# -*- coding: utf-8 -*-
"""
Renta Neta MercadoLibre — ELDOM EL BAZAR / DELERAL S.A.
=======================================================
Herramienta de cierre mensual: subís los reportes de cada costo y la app
netea todo para darte la renta neta de la operación de MercadoLibre.

Canales/fuentes que procesa:
  1. Reporte de Facturación de MercadoLibre (.xlsx)  -> comisión, envíos ME2/Colecta, publicidad, etc.
  2. Reporte de Notas de Crédito de MercadoLibre (.xlsx)
  3. Reporte de Costeo FacturApp - FACTURAS (.pdf)   -> venta y costo del producto
  4. Reporte de Costeo FacturApp - NOTAS DE CRÉDITO / DEVOLUCIONES (.pdf)
  5. Facturación mensual DAC (.pdf)                  -> flete ME1 (cuenta corriente)
  6. Nota de Crédito DAC / bonificación (.pdf)       -> descuento 20%
  7. Factura FLEX Distrilogic (.pdf)                 -> logística FLEX

Deploy: Streamlit Community Cloud o Render. Solo requiere app.py + requirements.txt.
"""

import io
import re
import pandas as pd
import pdfplumber
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

IVA_RATE = 0.22

# ---------------------------------------------------------------------------
# Utilidades de parseo numérico
# ---------------------------------------------------------------------------
def parse_amount(s):
    """Convierte un texto de importe a float, tolerando formato UY (1.234,56)
    y americano (1234.56)."""
    s = str(s).replace("$", "").replace("UYU", "").replace("\u00a0", " ").strip()
    s = s.replace(" ", "")
    if not s:
        return 0.0
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):          # coma decimal (UY)
            s = s.replace(".", "").replace(",", ".")
        else:                                     # punto decimal, coma miles
            s = s.replace(",", "")
    elif "," in s:                                # solo coma -> decimal
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def _pdf_text(f):
    f.seek(0)
    out = []
    with pdfplumber.open(f) as pdf:
        for pg in pdf.pages:
            out.append(pg.extract_text() or "")
    return "\n".join(out)


# ---------------------------------------------------------------------------
# 1-2) Cargos de MercadoLibre (facturación + notas de crédito, .xlsx)
# ---------------------------------------------------------------------------
def parse_ml_charges(fac_file, nc_file):
    frames = []
    for f in (fac_file, nc_file):
        if f is None:
            continue
        f.seek(0)
        df = pd.read_excel(f, sheet_name="REPORT", header=7).dropna(how="all")
        frames.append(df)
    if not frames:
        return None
    allml = pd.concat(frames, ignore_index=True)
    d = allml["Detalle"].astype(str)
    val = pd.to_numeric(allml["Valor del cargo"], errors="coerce").fillna(0)

    def s(mask):
        return round(float(val[mask].sum()), 2)

    return dict(
        comision=s(d.str.contains("cargo por venta", case=False, na=False)),
        envios=s(d.str.contains("envíos de Mercado Libre", case=False, na=False)
                 | d.str.contains("costo de envío", case=False, na=False)),
        prodads=s(d.str.contains("Product Ads", case=False, na=False)),
        dispads=s(d.str.contains("Display Ads", case=False, na=False)),
        aseso=s(d.str.contains("Asesoría", case=False, na=False)),
        manten=s(d.str.contains("mantenimiento de Mi página", case=False, na=False)),
        devfee=s(d.str.contains("Cargo por devolución", case=False, na=False)),
    )


# ---------------------------------------------------------------------------
# 3-4) Costeo FacturApp (.pdf) -> (costo, venta) con signo del reporte
# ---------------------------------------------------------------------------
def parse_costeo_pdf(f):
    if f is None:
        return (0.0, 0.0)
    text = _pdf_text(f)
    # Renglón oficial de totales: "TOTAL: costo$ costoU$S venta$ ventaU$S"
    summary = None
    for ln in text.split("\n"):
        m = re.search(
            r"TOTAL:\s*(-?\d+\.\d{2})\s+(-?\d+\.\d{2})\s+(-?\d+\.\d{2})\s+(-?\d+\.\d{2})",
            ln.strip())
        if m:
            summary = m
    if summary:
        return (float(summary.group(1)), float(summary.group(3)))
    # Fallback: sumar líneas de detalle
    costo = venta = 0.0
    pat = re.compile(
        r"(FACTURA CONTADO|DEVOLUCION CONTAD)\s+\S+.*?\$?\s*[\d.]+\s+\$?\s*"
        r"(-?\d+\.\d+)\s+(-?\d+\.\d+)\s+(-?\d+\.\d+)\s+(-?\d+\.\d+)")
    for ln in text.split("\n"):
        m = pat.search(ln)
        if m:
            costo += float(m.group(2))
            venta += float(m.group(4))
    return (round(costo, 2), round(venta, 2))


# ---------------------------------------------------------------------------
# 5) Facturación DAC (.pdf) -> total, salida (ventas ME1), entrada (proveedores)
# ---------------------------------------------------------------------------
def parse_dac_fact(f):
    if f is None:
        return dict(total=0.0, salida=0.0, entrada=0.0)
    total = sal = ent = 0.0
    f.seek(0)
    with pdfplumber.open(f) as pdf:
        for pg in pdf.pages:
            for ln in (pg.extract_text() or "").split("\n"):
                s = ln.strip()
                is_sal = "Cta.Cte. Remitente" in s
                is_ent = "Flete destino" in s
                if not (is_sal or is_ent):
                    continue
                nums = re.findall(r"\d[\d.]*,\d\d", s)
                if not nums:
                    continue
                v = parse_amount(nums[-1])
                total += v
                if is_sal:
                    sal += v
                else:
                    ent += v
    return dict(total=round(total, 2), salida=round(sal, 2), entrada=round(ent, 2))


# ---------------------------------------------------------------------------
# 6) Nota de crédito DAC (bonificación) -> total (con IVA), subtotal (sin IVA)
# ---------------------------------------------------------------------------
def parse_dac_nc(f):
    if f is None:
        return dict(total=0.0, subtotal=0.0)
    text = _pdf_text(f)
    total = sub = 0.0
    m = re.search(r"TOTAL:\s*\$?\s*([\d.,]+)", text)
    if m:
        total = parse_amount(m.group(1))
    m = re.search(r"Subtotal:\s*\$?\s*([\d.,]+)", text)
    if m:
        sub = parse_amount(m.group(1))
    return dict(total=round(total, 2), subtotal=round(sub, 2))


# ---------------------------------------------------------------------------
# 7) Factura FLEX Distrilogic -> total (con IVA), neto (sin IVA)
# ---------------------------------------------------------------------------
def parse_flex(f):
    if f is None:
        return dict(total=0.0, neto=0.0)
    text = _pdf_text(f)
    total = neto = 0.0
    m = re.search(r"TOTAL A PAGAR\s*\n?\s*(?:UYU)?\s*([\d.,]+)", text)
    if not m:
        m = re.search(r"TOTAL A PAGAR[^\d]*([\d.,]+)", text)
    if m:
        total = parse_amount(m.group(1))
    m2 = re.search(r"([\d.]+,\d\d)\s+([\d.]+,\d\d)\s+([\d.]+,\d\d)", text)
    if m2:
        neto = parse_amount(m2.group(1))
    if not neto and total:
        neto = round(total / (1 + IVA_RATE), 2)
    return dict(total=round(total, 2), neto=round(neto, 2))


# ---------------------------------------------------------------------------
# Cálculo del estado de resultados
# ---------------------------------------------------------------------------
def compute(ml, cost_fac, cost_nc, dac, dacnc, flex, impuestos=0.0):
    venta_bruta, venta_dev = cost_fac[1], cost_nc[1]
    costo_bruto, costo_dev = cost_fac[0], cost_nc[0]
    venta_neta = round(venta_bruta + venta_dev, 2)
    costo_neto = round(costo_bruto + costo_dev, 2)
    gb = round(venta_neta - costo_neto, 2)

    ml = ml or dict(comision=0, envios=0, prodads=0, dispads=0, aseso=0, manten=0, devfee=0)
    cargos_ml = round(ml["comision"] + ml["prodads"] + ml["dispads"]
                      + ml["aseso"] + ml["manten"] + ml["devfee"], 2)

    dac_neto = round(dac["total"] - dacnc["total"], 2)          # con IVA
    envios = round(ml["envios"] + dac_neto + flex["total"], 2)

    rno = round(gb - cargos_ml - envios, 2)
    final = round(rno - impuestos, 2)
    margen = (rno / venta_neta) if venta_neta else 0.0

    return dict(
        venta_bruta=venta_bruta, venta_dev=venta_dev, venta_neta=venta_neta,
        costo_bruto=costo_bruto, costo_dev=costo_dev, costo_neto=costo_neto,
        gb=gb, cargos_ml=cargos_ml, dac_neto=dac_neto, flex=flex["total"],
        me2=ml["envios"], envios=envios, rno=rno, final=final, margen=margen, ml=ml,
        dac=dac, dacnc=dacnc,
    )


# ---------------------------------------------------------------------------
# Exportación a Excel con formato
# ---------------------------------------------------------------------------
def build_excel(r):
    MONEY = "#,##0.00;(#,##0.00)"; PCT = "0.0%"
    BLUE = Font(name="Arial", color="0000FF", size=10)
    BLACK = Font(name="Arial", color="000000", size=10)
    BOLD = Font(name="Arial", bold=True, size=10)
    BOLDW = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    TITLE = Font(name="Arial", bold=True, size=13, color="7B1E22")
    CRIMSON = PatternFill("solid", fgColor="7B1E22")
    GREEN = PatternFill("solid", fgColor="E2EFDA")
    top = Border(top=Side(style="thin", color="000000"))
    dbl = Border(top=Side(style="thin", color="000000"),
                 bottom=Side(style="double", color="000000"))

    wb = Workbook(); ws = wb.active; ws.title = "Resultado"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 52
    ws.column_dimensions["C"].width = 20
    ws["B2"] = "ELDOM EL BAZAR — DELERAL S.A."; ws["B2"].font = TITLE
    ws["B3"] = "Renta neta operación MercadoLibre"; ws["B3"].font = BOLD

    row = [5]
    def line(label, value, font=BLACK, fmt=MONEY, fill=None, bd=None):
        rr = row[0]
        c1 = ws.cell(rr, 2, label); c1.font = font
        c2 = ws.cell(rr, 3, value); c2.font = font; c2.number_format = fmt
        c2.alignment = Alignment(horizontal="right")
        if fill:
            c1.fill = fill; c2.fill = fill
        if bd:
            c2.border = bd
        row[0] += 1
    def blank():
        row[0] += 1
    def hdr(t):
        rr = row[0]
        ws.cell(rr, 2, t).font = BOLDW
        ws.cell(rr, 2).fill = CRIMSON; ws.cell(rr, 3).fill = CRIMSON
        row[0] += 1

    hdr("INGRESOS")
    line("Ventas facturadas (FacturApp)", r["venta_bruta"], BLUE)
    line("(–) Devoluciones / Notas de crédito", r["venta_dev"], BLUE)
    line("Ventas netas", r["venta_neta"], BOLD, bd=top); blank()
    hdr("COSTO DE MERCADERÍA")
    line("Costo de productos vendidos", -abs(r["costo_bruto"]), BLUE)
    line("(+) Reversión por devoluciones", abs(r["costo_dev"]), BLUE)
    line("Costo neto", -abs(r["costo_neto"]), BOLD, bd=top); blank()
    line("GANANCIA BRUTA", r["gb"], BOLD, MONEY, GREEN)
    line("Margen bruto", r["gb"] / r["venta_neta"] if r["venta_neta"] else 0, BLACK, PCT); blank()
    hdr("CARGOS DE MERCADOLIBRE")
    line("Comisión por venta (neta)", -r["ml"]["comision"], BLUE)
    line("Publicidad — Product Ads", -r["ml"]["prodads"], BLUE)
    line("Publicidad — Display Ads", -r["ml"]["dispads"], BLUE)
    line("Asesoría Comercial", -r["ml"]["aseso"], BLUE)
    line("Mantenimiento Mi página (neto)", -r["ml"]["manten"], BLUE)
    line("Cargo por devolución (fee ML)", -r["ml"]["devfee"], BLUE)
    line("Subtotal cargos ML", -r["cargos_ml"], BOLD, bd=top); blank()
    hdr("ENVÍOS (3 canales)")
    line("ME2 / Colecta (MercadoLibre)", -r["me2"], BLUE)
    line("ME1 / DAC (neto cta.cte., con IVA)", -r["dac_neto"], BLUE)
    line("FLEX / Distrilogic (con IVA)", -r["flex"], BLUE)
    line("Subtotal envíos", -r["envios"], BOLD, bd=top); blank()
    line("RENTA NETA OPERATIVA (antes de imp.)", r["rno"], BOLDW, MONEY, CRIMSON)
    line("Margen neto operativo", r["margen"], BLACK, PCT); blank()
    hdr("IMPUESTOS")
    line("Impuestos (IVA neto DGI / IRAE)", -(r["final"] and (r["rno"] - r["final"]) or 0), BLUE)
    line("RENTA NETA FINAL", r["final"], BOLDW, MONEY, CRIMSON, bd=dbl)

    bio = io.BytesIO(); wb.save(bio); bio.seek(0)
    return bio.getvalue()


# ---------------------------------------------------------------------------
# Interfaz
# ---------------------------------------------------------------------------
st.set_page_config(page_title="Renta Neta ML — ELDOM", page_icon="📊", layout="wide")
st.title("📊 Renta Neta MercadoLibre")
st.caption("ELDOM EL BAZAR / DELERAL S.A. — cierre mensual. Subí cada reporte y presioná Procesar.")

with st.expander("¿Qué archivo va en cada casillero?"):
    st.markdown("""
- **Facturación ML** y **Notas de Crédito ML** (`.xlsx`): reportes de facturación de MercadoLibre. De ahí salen comisión, envíos ME2/Colecta, publicidad, asesoría, etc.
- **Costeo FacturApp – Facturas** y **– Notas de Crédito** (`.pdf`): reporte de costeo de FacturApp (medio de pago M LIBRE). De ahí salen la venta y el costo del producto.
- **DAC Facturación** y **DAC Nota de Crédito** (`.pdf`): flete ME1 por cuenta corriente y su bonificación del 20%.
- **FLEX Distrilogic** (`.pdf`): factura de logística FLEX.

Podés subir solo los que tengas; lo que falte queda en cero.
""")

c1, c2 = st.columns(2)
with c1:
    st.subheader("MercadoLibre")
    ml_fac = st.file_uploader("Facturación ML (.xlsx)", type=["xlsx"], key="mlf")
    ml_nc = st.file_uploader("Notas de Crédito ML (.xlsx)", type=["xlsx"], key="mln")
    st.subheader("Costeo FacturApp")
    co_fac = st.file_uploader("Costeo – Facturas (.pdf)", type=["pdf"], key="cof")
    co_nc = st.file_uploader("Costeo – Notas de Crédito (.pdf)", type=["pdf"], key="con")
with c2:
    st.subheader("Flete DAC (ME1)")
    dac_f = st.file_uploader("DAC – Facturación mensual (.pdf)", type=["pdf"], key="dacf")
    dac_n = st.file_uploader("DAC – Nota de Crédito / bonificación (.pdf)", type=["pdf"], key="dacn")
    st.subheader("Logística FLEX")
    flex_f = st.file_uploader("Distrilogic – Factura (.pdf)", type=["pdf"], key="flexf")

st.divider()
impuestos = st.number_input(
    "Impuestos a descontar (IVA neto DGI / IRAE) — opcional",
    min_value=0.0, value=0.0, step=1000.0, format="%.2f",
    help="Dejalo en 0 para ver la renta neta operativa. Cargá el IVA neto / IRAE del mes para la renta neta final.")

if st.button("⚙️  Procesar", type="primary", use_container_width=True):
    try:
        ml = parse_ml_charges(ml_fac, ml_nc)
        cost_fac = parse_costeo_pdf(co_fac)
        cost_nc = parse_costeo_pdf(co_nc)
        dac = parse_dac_fact(dac_f)
        dacnc = parse_dac_nc(dac_n)
        flex = parse_flex(flex_f)
        r = compute(ml, cost_fac, cost_nc, dac, dacnc, flex, impuestos)

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Ventas netas", f"$ {r['venta_neta']:,.0f}")
        m2.metric("Ganancia bruta", f"$ {r['gb']:,.0f}")
        m3.metric("Renta neta operativa", f"$ {r['rno']:,.0f}", f"{r['margen']*100:.1f}% s/ventas")
        m4.metric("Renta neta FINAL", f"$ {r['final']:,.0f}")

        tabla = pd.DataFrame([
            ("Ventas netas", r["venta_neta"]),
            ("Costo neto de mercadería", -abs(r["costo_neto"])),
            ("= Ganancia bruta", r["gb"]),
            ("Comisión ML (neta)", -r["ml"]["comision"]),
            ("Publicidad Product Ads", -r["ml"]["prodads"]),
            ("Publicidad Display Ads", -r["ml"]["dispads"]),
            ("Asesoría Comercial", -r["ml"]["aseso"]),
            ("Mantenimiento Mi página", -r["ml"]["manten"]),
            ("Cargo por devolución", -r["ml"]["devfee"]),
            ("Envío ME2 / Colecta (ML)", -r["me2"]),
            ("Envío ME1 / DAC (neto)", -r["dac_neto"]),
            ("Envío FLEX / Distrilogic", -r["flex"]),
            ("= Renta neta operativa", r["rno"]),
            ("(–) Impuestos", -(r["rno"] - r["final"])),
            ("= RENTA NETA FINAL", r["final"]),
        ], columns=["Concepto", "UYU"])
        st.dataframe(
            tabla.style.format({"UYU": "{:,.2f}"}),
            use_container_width=True, hide_index=True, height=560)

        st.download_button(
            "⬇️  Descargar Excel", data=build_excel(r),
            file_name="Renta_Neta_ML.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)

        with st.expander("Detalle de envíos y controles"):
            st.write(f"**ME2/Colecta (ML):** ${r['me2']:,.2f} (columna Valor del cargo)")
            st.write(f"**DAC:** facturación ${r['dac']['total']:,.2f} − bonificación "
                     f"${r['dacnc']['total']:,.2f} = neto ${r['dac_neto']:,.2f}  ·  "
                     f"salida ventas ${r['dac']['salida']:,.2f} / proveedores ${r['dac']['entrada']:,.2f}")
            st.write(f"**FLEX Distrilogic:** ${r['flex']:,.2f} con IVA "
                     f"(neto ${flex['neto']:,.2f})")
    except Exception as e:
        st.error(f"Error procesando los archivos: {e}")
        st.info("Verificá que cada archivo esté en el casillero correcto y con el formato esperado.")
