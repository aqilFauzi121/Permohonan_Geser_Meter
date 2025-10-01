import streamlit as st
import pandas as pd
from auth import get_gspread_client
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
import io

# Baca data pelanggan dari Google Sheets
def load_pelanggan():
    gc = get_gspread_client()
    sh = gc.open("DataPelanggan")   # ganti sesuai nama sheet
    ws = sh.worksheet("Sheet1")
    data = ws.get_all_records()
    return pd.DataFrame(data)

# Data barang (contoh, bisa juga dari sheet lain)
barang_list = {
    "Service wedge clamp 2/4 x 6/10 mm": 5526,
    "Strainhook / eker babi": 20000,
    "Cable support (50/80J/2009)": 1643,
    "Conn. press. AL/AL type 10-16 mm² / 50-70 mm² + Scoot + Cover": 30698,
    "Paku Beton": 500,
    "Pole Bracket 3-9": 18730,
    "Segel Plastik": 1000,
    "Twisted Cable 2 x 10 mm² – Al": 43714,
}

# Fungsi export PDF
def create_pdf(pelanggan, df_rekap, total, ppn, grand_total):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    styles = getSampleStyleSheet()
    elements = []

    # Judul
    elements.append(Paragraph("<b>REKAP HARGA PEKERJAAN</b>", styles['Title']))
    elements.append(Spacer(1, 12))

    # Data pelanggan
    elements.append(Paragraph(f"Nama: {pelanggan['Nama']}", styles['Normal']))
    elements.append(Paragraph(f"Alamat: {pelanggan['Alamat']}", styles['Normal']))
    elements.append(Spacer(1, 12))

    # Tabel barang
    table_data = [["No", "Barang", "Qty", "Harga Satuan", "Subtotal"]]
    for i, row in df_rekap.iterrows():
        table_data.append([
            i + 1,
            row["Barang"],
            row["Qty"],
            f"Rp {row['Harga Satuan']:,}",
            f"Rp {row['Subtotal']:,}",
        ])
    table = Table(table_data, colWidths=[1.5*cm, 7*cm, 2*cm, 3*cm, 3*cm])
    table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.grey),
        ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0,0), (-1,0), 8),
        ('GRID', (0,0), (-1,-1), 1, colors.black),
    ]))
    elements.append(table)
    elements.append(Spacer(1, 12))

    # Total
    elements.append(Paragraph(f"<b>TOTAL MATERIAL: Rp {total:,}</b>", styles['Normal']))
    elements.append(Paragraph(f"PPN (11%): Rp {ppn:,}", styles['Normal']))
    elements.append(Paragraph(f"<b>TOTAL BIAYA SETELAH PPN: Rp {grand_total:,}</b>", styles['Heading2']))

    doc.build(elements)
    pdf = buffer.getvalue()
    buffer.close()
    return pdf

# Streamlit App
st.title("Rekap Harga Pekerjaan")

df_pelanggan = load_pelanggan()
idpel = st.selectbox("Pilih ID Pelanggan:", df_pelanggan["IDPEL"].unique())
pelanggan = df_pelanggan[df_pelanggan["IDPEL"] == idpel].iloc[0]

st.write("**Nama:**", pelanggan["Nama"])
st.write("**Alamat:**", pelanggan["Alamat"])
st.divider()

st.subheader("Input Barang yang Digunakan")
kuantitas = {}
for barang, harga in barang_list.items():
    qty = st.number_input(f"{barang} (Harga Rp {harga:,})", min_value=0, step=1)
    if qty > 0:
        kuantitas[barang] = {"qty": qty, "harga": harga, "subtotal": qty * harga}

if kuantitas:
    df_rekap = pd.DataFrame([
        {"Barang": b, "Qty": v["qty"], "Harga Satuan": v["harga"], "Subtotal": v["subtotal"]}
        for b, v in kuantitas.items()
    ])

    total = df_rekap["Subtotal"].sum()
    ppn = int(total * 0.11)
    grand_total = total + ppn

    st.table(df_rekap)
    st.write("**TOTAL MATERIAL:** Rp", f"{total:,}")
    st.write("**PPN (11%):** Rp", f"{ppn:,}")
    st.write("### TOTAL BIAYA SETELAH PPN: Rp", f"{grand_total:,}")

    # Tombol download PDF
    pdf = create_pdf(pelanggan, df_rekap, total, ppn, grand_total)
    st.download_button("📥 Download Rekap PDF", data=pdf, file_name="rekap_harga.pdf", mime="application/pdf")
