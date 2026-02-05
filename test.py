import streamlit as st
import plotly.graph_objects as go
from reportlab.platypus import SimpleDocTemplate, Image, Table, Spacer, Paragraph
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import inch
from reportlab.lib.styles import getSampleStyleSheet
import os
import shutil

# -------------------------------------------------
# PAGE CONFIG
# -------------------------------------------------
st.set_page_config(page_title="Chart PDF Report", layout="wide")

# -------------------------------------------------
# CREATE SAMPLE CHARTS (REPLACE WITH YOUR VARIABLES)
# -------------------------------------------------
def create_charts():
    charts = []

    for i in range(10):  # 10 charts = 5 rows × 2 columns
        fig = go.Figure()
        fig.add_trace(
            go.Scatter(
                x=[1, 2, 3, 4, 5],
                y=[i, i+2, i+1, i+3, i+4],
                mode="lines+markers"
            )
        )
        fig.update_layout(
            title=f"Chart {i+1}",
            height=300,
            margin=dict(l=40, r=40, t=50, b=40)
        )
        charts.append(fig)

    return charts


# -------------------------------------------------
# SAVE CHARTS AS IMAGES
# -------------------------------------------------
def save_charts_as_images(charts, folder="temp_charts"):
    if os.path.exists(folder):
        shutil.rmtree(folder)

    os.makedirs(folder, exist_ok=True)
    image_paths = []

    for idx, fig in enumerate(charts):
        img_path = f"{folder}/chart_{idx+1}.png"
        fig.write_image(img_path, scale=2)  # requires kaleido
        image_paths.append(img_path)

    return image_paths


# -------------------------------------------------
# GENERATE PDF (2 CHARTS PER ROW, 5 ROWS)
# -------------------------------------------------
def generate_pdf(image_paths, pdf_path="Chart_Report.pdf"):
    styles = getSampleStyleSheet()
    pdf = SimpleDocTemplate(
        pdf_path,
        pagesize=A4,
        rightMargin=30,
        leftMargin=30,
        topMargin=30,
        bottomMargin=30
    )

    elements = []

    # -------- Title --------
    elements.append(Paragraph("<b>Energy Analysis Report</b>", styles["Title"]))
    elements.append(Spacer(1, 0.3 * inch))

    table_data = []
    row = []

    for i, img_path in enumerate(image_paths):
        img = Image(img_path, width=3.8 * inch, height=2.4 * inch)
        row.append(img)

        if len(row) == 2:
            table_data.append(row)
            row = []

        # 5 rows per page
        if len(table_data) == 5:
            elements.append(Table(table_data, colWidths=[4 * inch, 4 * inch]))
            elements.append(Spacer(1, 0.4 * inch))
            table_data = []

    if row:
        table_data.append(row)

    if table_data:
        elements.append(Table(table_data, colWidths=[4 * inch, 4 * inch]))
    pdf.build(elements)

# -------------------------------------------------
# STREAMLIT UI
# -------------------------------------------------
st.title("📊 Streamlit Charts → PDF Report")
charts = create_charts()
# -------- Display charts in Streamlit (2 per row) --------
for i in range(0, len(charts), 2):
    col1, col2 = st.columns(2)
    with col1:
        st.plotly_chart(charts[i], use_container_width=True)
    with col2:
        if i + 1 < len(charts):
            st.plotly_chart(charts[i + 1], use_container_width=True)
st.divider()
# -------------------------------------------------
# PDF GENERATION BUTTON
# -------------------------------------------------
if st.button("📄 Generate PDF Report"):
    with st.spinner("Generating PDF report..."):
        image_paths = save_charts_as_images(charts)
        generate_pdf(image_paths)

    with open("Chart_Report.pdf", "rb") as pdf_file:
        st.download_button(
            label="⬇️ Download PDF Report",
            data=pdf_file,
            file_name="Chart_Report.pdf",
            mime="application/pdf"
        )

    st.success("PDF generated successfully!")
