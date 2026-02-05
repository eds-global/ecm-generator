import os, sys, time, webbrowser
import pandas as pd
import streamlit as st
import plotly.express as px
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import json
import streamlit.components.v1 as components
import pdfplumber
import re
import io
import matplotlib.image as mpimg
from streamlit.runtime.uploaded_file_manager import UploadedFile
from io import BytesIO
import time
import random
from datetime import datetime
from shutil import copyfile
import glob as gb
import subprocess
import shutil
from pathlib import Path
from collections import defaultdict
import streamlit.web.cli as stcli
from src import insertWall, insertConst, orient, lighting, equip, windows, insertRoof, wwr
import traceback
from helper import *
from report_ext import *
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
import statsmodels.api as sm 
from src import ModifyWallRoof
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
import plotly.io as pio
from src import lv_b, ls_c, lv_d, pv_a_loop, sv_a, beps, bepu, lvd_summary, sva_zone, locationInfo, masterFile, sva_sys_type, pv_a_pump, pv_a_heater, pv_a_equip, pv_a_tower, ps_e, inp_shgc
from reportlab.platypus import SimpleDocTemplate, Image, Table, Spacer, Paragraph
from reportlab.platypus import BaseDocTemplate, Frame, PageTemplate
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import inch
from reportlab.lib.styles import getSampleStyleSheet
import plotly.io as pio
from reportlab.lib.units import inch
# pio.kaleido.scope.default_format = "png"
# pio.kaleido.scope.default_width = 1200
# pio.kaleido.scope.default_height = 800


# --- Streamlit Page Config ---
st.set_page_config(page_title="AWESIM", page_icon="⚡", layout='wide')

# --- Inject Custom CSS ---
st.markdown("""
    <style>
        .block-container { padding-top: 0rem !important; }
        header, main { margin-top: 0rem !important; padding-top: 0rem !important; }
        .stButton>button { box-shadow: 1px 1px 1px rgba(0, 0, 0, 0.8); }
        .heading-with-shadow {
            text-align: left;
            color: red;
            text-shadow: 0px 8px 4px rgba(255, 255, 255, 0.4);
            background-color: white;
        }
        body {
            background-color: #bfe1ff;
            animation: changeColor 5s infinite;
        }
        #MainMenu, footer, .viewerBadge_container__1QSob {visibility: hidden;}
        header .stApp [title="View source on GitHub"] { display: none; }
        .stApp header, .stApp footer {visibility: hidden;}
        .stButton button {
            height: 30px;
            width: 166px;
        }
    </style>
""", unsafe_allow_html=True)

def resource_path(relative_path):
    """Get absolute path to resource (works for dev and PyInstaller exe)"""
    if getattr(sys, 'frozen', False):  # Running in exe
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, relative_path)

def remove_utility(inp_data):
    start_marker = "Utility Rates"
    end_marker = "Output Reporting"

    start_index, end_index = None, None

    # Find first occurrence of start_marker
    for i, line in enumerate(inp_data):
        if start_marker in line:
            start_index = i + 3

    # Find first occurrence of end_marker after start_marker
    if start_index is not None:
        for i, line in enumerate(inp_data[start_index:], start=start_index):
            if end_marker in line:
                end_index = i - 3
                break

    if start_index is None or end_index is None:
        raise ValueError("Could not find section markers in INP file.")

    # Remove everything between start_index and end_index (inclusive)
    new_data = inp_data[:start_index] + inp_data[end_index + 1:]
    return new_data

def remove_betweenLightEquip(inp_data):
    start_marker = "Floors / Spaces / Walls / Windows / Doors"
    end_marker = "Electric & Fuel Meters"

    start_index, end_index = None, None

    # Find first occurrence of start_marker
    for i, line in enumerate(inp_data):
        if start_marker in line:
            start_index = i + 2
            break

    # Find first occurrence of end_marker after start_marker
    if start_index is not None:
        for i, line in enumerate(inp_data[start_index:], start=start_index):
            if end_marker in line:
                end_index = i - 2
                break

    if start_index is None or end_index is None:
        raise ValueError("Could not find section markers in INP file.")

    new_data = []
    cleaning = False  # to track inside LIGHTING/EQUIPMENT block

    for i, line in enumerate(inp_data):
        if start_index <= i <= end_index:
            if line.strip().startswith("LIGHTING-W/AREA") or line.strip().startswith("EQUIPMENT-W/AREA"):
                cleaning = True
                new_data.append(line)  # keep the main line
                continue
            if cleaning:
                if line.strip().startswith("*"):
                    continue  # skip continuation lines
                else:
                    cleaning = False  # stop cleaning when normal line comes
        new_data.append(line)
    return new_data

def energy_param_plot_wall(
    x_param_df,
    baseline_df,
    x_col,
    x_label,
    title,
    show_legend=True):

    df = x_param_df.copy()

    # =========================
    # Create DensityType safely
    # =========================
    if "DensityType" not in df.columns:
        df["DensityIndex"] = (
            df["FileName"]
            .astype(str)
            .str.split("_")
            .str[1]
            .astype(int)
        )

        df["DensityType"] = df["DensityIndex"].apply(
            lambda x: "Mid Density" if x >= 13 else "Low Density"
        )

    fig = go.Figure()

    baseline_x = baseline_df[x_col].iloc[0]

    # =========================
    # Low Density curve
    # =========================
    mid_df = (
        df[(df["DensityType"] == "Low Density") & (df[x_col] != baseline_x)]
        .sort_values(by=x_col)
    )

    fig.add_trace(
        go.Scatter(
            x=mid_df[x_col],
            y=mid_df["Energy_Outcome(KWH)"],
            mode="lines",
            line=dict(
                color="#B0B0B0",
                width=2,
                shape="spline",
                smoothing=1.2
            ),
            name="Low Density",
            hovertemplate=(
                f"<b>{x_label}</b>: %{{x:.2f}}<br>"
                "<b>Energy</b>: %{y:,.0f} kWh<br>"
                "<b>Type</b>: Low Density<extra></extra>"
            )
        )
    )

    # =========================
    # Mid Density curve
    # =========================
    high_df = (
        df[(df["DensityType"] == "High Density") & (df[x_col] != baseline_x)]
        .sort_values(by=x_col)
    )

    fig.add_trace(
        go.Scatter(
            x=high_df[x_col],
            y=high_df["Energy_Outcome(KWH)"],
            mode="lines",
            line=dict(
                color="#6B6B6B",
                width=2,
                shape="spline",
                smoothing=1.2
            ),
            name="High Density",
            hovertemplate=(
                f"<b>{x_label}</b>: %{{x:.2f}}<br>"
                "<b>Energy</b>: %{y:,.0f} kWh<br>"
                "<b>Type</b>: High Density<extra></extra>"
            )
        )
    )

    # =========================
    # As Designed (only point kept)
    # =========================
    fig.add_trace(
        go.Scatter(
            x=baseline_df[x_col],
            y=baseline_df["Energy_Outcome(KWH)"],
            mode="markers",
            marker=dict(
                size=14,
                color="red",
                symbol="circle",
                line=dict(width=2, color="red")
            ),
            name="As Designed",
            hovertemplate=(
                f"<b>{x_label}</b>: %{{x:.2f}}<br>"
                "<b>Energy</b>: %{y:,.0f} kWh<br>"
                "<b>Type</b>: As Designed<extra></extra>"
            )
        )
    )

    # =========================
    # Layout
    # =========================
    fig.update_layout(
        # title=title,
        template="plotly_white",
        height=420,
        margin=dict(l=50, r=30, t=70, b=90),
        font=dict(size=13),
        legend=dict(
            orientation="h",
            yanchor="top",
            y=-0.25,
            xanchor="center",
            x=0.5
        ),
        showlegend=show_legend
    )

    fig.update_xaxes(title=x_label)
    fig.update_yaxes(title="Energy Use (kWh)", tickformat=".4s")

    return fig

# ------------ Graph ----------- #
def energy_param_plot(
    x_param_df,
    baseline_df,
    x_col,
    x_label,
    title,
    show_legend=True):

    df = x_param_df.copy()

    # =========================
    # Create DensityType safely
    # =========================
    if "DensityType" not in df.columns:
        df["DensityIndex"] = (
            df["FileName"]
            .astype(str)
            .str.split("_")
            .str[1]
            .astype(int)
        )

        df["DensityType"] = df["DensityIndex"].apply(
            lambda x: "High Density" if x >= 13 else "Mid Density"
        )

    fig = go.Figure()

    baseline_x = baseline_df[x_col].iloc[0]

    # =========================
    # Mid Density curve
    # =========================
    mid_df = (
        df[(df["DensityType"] == "Low Density") & (df[x_col] != baseline_x)]
        .sort_values(by=x_col)
    )

    fig.add_trace(
        go.Scatter(
            x=mid_df[x_col],
            y=mid_df["Energy_Outcome(KWH)"],
            mode="lines",
            line=dict(
                color="#B0B0B0",
                width=2,
                shape="spline",
                smoothing=1.2
            ),
            name="Low Density",
            hovertemplate=(
                f"<b>{x_label}</b>: %{{x:.2f}}<br>"
                "<b>Energy</b>: %{y:,.0f} kWh<br>"
                "<b>Type</b>: Low Density<extra></extra>"
            )
        )
    )

    # =========================
    # High Density curve
    # =========================
    high_df = (
        df[(df["DensityType"] == "High Density") & (df[x_col] != baseline_x)]
        .sort_values(by=x_col)
    )

    fig.add_trace(
        go.Scatter(
            x=high_df[x_col],
            y=high_df["Energy_Outcome(KWH)"],
            mode="lines",
            line=dict(
                color="#6B6B6B",
                width=2,
                shape="spline",
                smoothing=1.2
            ),
            name="High Density",
            hovertemplate=(
                f"<b>{x_label}</b>: %{{x:.2f}}<br>"
                "<b>Energy</b>: %{y:,.0f} kWh<br>"
                "<b>Type</b>: High Density<extra></extra>"
            )
        )
    )

    # =========================
    # As Designed (only point kept)
    # =========================
    fig.add_trace(
        go.Scatter(
            x=baseline_df[x_col],
            y=baseline_df["Energy_Outcome(KWH)"],
            mode="markers",
            marker=dict(
                size=14,
                color="red",
                symbol="circle",
                line=dict(width=2, color="red")
            ),
            name="As Designed",
            hovertemplate=(
                f"<b>{x_label}</b>: %{{x:.2f}}<br>"
                "<b>Energy</b>: %{y:,.0f} kWh<br>"
                "<b>Type</b>: As Designed<extra></extra>"
            )
        )
    )

    # =========================
    # Layout
    # =========================
    fig.update_layout(
        # title=title,
        template="plotly_white",
        height=420,
        margin=dict(l=50, r=30, t=70, b=90),
        font=dict(size=13),
        legend=dict(
            orientation="h",
            yanchor="top",
            y=-0.25,
            xanchor="center",
            x=0.5
        ),
        showlegend=show_legend
    )

    fig.update_xaxes(title=x_label)
    fig.update_yaxes(title="Energy Use (kWh)", tickformat=".4s")

    return fig

def energy_param_plots(
    x_param_df,
    baseline_df,
    x_col,
    x_label,
    title,
    show_legend=True):

    df = x_param_df.copy()

    # =========================
    # Create DensityType safely
    # =========================
    if "DensityType" not in df.columns:
        df["DensityIndex"] = (
            df["FileName"]
            .astype(str)
            .str.split("_")
            .str[1]
            .astype(int)
        )

        df["DensityType"] = df["DensityIndex"].apply(
            lambda x: "High Density" if x >= 13 else "Mid Density"
        )

    fig = go.Figure()

    baseline_x = baseline_df[x_col].iloc[0]

    # =========================
    # Mid Density
    # =========================
    mid_df = (
        df[(df["DensityType"] == "Mid Density") & (df[x_col] != baseline_x)]
        .sort_values(by=x_col)
    )

    # markers
    fig.add_trace(
        go.Scatter(
            x=mid_df[x_col],
            y=mid_df["Energy_Outcome(KWH)"],
            mode="markers",
            marker=dict(size=12, color="lightblue", opacity=0.9),
            customdata=mid_df["Energy_Saving_%"],
            hovertemplate=(
                f"<b>{x_label}</b>: %{{x:.2f}}<br>"
                "<b>Energy</b>: %{y:,.0f} kWh<br>"
                "<b>Saving</b>: %{customdata:.1f}%<br>"
                "<b>Type</b>: Mid Density<extra></extra>"
            ),
            name="Mid Density"
        )
    )

    # line
    fig.add_trace(
        go.Scatter(
            x=mid_df[x_col],
            y=mid_df["Energy_Outcome(KWH)"],
            mode="lines",
            line=dict(color="lightblue", width=1.5, dash="dot"),
            hoverinfo="skip",
            showlegend=False
        )
    )

    # =========================
    # High Density
    # =========================
    high_df = (
        df[(df["DensityType"] == "High Density") & (df[x_col] != baseline_x)]
        .sort_values(by=x_col)
    )

    # markers
    fig.add_trace(
        go.Scatter(
            x=high_df[x_col],
            y=high_df["Energy_Outcome(KWH)"],
            mode="markers",
            marker=dict(size=12, color="blue", opacity=0.9),
            customdata=high_df["Energy_Saving_%"],
            hovertemplate=(
                f"<b>{x_label}</b>: %{{x:.2f}}<br>"
                "<b>Energy</b>: %{y:,.0f} kWh<br>"
                "<b>Saving</b>: %{customdata:.1f}%<br>"
                "<b>Type</b>: High Density<extra></extra>"
            ),
            name="High Density"
        )
    )

    # line
    fig.add_trace(
        go.Scatter(
            x=high_df[x_col],
            y=high_df["Energy_Outcome(KWH)"],
            mode="lines",
            line=dict(color="blue", width=1.5, dash="dot"),
            hoverinfo="skip",
            showlegend=False
        )
    )

    # =========================
    # Baseline (As Designed)
    # =========================
    fig.add_trace(
        go.Scatter(
            x=baseline_df[x_col],
            y=baseline_df["Energy_Outcome(KWH)"],
            mode="markers",
            marker=dict(size=14, color="red", line=dict(width=2)),
            hovertemplate=(
                f"<b>{x_label}</b>: %{{x:.2f}}<br>"
                "<b>Energy</b>: %{y:,.0f} kWh<br>"
                "<b>Saving</b>: 0.0%<extra></extra>"
            ),
            name="As Designed"
        )
    )

    # =========================
    # Layout
    # =========================
    fig.update_layout(
        # title=title,
        template="plotly_white",
        height=420,
        margin=dict(l=50, r=30, t=70, b=90),
        font=dict(size=13),
        legend=dict(
            orientation="h",
            yanchor="top",
            y=-0.25,
            xanchor="center",
            x=0.5
        ),
        showlegend=show_legend
    )

    fig.update_xaxes(title=x_label)
    fig.update_yaxes(title="Energy Use (kWh)", tickformat=".4s")

    return fig


def make_single_plot(df, x_col, title, x_label):
    df = df.copy()

    df["CaseType"] = df.index.map(lambda i: "As Designed" if i == 0 else "Parameter")

    def human_readable(num):
        if num >= 1_000_000_000:
            return f"{num/1_000_000_000:.1f}B"
        elif num >= 1_000_000:
            return f"{num/1_000_000:.1f}M"
        elif num >= 1_000:
            return f"{num/1_000:.1f}k"
        else:
            return str(num)

    df["Energy_human"] = df["Energy_Outcome(KWH)"].apply(human_readable)

    fig = go.Figure()

    # Parameter
    ecm_df = df[df["CaseType"] == "Parameter"]
    fig.add_trace(go.Scatter(
        x=ecm_df[x_col],
        y=ecm_df["Energy_Outcome(KWH)"],
        mode="lines",   # ⬅️ line instead of only points
        name="Parameter",
        line=dict(
            color="#6B6B6B",    # ⬅️ nicer blue (not grey)
            width=2
        ),
        marker=dict(
            size=8,
            color="#6B6B6B"
        ),
        customdata=ecm_df[["Energy_human"]].values,
        hovertemplate=(
            f"<b>{x_label}</b>: %{{x:.1f}}<br>"
            "<b>Energy</b>: %{customdata[0]}<extra></extra>"
        )
    ))

    # fig.add_trace(go.Scatter(
    #     x=ecm_df[x_col],
    #     y=ecm_df["Energy_Outcome(KWH)"],
    #     mode="markers",
    #     name="Parameter",
    #     marker=dict(size=12, color="#6B6B6B", line=dict(width=1, color="#6B6B6B")),
    #     customdata=ecm_df[["Energy_human"]].values,
    #     hovertemplate=(
    #         f"<b>{x_label}</b>: %{{x:.1f}}<br>"
    #         "<b>Energy</b>: %{customdata[0]}<extra></extra>"
    #     )
    # ))

    # As Designed
    ad_df = df[df["CaseType"] == "As Designed"]
    fig.add_trace(go.Scatter(
        x=ad_df[x_col],
        y=ad_df["Energy_Outcome(KWH)"],
        mode="markers",
        name="As Designed",
        marker=dict(size=12, color="red", line=dict(width=2, color="red")),
        customdata=ad_df[["Energy_human"]].values,
        hovertemplate=(
            f"<b>{x_label}</b>: %{{x:.1f}}<br>"
            "<b>Energy</b>: %{customdata[0]}<extra></extra>"
        )
    ))

    fig.update_layout(
        # title=title,
        xaxis_title=x_label,
        yaxis_title="Energy Use (kWh)",
        yaxis=dict(tickformat=".1s"),
        legend=dict(orientation="h", y=-0.3, x=0.5, xanchor="center"),
        height=450
    )
    fig.update_yaxes(title="Energy Use (kWh)", tickformat=".4s")
    return fig

def card(col, title, value):
    col.markdown(f"""
    <div class="metric-card">
        <div class="metric-title">{title}</div>
        <div class="metric-value">{value}</div>
    </div>
    """, unsafe_allow_html=True)

def impact(col, label, value):
    col.markdown(f"""
    <div class="impact-card">
        <div class="metric-title">{label}</div>
        <div class="impact-value">{value if value is not None else "—"}%</div>
    </div>
    """, unsafe_allow_html=True)

def save_figs_as_images(figs, folder="pdf_charts"):
    if os.path.exists(folder):
        shutil.rmtree(folder)
    os.makedirs(folder)

    paths = []
    for i, fig in enumerate(figs):
        path = f"{folder}/chart_{i+1}.png"
        fig.write_image(path, scale=2)
        paths.append(path)

    return paths

def add_header_footer(canvas, doc, title, right_logo_path, left_logo_path):

    canvas.saveState()
    width, height = A4

    # ---------- HEADER SETTINGS ----------
    y_logo = height - 55

    # Logo sizes
    left_logo_w, left_logo_h = 1.1 * inch, 0.5 * inch
    second_left_logo_w, second_left_logo_h = 0.4 * inch, 0.4 * inch
    right_logo_w, right_logo_h = 0.7 * inch, 0.5 * inch

    x_start = 30
    gap = 10  # space between left logos

    # ---------- LEFT LOGO 1 ----------
    if left_logo_path and os.path.exists(left_logo_path):
        canvas.drawImage(
            left_logo_path,
            x_start,
            y_logo,
            width=left_logo_w,
            height=left_logo_h,
            preserveAspectRatio=True,
            mask="auto"
        )

    # # ---------- LEFT LOGO 2 ----------
    # if second_left_logo_path and os.path.exists(second_left_logo_path):
    #     canvas.drawImage(
    #         second_left_logo_path,
    #         x_start + left_logo_w + gap,
    #         y_logo + (left_logo_h - second_left_logo_h) / 2,
    #         width=second_left_logo_w,
    #         height=second_left_logo_h,
    #         preserveAspectRatio=True,
    #         mask="auto"
    #     )

    # ---------- RIGHT LOGO ----------
    if right_logo_path and os.path.exists(right_logo_path):
        canvas.drawImage(
            right_logo_path,
            width - 30 - right_logo_w,
            y_logo + (left_logo_h - right_logo_h) / 2,
            width=right_logo_w,
            height=right_logo_h,
            preserveAspectRatio=True,
            mask="auto"
        )

    # ---------- TITLE ----------
    canvas.setFont("Helvetica-Bold", 14)
    canvas.setFillColor(colors.black)
    canvas.drawCentredString(width / 2, height - 35, title)

    # ---------- DIVIDER ----------
    canvas.setStrokeColor(colors.grey)
    canvas.setLineWidth(0.8)
    canvas.line(30, height - 65, width - 30, height - 65)

    # ---------- FOOTER ----------
    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(colors.grey)

    canvas.drawString(30, 25, "EDS | Automating Workflows for Energy Simulation")
    canvas.drawRightString(width - 30, 25, f"Page {doc.page}")

    canvas.restoreState()

from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
def add_later_pages(canvas, doc):
    canvas.saveState()
    # Move coordinate system UP → looks like smaller top margin
    add_footer(canvas, doc)
    canvas.restoreState()

def add_footer(canvas, doc):
    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(colors.grey)

    canvas.drawString(
        30,
        25,
        "EDS | Automating Workflows for Energy Simulation"
    )

    canvas.drawRightString(
        A4[0] - 30,
        25,
        f"Page {doc.page}"
    )

from reportlab.platypus import PageBreak
from reportlab.platypus import ListFlowable, ListItem

def create_pdf(image_paths, project_info, values, pdf_name="Energy_Parametric_Report.pdf"):
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image)
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch
    from reportlab.lib import colors

    styles = getSampleStyleSheet()

    # -------------------------------
    # KPI paragraph styles (MATCH INFO TABLE)
    # -------------------------------
    kpi_normal = ParagraphStyle(
        name="KpiNormal",
        fontName="Helvetica",
        fontSize=9,
        leading=11
    )

    kpi_bold = ParagraphStyle(
        name="KpiBold",
        fontName="Helvetica-Bold",
        fontSize=9,
        leading=11
    )

    pdf = SimpleDocTemplate(
        pdf_name,
        pagesize=A4,
        leftMargin=30,
        rightMargin=30,
        topMargin=70,
        bottomMargin=45
    )

    elements = []
    # -------------------------------
    # HEADER INFO
    # -------------------------------
    title = "Automating Workflows for Energy Simulation"
    logo_path = "images/EDSlogo.jpg"
    awesim_logo = "images/awesim.png"

    # -------------------------------
    # PROJECT INFORMATION TABLE (UNCHANGED)
    # -------------------------------
    elements.append(
        Paragraph(
            "<b>Project Information</b>",
            ParagraphStyle(
                name="KpiTitle",
                fontSize=10,
                spaceAfter=8
            )
        )
    )

    info_data = [
        ["Project Name", project_info.get("project_name", "—")],
        ["Country", project_info.get("country", "—")],
        ["City", project_info.get("city", "—")],
        ["Climate Zone", project_info.get("climate_zone", "—")],
        ["Typology", project_info.get("typology", "—")],
        ["Weather File", project_info.get("weather", "—")]
    ]

    info_table = Table(info_data, colWidths=[2.0 * inch, 5.4 * inch])
    info_table.setStyle([
        ("BACKGROUND", (0, 0), (0, -1), colors.whitesmoke),
        ("BOX", (0, 0), (-1, -1), 0.5, colors.grey),
        ("INNERGRID", (0, 0), (-1, -1), 0.25, colors.lightgrey),
        ("FONT", (0, 0), (-1, -1), "Helvetica", 9),
        ("FONT", (0, 0), (0, -1), "Helvetica-Bold", 9),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ])

    elements.append(info_table)
    elements.append(Spacer(1, 0.32 * inch))
    # -------------------------------
    # KPI TABLE DATA
    # -------------------------------
    kpi_pairs = [
        ("Area (ft²)", f"{values.get('area', 0):,.0f}" if values.get("area") else "—"),
        ("Conditioned Area (%)", f"{values.get('condArea', 0):.1f}" if values.get("condArea") else "—"),
        ("Window-to-Wall Ratio (%)", f"{values.get('wwr', 0):.1f}" if values.get("wwr") else "—"),
        ("Wall-to-Floor Ratio", f"{values.get('wfr', 0):.2f}" if values.get("wfr") else "—"),
        ("Above-Grade Area (%)", f"{values.get('agArea', 0):,.1f}" if values.get("agArea") else "—"),
        ("Envelope-to-Floor Ratio", f"{values.get('envelopeFloorArea', 0):.2f}" if values.get("envelopeFloorArea") else "—"),
        ("Estimated Hours of Use", f"{values.get('estimateHrsUse', 0):,.0f} h" if values.get("estimateHrsUse") else "—"),
        ("", values.get("", "")),
    ]

    table_data = []
    for i in range(0, len(kpi_pairs), 2):
        left = kpi_pairs[i]
        right = kpi_pairs[i + 1]

        table_data.append([
            Paragraph(left[0], kpi_bold),
            Paragraph(str(left[1]), kpi_normal),
            Paragraph(right[0], kpi_bold),
            Paragraph(str(right[1]), kpi_normal),
        ])

    kpi_table = Table(
        table_data,
            colWidths=[
                1.85 * inch,   # Label (left)
                1.85 * inch,   # Value
                1.85 * inch,   # Label (right)
                1.85 * inch    # Value
            ]
    )
    
    kpi_table.setStyle([
        ("BOX", (0, 0), (-1, -1), 0.5, colors.grey),
        ("INNERGRID", (0, 0), (-1, -1), 0.25, colors.lightgrey),
        ("BACKGROUND", (0, 0), (0, -1), colors.whitesmoke),
        ("BACKGROUND", (2, 0), (2, -1), colors.whitesmoke),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ])

    elements.append(kpi_table)
    elements.append(Spacer(1, 0.35 * inch))

    # -------------------------------
    # CHART GRID (2 PER ROW)
    # -------------------------------
    table_data = []
    row = []

    for img_path in image_paths:
        img = Image(img_path, width=3.8 * inch, height=2.4 * inch)
        row.append(img)

        if len(row) == 2:
            table_data.append(row)
            row = []

        if len(table_data) == 5:
            elements.append(Table(table_data, colWidths=[3.7 * inch, 3.7 * inch]))
            elements.append(Spacer(1, 0.35 * inch))
            table_data = []

    if row:
        table_data.append(row)

    if table_data:
        elements.append(Table(table_data, colWidths=[3.7 * inch, 3.7 * inch]))
    elements.append(Spacer(1, 0.35 * inch))
    elements.append(PageBreak())

    elements.append(
        Paragraph(
            "<b>About <font color='#CA3232'>AWESIM</font></b>",
            ParagraphStyle(
                name="AboutTitle",
                fontSize=11,
                spaceAfter=8
            )
        )
    )

    elements.append(
        Paragraph(
            "AWESIM enables an energy simulation expert to discover optimal design parameters "
            "through systematic parametric investigations in just a few clicks. "
            "These investigations are accessible through interactive charts and "
            "downloadable reports. Please feel free to share your feedback with us at "
            "<u>support@edsglobal.com</u>.",
            ParagraphStyle(
                name="AboutText",
                fontSize=9,
                leading=13
            )
        )
    )
    elements.append(Spacer(1, 0.25 * inch))

    elements.append(
        Paragraph(
            "<b>About <font color='#CA3232'>EDS</font></b>",
            ParagraphStyle(
                name="AboutTitle",
                fontSize=11,
                spaceAfter=8
            )
        )
    )

    elements.append(
        Paragraph(
            "Environmental Design Solutions [EDS] is a sustainability advisory firm. "
            "Since 2002, EDS has worked on over 500 green building and energy efficiency projects worldwide. "
            "The team focuses on climate change mitigation, low-carbon design, building simulation, performance audits, and capacity building. ",
            ParagraphStyle(
                name="AboutText",
                fontSize=9,
                leading=13
            )
        )
    )
    elements.append(Spacer(1, 0.25 * inch))

    elements.append(
        Paragraph(
            "<b>Disclaimer</b>",
            ParagraphStyle(
                name="DA_Title",
                fontSize=11,
                spaceAfter=8
            )
        )
    )

    elements.append(
        Paragraph(
            "AWECM Sim (referred to as the “Application” hereafter) is an outcome of the best "
            "efforts of building simulation experts and IT developers at "
            "<b>Environmental Design Solutions Limited (EDS).</b>"
            "<br/>"

            "<br/>• While EDS has undertaken rigorous due diligence and testing of this Application, "
            "EDS does not assume responsibility for outcomes resulting from its use. "
            "By using this Application, the User indemnifies EDS against any damages whatsoever."
            "<br/>"

            "<br/>• This Application is hosted on private server infrastructure managed by EDS. "
            "EDS does not guarantee uninterrupted availability of the Application. "
            "By using this Application, the User agrees to share uploaded information with EDS "
            "for analysis and future research purposes."
            "<br/>"

            "<br/>• This Application utilizes the following open-source or free-to-use resources:"
            "<br/>"
            " – DOE-2.2<br/>"
            " – Streamlit<br/>"
            " – Python"
            "<br/>"

            "<br/>• EDS is not liable to inform Users about updates to the Application or underlying resources.",
            ParagraphStyle(
                name="DA_Text",
                fontSize=9,
                leading=14
            )
        )
    )

    elements.append(Spacer(1, 0.25 * inch))
    elements.append(
        Paragraph(
            "<b>Acknowledgement</b>",
            ParagraphStyle(
                name="DA_Title",
                fontSize=11,
                spaceAfter=8
            )
        )
    )

    elements.append(
        Paragraph(
            "This application uses:"
            "<br/>"
            "• eQuest © James J. Hirsch &amp; Associates under license."
            "<br/>"
            "• Streamlit, © Streamlit Inc., licensed under Apache 2.0"
            "<br/>"
            "• Python © Python Software Foundation, licensed under the PSF License Version 2",
            ParagraphStyle(
                name="DA_Text",
                fontSize=9,
                leading=14
            )
        )
    )

    elements.append(Spacer(1, 0.25 * inch))
    image_path = "images/image123456.png"

    # elements.append(Spacer(1, 0.2 * inch))

    elements.append(
        Image(
            image_path,
            width=7.4 * inch,   # adjust as needed
            height=1.85 * inch   # adjust as needed
        )
    )

    elements.append(Spacer(1, 0.5 * inch))

    # # -------------------------------
    # # DISCLAIMER
    # # -------------------------------
    # elements.append(
    #     Paragraph(
    #         "<b>Disclaimer</b>",
    #         ParagraphStyle(
    #             name="DisclaimerTitle",
    #             fontSize=10,
    #             spaceAfter=6
    #         )
    #     )
    # )

    # elements.append(
    #     Paragraph(
    #         "This report is generated based on energy simulation results derived from "
    #         "standard inputs and user-provided data. Actual building performance may vary "
    #         "due to construction practices, operational behavior, and climatic variations. "
    #         "AWESIM is the outcome of best efforts by simulation experts at EDS and does not "
    #         "assume responsibility for outcomes from this application. The user indemnifies "
    #         "EDS of any damages.",
    #         ParagraphStyle(
    #             name="DisclaimerText",
    #             fontSize=9,
    #             leading=12,
    #             textColor=colors.grey
    #         )
    #     )
    # )

    pdf.build(
        elements,
        onFirstPage=lambda c, d: add_header_footer(c, d, title, logo_path, awesim_logo),
        onLaterPages=add_later_pages
    )


# def create_pdf(image_paths, project_info, values, pdf_name="Energy_Parametric_Report.pdf"):
#     styles = getSampleStyleSheet()

#     pdf = SimpleDocTemplate(
#         pdf_name,
#         pagesize=A4,
#         leftMargin=30,
#         rightMargin=30,
#         topMargin=70,
#         bottomMargin=45   # ⬅ ensures footer never overlaps
#     )
#     elements = []
#     # -------------------------------
#     # HEADER INFO
#     # -------------------------------
#     title = "Automating Workflows for Energy Simulation"
#     logo_path = "images/EDSlogo.jpg"  # OR absolute path
#     awesim_logo = "images/analysis.png"
#     # -------------------------------
#     # PROJECT INFORMATION TABLE
#     # -------------------------------
#     info_data = [
#         ["Project Name", project_info.get("project_name", "—")],
#         ["Country", project_info.get("country", "—")],
#         ["City", project_info.get("city", "—")],
#         ["Climate Zone", project_info.get("climate_zone", "—")],
#         ["Typology", project_info.get("typology", "—")],
#         ["Weather File", project_info.get("weather", "—")],
#         ["Report Date", project_info.get("report_date", "—")],
#     ]

#     info_table = Table(info_data, colWidths=[2.0 * inch, 5.4 * inch])
#     info_table.setStyle([
#         ("BACKGROUND", (0, 0), (0, -1), colors.whitesmoke),
#         ("BOX", (0, 0), (-1, -1), 0.5, colors.grey),
#         ("INNERGRID", (0, 0), (-1, -1), 0.25, colors.lightgrey),
#         ("FONT", (0, 0), (-1, -1), "Helvetica", 9),
#         ("FONT", (0, 0), (0, -1), "Helvetica-Bold", 9),
#         ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
#     ])

#     elements.append(info_table)
#     elements.append(Spacer(1, 0.32 * inch))

#     # -------------------------------
#     # KEY DESIGN & ENERGY PARAMETERS
#     # -------------------------------

#     elements.append(
#         Paragraph(
#             "<b>Energy Parameters</b>",
#             ParagraphStyle(
#                 name="KpiTitle",
#                 fontSize=10,
#                 spaceAfter=8
#             )
#         )
#     )
#     values_data = [
#         ["Area (ft²)", values.get("area", "—")],
#         ["Conditioned Area", values.get("condArea", "—")],
#         ["Window-to-Wall Ratio", values.get("wwr", "—")],
#         ["Wall-to-Floor Ratio", values.get("wfr", "—")],
#         ["Climate Zone", values.get("climate", "—")],
#         ["Above-Grade Area", values.get("agArea", "—")],
#         ["Envelope-to-Floor Ratio", values.get("envelopeFloorArea", "—")],
#         ["Estimated Hours of Use", values.get("estimateHrsUse", "—")],

#         ["Wall (%)", values.get("wallpct", "—")],
#         ["Roof (%)", values.get("roofpct", "—")],
#         ["Glazing (%)", values.get("glazingpct", "—")],
#         ["WWR (%)", values.get("wwrpct", "—")],
#         ["Lighting (%)", values.get("lightpct", "—")],
#         ["Equipment (%)", values.get("equippct", "—")],
#         ["Energy Range", values.get("EnergyRange", "—")],
#         ["Demand Range", values.get("rangeDemand", "—")],
#     ]

#     kpi_pairs = [
#         ("Area (ft²)", f"{values.get('area', 0):,.0f}" if values.get("area") else "—"),
#         ("Conditioned Area (%)", f"{values.get('condArea', 0):.1f}" if values.get("condArea") else "—"),
#         ("Window-to-Wall Ratio (%)", f"{values.get('wwr', 0):.1f}" if values.get("wwr") else "—"),
#         ("Wall-to-Floor Ratio", f"{values.get('wfr', 0):.2f}" if values.get("wfr") else "—"),
#         ("Climate Zone", values.get("climate", "—")),
#         ("Above-Grade Area (%)", f"{values.get('agArea', 0):,.1f}" if values.get("agArea") else "—"),
#         ("Envelope-to-Floor Ratio", f"{values.get('envelopeFloorArea', 0):.2f}" if values.get("envelopeFloorArea") else "—"),
#         ("Estimated Hours of Use", f"{values.get('estimateHrsUse', 0):,.0f} h" if values.get("estimateHrsUse") else "—"),

#         ("Wall ERP (%)", f"{values.get('wallpct', 0):.1f}" if values.get("wallpct") else "0.0"),
#         ("Roof ERP (%)", f"{values.get('roofpct', 0):.1f}" if values.get("roofpct") else "0.0"),
#         ("Glazing ERP (%)", f"{values.get('glazingpct', 0):.1f}" if values.get("glazingpct") else "0.0"),
#         ("WWR ERP (%)", f"{values.get('wwrpct', 0):.1f}" if values.get("wwrpct") else "0.0"),
#         ("Lighting ERP (%)", f"{values.get('lightpct', 0):.1f}" if values.get("lightpct") else "0.0"),
#         ("Equipment ERP (%)", f"{values.get('equippct', 0):.1f}" if values.get("equippct") else "0.0"),

#         ("Energy Range (kWh)", f"{values.get('EnergyRange', 0):.0f}"),
#         ("Demand Range (kW)", f"{values.get('rangeDemand', 0):.0f}")
#     ]


#     table_data = []
#     for i in range(0, len(kpi_pairs), 2):
#         left = kpi_pairs[i]
#         right = kpi_pairs[i + 1]

#         # table_data.append([
#         #     Paragraph(f"<b>{left[0]}</b>", styles["Normal"]),
#         #     Paragraph(left[1], styles["Normal"]),
#         #     Paragraph(f"<b>{right[0]}</b>", styles["Normal"]),
#         #     Paragraph(right[1], styles["Normal"]),
#         # ])
#         table_data.append([
#             Paragraph(f"<b>{left[0]}</b>", styles["Normal"]),
#             Paragraph(str(left[1]), styles["Normal"]),
#             Paragraph(f"<b>{right[0]}</b>", styles["Normal"]),
#             Paragraph(str(right[1]), styles["Normal"]),
#         ])

#     kpi_table = Table(
#         table_data,
#         colWidths=[
#             2.0 * inch,  # Label (left)
#             1.7 * inch,  # Value
#             1.8 * inch,  # Label (right)
#             1.7 * inch   # Value
#         ],
#         hAlign="LEFT"
#     )

#     kpi_table.setStyle([
#         # Grid & borders (same visual weight as info table)
#         ("BOX", (0, 0), (-1, -1), 0.5, colors.grey),
#         ("INNERGRID", (0, 0), (-1, -1), 0.25, colors.lightgrey),

#         # Backgrounds
#         ("BACKGROUND", (0, 0), (0, -1), colors.whitesmoke),  # 1st column (labels)
#         ("BACKGROUND", (2, 0), (2, -1), colors.whitesmoke),  # 3rd column (labels)

#         # Alignment
#         ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
#         ("ALIGN", (1, 0), (1, -1), "LEFT"),
#         ("ALIGN", (3, 0), (3, -1), "LEFT"),

#         # Fonts
#         ("FONT", (0, 0), (0, -1), "Helvetica-Bold", 9),
#         ("FONT", (2, 0), (0, -1), "Helvetica-Bold", 9),
#         ("FONT", (1, 0), (-1, -1), "Helvetica", 9),
#         ("FONT", (3, 0), (-1, -1), "Helvetica", 9),

#         # Padding → THIS fixes the cramped look
#         ("LEFTPADDING", (0, 0), (-1, -1), 8),
#         ("RIGHTPADDING", (0, 0), (-1, -1), 8),
#         ("TOPPADDING", (0, 0), (-1, -1), 6),
#         ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
#     ])

#     elements.append(kpi_table)
#     elements.append(Spacer(1, 0.35 * inch))

#     # Add Note - 
#     elements.append(
#         Paragraph(
#             "<b>Note</b>",
#             ParagraphStyle(
#                 name="NoteTitle",
#                 fontSize=10,
#                 spaceAfter=6
#             )
#         )
#     )
#     elements.append(
#         Paragraph(
#             "Energy Reduction Potential (ERP) represents the relative reduction in total energy use "
#             "compared to the As Designed case, evaluated by varying one parameter at a time while keeping all other "
#             "inputs constant.<br/><br/>"
#             "<b>ERP (%)</b> is computed as:<br/>"
#             "(As Designed energy - Minimum energy outcome for the parameter) / As Designed energy x 100.<br/><br/>",
#             ParagraphStyle(
#                 name="NoteText",
#                 fontSize=9,
#                 leading=12,
#                 textColor=colors.grey
#             )
#         )
#     )

#     # -------------------------------
#     # CHART GRID (2 PER ROW)
#     # -------------------------------
#     table_data = []
#     row = []

#     for img_path in image_paths:
#         img = Image(img_path, width=3.8 * inch, height=2.4 * inch)
#         row.append(img)

#         if len(row) == 2:
#             table_data.append(row)
#             row = []

#         if len(table_data) == 5:   # 5 rows per page
#             elements.append(Table(table_data, colWidths=[4 * inch, 4 * inch]))
#             elements.append(Spacer(1, 0.35 * inch))
#             table_data = []

#     if row:
#         table_data.append(row)

#     if table_data:
#         elements.append(Table(table_data, colWidths=[4 * inch, 4 * inch]))
#     elements.append(Spacer(1, 0.35 * inch))

#     elements.append(
#         Paragraph(
#             "<b>Disclaimer</b>",
#             ParagraphStyle(
#                 name="DisclaimerTitle",
#                 fontSize=10,
#                 spaceAfter=6
#             )
#         )
#     )

#     elements.append(
#         Paragraph(
#             "This report is generated based on energy simulation results derived from "
#             "standard inputs, and user-provided data. Actual building "
#             "performance may vary due to construction practices, operational behavior, "
#             "and climatic variations. AWESIM is the outcome of best efforts by simulation "
#             "experts at EDS and does not assume responsibility for outcomes from this application. "
#             "ECM comparison through interactive charts and downloadable reports. The user indemnifies EDS of any damages.",
#             ParagraphStyle(
#                 name="DisclaimerText",
#                 fontSize=9,
#                 leading=12,
#                 textColor=colors.grey
#             )
#         )
#     )

#     pdf.build(
#         elements,
#         onFirstPage=lambda c, d: add_header_footer(c, d, title, awesim_logo, logo_path),
#         onLaterPages=add_later_pages
#     )

def add_disclaimer(canvas, doc):
    canvas.saveState()
    disclaimer_text = (
        "This report is generated using AWESIM for preliminary energy analysis only. "
        "EDS does not guarantee accuracy of results.")

    canvas.setFont("Helvetica", 8)
    canvas.setFillColorRGB(0.4, 0.4, 0.4)

    # Draw a separator line
    canvas.line(
        doc.leftMargin,
        0.9 * inch,
        doc.pagesize[0] - doc.rightMargin,
        0.9 * inch
    )

    # Footer text (centered)
    canvas.drawCentredString(
        doc.pagesize[0] / 2,
        0.6 * inch,
        disclaimer_text
    )

    # Page number (optional)
    canvas.drawRightString(
        doc.pagesize[0] - doc.rightMargin,
        0.6 * inch,
        f"Page {doc.page}"
    )

    canvas.restoreState()

# --------------------------------------------------
# ENERGY REDUCTION FUNCTION
# --------------------------------------------------
def calc_energy_reduction(df, baseline):
    if df.empty or baseline == 0:
        return None
    min_energy = df["Energy_Outcome(KWH)"].min()
    return round((baseline - min_energy) / baseline * 100, 1)

if "script_choice" not in st.session_state:
    st.session_state.script_choice = "home"

if "tools_dropdown" not in st.session_state:
    st.session_state.tools_dropdown = "Select"

if "reset_tools" not in st.session_state:
    st.session_state.reset_tools = False

# --- Header Section ---
# 🔁 Reset dropdown BEFORE widget is created
if st.session_state.reset_tools:
    st.session_state.tools_dropdown = "Select"
    st.session_state.reset_tools = False

col1, col2, col3, _, col4, col5 = st.columns([0.16,0.16,2.3,0.4,0.5,0.5])
with col1:
    st.image("images/analysis.png")
with col2:
    st.image("images/eds.png")
with col3:
    st.markdown("""
    <div style="text-align:center;">
        <h3 style="margin-bottom:2px; color:black;">
            <span style="color: rgb(202, 50, 50);">
                Automating Workflows for Energy Simulation
            </span>
        </h3>
    </div>
    """, unsafe_allow_html=True)

with col4:
    tool = st.selectbox(
        "",
        ["ParSim", "ComSim"],
        key="tools_dropdown",
        label_visibility="collapsed"
    )

    if tool == "ParSim":
        st.session_state.script_choice = "tool1"

    elif tool == "ComSim":
        st.session_state.script_choice = "tool2"

with col5:
    if st.button("Home", key="home"):
        st.session_state.script_choice = "home"
        st.session_state.reset_tools = True   # ✅ trigger reset

# Reduce gap before red line
st.markdown("""<div style='margin-top:-40px;'><hr style="border:1px solid red"></div>""",unsafe_allow_html=True)

# ---------------------- HEADER STYLES -------------------------
st.markdown("""<style></style>""", unsafe_allow_html=True)
if st.session_state.script_choice == "home":
    # ---------------------- THREE COLUMN CARDS -------------------------
    col1, col2, col3 = st.columns(3)
    # --------- Card 1: AWESim ----------
    with col1:
        st.image("images/awesim.png")
    # --------- Card 2: PARSim ----------
    with col2:
        col1, col2 = st.columns(2)
        with col1:
            st.image("images/6.png")
        with col2:
            if st.button("ParSim"):
                st.session_state.script_choice = "tool1"
                
    # --------- Card 3: COMSim ----------
    with col3:
        col1, col2 = st.columns(2)
        with col1:
            st.image("images/8.png")
        with col2:
            if st.button("ComSim", key="1"):
                st.session_state.script_choice = "tool2"

    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown(
            '<span style="font-size:24px; font-weight:600;">About <span style="color: rgb(202,50,50);">AWESim</span></span>',
            unsafe_allow_html=True
        )
        # st.markdown('<span style="font-size:24px; font-weight:600;">About <span style="color:red;">AWESim</span></span>',unsafe_allow_html=True)
        st.write("""AWESIM enables an energy simulation expert to discover optimal design parameters through systematic parametric investigations in just a few clicks. Please share your feedback with us.
        These investigations are accessible through interactive charts and downloadable reports. Please feel free to share your feedback with us at **support@edsglobal.com**
        """)
    with col2:
        st.markdown('<span style="font-size:24px; font-weight:600;">Parametric Analysis</span>',unsafe_allow_html=True)
        st.write("""AWESim enables systematic parametric investigations with ease. Interactive charts and downloadable reports provide clear insights. 
        These investigations are accessible through interactive charts and downloadable reports. We keep improving this tool — please send feedback.""")
    with col3:
        st.markdown('<span style="font-size:24px; font-weight:600;">ECM Comparison</span>',unsafe_allow_html=True)
        st.write("""AWESim enables quick ECM comparison through interactive charts and downloadable reports. Please share your feedback with us. These investigations are accessible through interactive charts and downloadable reports. Please share your feedback with us.""")

    # ---------------------- THREE COLUMN CARDS -------------------------
    col1, col2, col3 = st.columns(3)
    # --------- Card 1: AWESim ----------
    with col1:
        st.markdown('<span style="font-size:24px; font-weight:600;">About <span style="color: rgb(202,50,50);">EDS</span></span>',unsafe_allow_html=True)
        st.write("""Environmental Design Solutions [EDS] is a sustainability advisory firm. Since 2002, EDS has worked on over 500 green building and energy efficiency projects worldwide.
        The team focuses on climate change mitigation, low-carbon design, building simulation, performance audits, and capacity building.""")
        st.markdown("</div>", unsafe_allow_html=True)
    # --------- Card 2: PARSim ----------
    with col2:
        st.markdown('<span style="font-size:24px; font-weight:600;">This Version</span>',unsafe_allow_html=True)
        st.write("""
        **Build v1.0.0**

        **Fixes**-
        No Fixes right now!  
        """)
        st.markdown("</div>", unsafe_allow_html=True)
    # --------- Card 3: COMSim ----------
    with col3:
        st.markdown('<span style="font-size:24px; font-weight:600;">Disclaimer & Acknowledgement</span>',unsafe_allow_html=True)
        st.markdown("""
            **AWECM Sim** (referred to as the *“Application”* hereafter) is an outcome of the best efforts of  
            building simulation experts and IT developers at **Environmental Design Solutions Limited (EDS)**.""")
        if st.button("Read more"):
            st.markdown("""
                - While EDS has undertaken rigorous due diligence and testing of this Application,  
                EDS does **not** assume responsibility for outcomes resulting from its use.  
                By using this Application, the *User* indemnifies EDS against any damages whatsoever.

                - This Application is hosted on private server infrastructure managed by EDS.  
                EDS does not guarantee uninterrupted availability of the Application.  
                By using this Application, the User agrees to share uploaded information with EDS  
                for analysis and future research purposes.

                - This Application utilizes open-source or free-to-use resources:
                    - DOE-2.2  
                    - Streamlit  
                    - Python

                - EDS is not liable to inform Users about updates to the Application or underlying resources.
            """)
        st.markdown("</div>", unsafe_allow_html=True)

if st.session_state.script_choice == "tool1":
    # Load location database
    csv_path = resource_path(os.path.join("database", "Simulation_locations.csv"))
    weather_df = pd.read_csv(csv_path)
    db_path = resource_path(os.path.join("database", "AllData.xlsx"))
    output_csv = resource_path(os.path.join("database", "Randomized_Sheet.xlsx"))
    location_path = resource_path(os.path.join("database", "Simulation_locations.csv"))
    location = pd.read_csv(location_path)
    locations = sorted(location['Sim_location'].unique())
    updated_df = pd.read_excel(output_csv)

    updated_df["Wall_Roof_Glazing_WWR_GlazingR_Light_Equip"] = (
        updated_df["Wall"].astype(str) + "_" +
        updated_df["Roof"].astype(str) + "_" +
        updated_df["Glazing"].astype(str) + "_" +
        updated_df["WWR"].astype(str) + "_" +
        updated_df["GlazingR"].astype(str) + "_" +
        updated_df["Light"].astype(str) + "_" +
        updated_df["Equip"].astype(str)
    )
    # Inputs
    col1, col2, col3 = st.columns([1,1,2])
    st.markdown("""
        <style>
        div[data-testid="stFileUploader"] section {
            padding: 0.01rem 0 !important;  /* Reduce vertical padding */
        }
        div[data-testid="stFileUploader"] div[role="button"] {
            padding: 0.0rem 0.0rem !important;  /* Reduce height of clickable area */
            font-size: 0.00rem !important;      /* Smaller text */
        }
        </style>
    """, unsafe_allow_html=True)
    # --- NEW: handle both Country and Location ---
    if "Country" in location.columns:
        # If CSV already has Country column
        countries = sorted(location["Country"].unique().tolist())
    else:
        # If your current CSV has only Indian cities — default to India
        countries = ["India"]
    with col1:
        st.write("📝 Project Name")
        project_name = st.text_input("", placeholder="Enter project name", label_visibility="collapsed")
        
        # project_name = st.text_input("📝 Project Name", placeholder="Enter project name")
        project_name_clean = project_name.replace(" ", "")
        user_nm = project_name_clean
        if project_name_clean:
            parent_dir = os.path.dirname(os.getcwd())
            batch_outputs_dir = os.path.join(parent_dir, "Batch_Outputs")
            project_folder = os.path.join(batch_outputs_dir, project_name_clean)
            # Check if project folder already exists
            if os.path.exists(project_folder):
                st.warning("⚠️ Project name already exists! Please select another name.")
                # st.stop()
    # with col2:
        # Add "Other" option to countries
        countries.append("Custom Weather")
        st.write("🌎 Select Country")
        selected_country = st.selectbox("", countries, label_visibility="collapsed")

        # Filter and sort locations for the selected country (if not "Other")
        if selected_country != "Custom Weather":
            if "Country" in location.columns:
                filtered_locations = (
                    location[location["Country"] == selected_country]["Sim_location"]
                    .dropna()
                    .unique()
                    .tolist()
                )
                filtered_locations = sorted(filtered_locations)
            else:
                filtered_locations = sorted(location["Sim_location"].dropna().tolist())
        else:
            filtered_locations = []  # No locations for "Other"
    with col2:
        # Main typologies
        main_typologies = ["Office", "Retail", "Hospital", "Hotel", "Residential"]
        st.write("🌆 Select Typology")
        selected_typology = st.selectbox("", main_typologies, label_visibility="collapsed", key="typology")

    bin_name = ""
    # Only show City dropdown if not "Other"
    if selected_country != "Custom Weather":
        with col2:
            st.write("🌎 Select City")
            user_input = st.selectbox("", filtered_locations, label_visibility="collapsed").lower()
            selected_ = location[location["Sim_location"].str.lower() == user_input.lower()]
            # Extract Ashrae Climate Zone
            if not selected_.empty:
                ashrae_zone = selected_["Ashrae Climate Zone"].iloc[0]
    else:
        with col2:
            user_input = "Other-City"
            # When "Other" is selected, show .bin upload option
            st.write("📤 Upload .bin file")
            uploaded_bin = st.file_uploader("", type=["bin"], label_visibility="collapsed")
            if uploaded_bin is not None:
                save_folder = r"C:\doe22\weather"
                os.makedirs(save_folder, exist_ok=True)
                # Always rename uploaded file to 1.bin
                save_path = os.path.join(save_folder, "1.bin")
                # Save uploaded file as 1.bin
                with open(save_path, "wb") as f:
                    f.write(uploaded_bin.getbuffer())
                bin_name = "1"   # without extension
            
    with col3:
        uploaded_file = st.file_uploader("📤 Upload eQUEST INP file", type=["inp"])
        if uploaded_file:
            # st.write(uploaded_file)
            if uploaded_file.name != 'F1.inp':
                uploaded_file.name = user_nm + '.inp'

            # Go one step outside current working directory
            parent_dir = os.path.dirname(os.getcwd())
            batch_outputs_dir = os.path.join(parent_dir, "Batch_Outputs")
            os.makedirs(batch_outputs_dir, exist_ok=True)
            uploaded_inp_path = os.path.join(batch_outputs_dir, uploaded_file.name)
            with open(uploaded_inp_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            output_inp_folder = os.path.dirname(uploaded_inp_path)
            inp_folder = output_inp_folder
            project_folder = os.path.join(batch_outputs_dir, project_name_clean)
    run_cnt = 1
    location_id, weather_path = "", ""
    if user_input != "Other-City":
        matched_row = weather_df[weather_df['Sim_location'].str.lower().str.contains(user_input)]
    else:
        matched_row = pd.DataFrame()
    if not matched_row.empty:
        location_id = matched_row.iloc[0]['Location_ID']
        weather_path = matched_row.iloc[0]['Weather_file']
    elif user_input:
        weather_path = bin_name
    
    if "all_figs" not in st.session_state:
        st.session_state.all_figs = []

    if st.button("Simulate 🚀"):
        st.cache_data.clear()
        st.cache_resource.clear()
        if uploaded_file is None:
            st.warning("⚠️ Please Upload .INP File!")
            st.stop()
        if bin_name is None:
            st.warning("⚠️ Please Upload .BIN File!")
            st.stop()
        if not project_name_clean:
            st.warning("⚠️ Please enter a project name.")
            st.stop()
        with st.spinner("⚡ Processing... This may take a few minutes."):
            os.makedirs(output_inp_folder, exist_ok=True)
            new_batch_id = f"{int(time.time())}"  # unique ID

            selected_rows = updated_df[updated_df['Batch_ID'] == run_cnt]
            batch_output_folder = os.path.join(output_inp_folder, f"{user_nm}")
            os.makedirs(batch_output_folder, exist_ok=True)

            num = 1
            modified_files = []
            for _, row in selected_rows.iterrows():
                selected_inp = uploaded_file.name
                new_inp_name = f"{row['Wall']}_{row['Roof']}_{row['Glazing']}_{row['GlazingR']}_{row['Light']}_{row['WWR']}_{row['Equip']}_{selected_inp}"
                new_inp_path = os.path.join(batch_output_folder, new_inp_name)

                inp_file_path = os.path.join(inp_folder, selected_inp)
                if not os.path.exists(inp_file_path):
                    st.error(f"File {inp_file_path} not found. Skipping.")
                    continue

                # st.info(f"Modifying INP file {num}: {selected_inp} -> {new_inp_name}")
                num += 1

                # Apply modifications
                inp_content = wwr.process_window_insertion_workflow(inp_file_path, row["WWR"])
                # inp_content = orient.updateOrientation(inp_content, row["Orient"])
                inp_content = lighting.updateLPD(inp_content, row['Light'])
                # inp_content = insertWall.update_Material_Layers_Construction(inp_content, row["Wall"])
                # inp_content = insertRoof.update_Material_Layers_Construction(inp_content, row["Roof"])
                # inp_content = insertRoof.removeDuplicates(inp_content)
                inp_content = equip.updateEquipment(inp_content, row['Equip'])
                inp_content = windows.insert_glass_types_multiple_outputs(inp_content, row['Glazing'])
                inp_content = windows.insert_glass_UVal(inp_content, row['GlazingR'])
                inp_content =remove_utility(inp_content)
                # inp_content = windows.readSCUVal(inp_content, row['Glazing'])
                if row['Light'] > 0:
                    inp_content =remove_betweenLightEquip(inp_content)
                count = ModifyWallRoof.count_exterior_walls(inp_content)
                if count > 1:
                    inp_content = ModifyWallRoof.fix_walls(inp_content, row["Wall"])
                    inp_content = ModifyWallRoof.fix_roofs(inp_content, row["Roof"])
                    # inp_content = insertRoof.removeDuplicates(inp_content)
                    with open(new_inp_path, 'w') as file:
                        file.writelines(inp_content)
                    modified_files.append(new_inp_name)
                else:
                    st.write("No Exterior-Wall Exists!")
        
            simulate_files = []
            if uploaded_file is None:
                st.error("Please upload an INP file before starting the simulation.")
            else:
                # st.markdown(f"<span style='color:green;'>✅ Updating DAYLIGHTING from YES to NO!</span>", unsafe_allow_html=True)
                script_dir = os.path.dirname(os.path.abspath(__file__))
                shutil.copy(os.path.join(script_dir, "script.bat"), batch_output_folder)
                inp_files = [f for f in os.listdir(batch_output_folder) if f.lower().endswith(".inp")]
                for inp_file in inp_files:
                    file_path = os.path.join(batch_output_folder, os.path.splitext(inp_file)[0])
                    subprocess.call(
                        [os.path.join(batch_output_folder, "script.bat"), file_path, weather_path],
                        shell=True
                    )
                    simulate_files.append(inp_file)
            
                subprocess.call([os.path.join(batch_output_folder, "script.bat"), batch_output_folder, weather_path], shell=True)
                required_sections = ['BEPS', 'BEPU', 'LS-C', 'LV-B', 'LV-D', 'PS-E', 'SV-A']
                log_file_path = check_missing_sections(batch_output_folder, required_sections, new_batch_id, user_nm)
                get_failed_simulation_data(batch_output_folder, log_file_path)
                clean_folder(batch_output_folder)
                combined_Data = get_files_for_data_extraction(batch_output_folder, log_file_path, new_batch_id, location_id, user_nm, user_input, selected_typology)
                combined_Data = combined_Data.reset_index(drop=True)
                
        st.session_state.all_figs = []   # 🔥 MUST RESET HERE
        # exportCSV = resource_path(os.path.join("2026-01-29T11-34_export.csv"))
        # combined_Data = pd.read_csv(exportCSV)
        combined_Data["Equip(W/Sqft)"] = combined_Data["Equipment-Total(W)"] / combined_Data["Floor-Total-Above-Grade(SQFT)"]
        combined_Data["Light(W/Sqft)"] = combined_Data["Power Lighting Total(W)"] / combined_Data["Floor-Total-Above-Grade(SQFT)"]
        path = os.path.join(os.path.dirname(__file__), "AllData.xlsx")
        r_val = pd.read_excel(path, sheet_name="wallNew")
        wall_to_rvalue = {i+1: float(r) for i, r in enumerate(r_val["R-Value"].dropna())}

        # Extract wall and roof codes
        combined_Data["WallCode"] = combined_Data["FileName"].str.split("_").str[0].astype(int)
        combined_Data["RoofCode"] = combined_Data["FileName"].str.split("_").str[1].astype(int)

        # Map to R-values
        combined_Data["R-Value-Wall"] = combined_Data["WallCode"].map(wall_to_rvalue)
        combined_Data["R-Value-Roof"] = combined_Data["RoofCode"].map(wall_to_rvalue)

        # Drop intermediate columns if you don’t need them
        combined_Data.drop(columns=["WallCode", "RoofCode"], inplace=True)
        baseline_energy = combined_Data.loc[combined_Data.index[0], "Energy_Outcome(KWH)"]
        combined_Data["Energy_Saving_%"] = ((baseline_energy - combined_Data["Energy_Outcome(KWH)"]) / baseline_energy * 100)
        baseline = combined_Data.iloc[[0]]
        others   = combined_Data.iloc[1:]
        # Split FileName into parts
        split_cols = combined_Data['FileName'].str.split("_", expand=True)

        # Assign names to first 8 parts
        split_cols.columns = ["wall", "roof", "glazing", "glazingr", "light", "wwr", "equip", "suffix"]
        # Convert numeric columns (except suffix)
        # st.write(combined_Data)
        for col in split_cols.columns[:-1]:
            split_cols[col] = pd.to_numeric(split_cols[col], errors="coerce")

        # Merge back
        combined_Data_expanded = pd.concat([combined_Data, split_cols], axis=1)

        # Filter subsets
        wall_df    = combined_Data_expanded[(combined_Data_expanded["wall"] > 0) | (combined_Data_expanded.index == 0)]
        roof_df    = combined_Data_expanded[(combined_Data_expanded["roof"] > 0) | (combined_Data_expanded.index == 0)]
        glazing_df = combined_Data_expanded[(combined_Data_expanded["glazing"] > 0) | (combined_Data_expanded.index == 0)]
        wwr_df     = combined_Data_expanded[(combined_Data_expanded["wwr"] > 0) | (combined_Data_expanded.index == 0)]
        glazingr_df  = combined_Data_expanded[(combined_Data_expanded["glazingr"] > 0) | (combined_Data_expanded.index == 0)]
        light_df   = combined_Data_expanded[(combined_Data_expanded["light"] > 0) | (combined_Data_expanded.index == 0)]
        equip_df   = combined_Data_expanded[(combined_Data_expanded["equip"] > 0) | (combined_Data_expanded.index == 0)]
        combined_Data.columns = combined_Data.columns.str.strip()
        row0 = combined_Data.iloc[0]
        area_ft2 = (row0.get("Floor-Total-Above-Grade(SQFT)", 0) + row0.get("Floor-Total-Below-Grade(SQFT)", 0))
        conditioned_pct = ((row0.get("Floor-Total-Conditioned-Grade(SQFT)", 0) / area_ft2) * 100 if area_ft2 else 0)
        wwr = row0.get("WWR", 0)
        climate_zone = row0.get("Climate Zone", "—")
        above_grade_pct = ((row0.get("Floor-Total-Above-Grade(SQFT)", 0) / area_ft2) * 100 if area_ft2 else 0)
        roof_floor_ratio = (row0.get("ROOF-AREA(SQFT)", 0) / area_ft2 if area_ft2 else None)
        usage_factor = row0.get("EFLH", None)
        if "Energy_Outcome(KWH)" in combined_Data.columns:
            energy_min = combined_Data["Energy_Outcome(KWH)"].min()
            energy_max = combined_Data["Energy_Outcome(KWH)"].max()
        else:
            energy_min = energy_max = None
        # =============================
        # SAFE FORMATTING
        # =============================
        roof_floor_disp = f"{roof_floor_ratio:.2f}" if roof_floor_ratio is not None else "—"
        energy_min_disp = f"{energy_min:,.0f}" if energy_min is not None else "—"
        energy_max_disp = f"{energy_max:,.0f}" if energy_max is not None else "—"
        usage_factor_disp = f"{usage_factor:.1f}" if usage_factor is not None else "—"
        # =============================
        # HEADER
        # =============================
        # st.markdown("<h5 style='text-align:left;'>Analysis Summary</h5>", unsafe_allow_html=True)
        baseline_energy = combined_Data_expanded.loc[0, "Energy_Outcome(KWH)"]
        # --------------------------------------------------
        # BASE CALCULATIONS (FIRST ROW ONLY WHERE REQUIRED)
        # --------------------------------------------------
        above_grade_area = row0["Floor-Total-Above-Grade(SQFT)"]
        below_grade_area = row0["Floor-Total-Below-Grade(SQFT)"]
        area_ft2 = above_grade_area + below_grade_area
        conditioned_area = row0["Floor-Total-Conditioned-Grade(SQFT)"]
        conditioned_pct = (conditioned_area / area_ft2) * 100 if area_ft2 else 0
        wwr = row0["WWR"]
        above_grade_pct = (above_grade_area / area_ft2) * 100 if area_ft2 else 0
        roof_floor_ratio = ((row0["Wall-Total-Above-Grade(SQFT)"] + row0["ROOF-AREA(SQFT)"]) / area_ft2 if area_ft2 else None)
        wall_floor_ratio = (row0["Wall-Total-Above-Grade(SQFT)"] / area_ft2 if area_ft2 else None)
        usage_factor = round(row0["EFLH"], 1) if "EFLH" in row0 else None
        # --------------------------------------------------
        # ENERGY RANGE (WHOLE DATASET)
        # --------------------------------------------------
        energy_min = combined_Data_expanded["Energy_Outcome(KWH)"].min()
        energy_max = combined_Data_expanded["Energy_Outcome(KWH)"].max()
        demand_min = combined_Data_expanded["TOTAL-LOAD(KW)"].min()
        demand_max = combined_Data_expanded["TOTAL-LOAD(KW)"].max()
        baseline_energy = combined_Data_expanded.loc[0, "Energy_Outcome(KWH)"]

        wall_pct    = calc_energy_reduction(wall_df, baseline_energy)
        roof_pct    = calc_energy_reduction(roof_df, baseline_energy)
        glazing_pct = calc_energy_reduction(glazing_df, baseline_energy)
        wwr_pct     = calc_energy_reduction(wwr_df, baseline_energy)
        glazingr_pct  = calc_energy_reduction(glazingr_df, baseline_energy)
        light_pct   = calc_energy_reduction(light_df, baseline_energy)
        equip_pct   = calc_energy_reduction(equip_df, baseline_energy)
        range_energy = energy_max - energy_min
        range_demand = demand_max - demand_min

        # ---------------- STYLES ----------------
        st.markdown("""
        <style>
        .analysis-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(140px, 1fr));
            gap: 10px;
        }

        .card {
            position: relative;
            border-radius: 10px;
            padding: 8px;
            min-height: 70px;
            background: white;
            box-shadow: 0 4px 10px rgba(0,0,0,0.04);
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            text-align: center;
            cursor: help;
            transition: all 0.25s ease;
        }

        .card:hover {
            transform: translateY(-3px);
            box-shadow: 0 8px 20px rgba(0,0,0,0.08);
        }

        .card::before {
            content: "";
            position: absolute;
            left: 0;
            top: 0;
            width: 4px;
            height: 100%;
            
            border-radius: 10px 0 0 10px;
        }

        .info-icon {
            position: absolute;
            top: 6px;
            right: 8px;
            font-size: 12px;
            color: #6b7280;
        }

        .card .tooltip {
            visibility: hidden;
            opacity: 0;
            width: 220px;
            background-color: #111827;
            color: #ffffff;
            border-radius: 6px;
            padding: 8px 10px;
            font-size: 12px;
            line-height: 1.4;
            position: absolute;
            bottom: 115%;
            left: 50%;
            transform: translateX(-50%);
            transition: opacity 0.2s ease-in-out;
            z-index: 100;
        }

        .card:hover .tooltip {
            visibility: visible;
            opacity: 1;
        }

        .card small {
            font-size: 0.7rem;
            font-weight: 500;
            color: #000;
        }

        .card h6 {
            margin-top: 6px;
            font-size: 1.05rem;
            font-weight: 600;
            color: #0f172a;
        }

        .section-title {
            font-size: 1rem;
            font-weight: 600;
            color: #dc2626;
            margin: 18px 0 10px 0;
        }
        .card {
            position: relative;
        }

        .info-icon {
            font-size: 12px;
            cursor: pointer;
            color: #6b7280; /* subtle gray */
        }

        .tooltip {
            position: absolute;
            bottom: 120%;
            left: 50%;
            transform: translateX(-50%) translateY(6px);
            background: #111827;   /* dark slate */
            color: #f9fafb;
            padding: 8px 12px;
            border-radius: 8px;
            font-size: 12px;
            line-height: 1.4;
            max-width: 220px;
            text-align: left;
            white-space: normal;
            box-shadow: 0 8px 24px rgba(0,0,0,0.18);
            opacity: 0;
            pointer-events: none;
            transition: opacity 0.2s ease, transform 0.2s ease;
            z-index: 10;
        }

        /* Tooltip arrow */
        .tooltip::after {
            content: "";
            position: absolute;
            top: 100%;
            left: 50%;
            transform: translateX(-50%);
            border-width: 6px;
            border-style: solid;
            border-color: #111827 transparent transparent transparent;
        }

        /* Show on hover */
        .card:hover .tooltip {
            opacity: 1;
            transform: translateX(-50%) translateY(0);
        }

        </style>
        """, unsafe_allow_html=True)
        st.markdown("""<div style="border-left: 3px solid red; height: 100%; margin: 0 15px;"></div>""",unsafe_allow_html=True)
        col1, col_line, col2 = st.columns([5, 0.2, 5])
        with col1:
            st.markdown(f"""
                <h5 class="section-title" style="color:#dc2626">Analysis Summary</h5>
                <div class="analysis-grid">
                    <div class="card"><span class="info-icon">ⓘ</span><div class="tooltip">Total floor area of the building, including both above-grade and below-grade areas.</div><small>Area (ft²)</small><h6>{area_ft2:,.0f}</h6></div>
                    <div class="card"><span class="info-icon">ⓘ</span><div class="tooltip">Percentage of floor area that is conditioned, including above-grade spaces and below-grade floor areas.</div><small>Conditioned Area</small><h6>{conditioned_pct:.1f}%</h6></div>
                    <div class="card"><span class="info-icon">ⓘ</span><div class="tooltip">Ratio of window area to wall area</div><small>Window-to-Wall Ratio</small><h6>{wwr*100:.1f}%</h6></div>
                    <div class="card"><span class="info-icon">ⓘ</span><div class="tooltip">Wall area divided by floor area</div><small>Wall-to-Floor Ratio</small><h6>{wall_floor_ratio*100:.1f}%</h6></div>
                    <div class="card"><span class="info-icon">ⓘ</span><div class="tooltip">ASHRAE climate classification</div><small>Climate Zone</small><h6 style="color:#16a34a;">{ashrae_zone}</h6></div>
                    <div class="card"><span class="info-icon">ⓘ</span><div class="tooltip">Above-grade envelope percentage</div><small>Above-Grade Area</small><h6>{above_grade_pct:.1f}%</h6></div>
                    <div class="card"><span class="info-icon">ⓘ</span><div class="tooltip">Roof area relative to floor area</div><small>Envelope-to-Floor Ratio</small><h6>{roof_floor_ratio:.2f}</h6></div>
                    <div class="card"><span class="info-icon">ⓘ</span><div class="tooltip">Estimated annual operating hours</div><small>Estimated Hours of Use</small><h6>{usage_factor}</h6></div>
                </div>
            """, 
            unsafe_allow_html=True)
        with col_line:
            st.markdown("<div style='border-left:3px solid red; height:210px'></div>",unsafe_allow_html=True)
        with col2:
            st.markdown(f"""
                <h5 class="section-title" style="color:#dc2626">Parametric Energy Reduction & Ranges</h5>
                <div class="analysis-grid">
                    <div class="card"><span class="info-icon">ⓘ</span><div class="tooltip">Maximum percentage reduction in annual energy consumption achieved through wall design variations, compared to the baseline case.</div><small>Wall</small><h6>{wall_pct:.1f}%</h6></div>
                    <div class="card"><span class="info-icon">ⓘ</span><div class="tooltip">Maximum percentage reduction in annual energy consumption achieved through roof design variations, compared to the baseline case.</div><small>Roof</small><h6>{roof_pct:.1f}%</h6></div>
                    <div class="card"><span class="info-icon">ⓘ</span><div class="tooltip">Maximum percentage reduction in annual energy consumption achieved through glazing performance variations, compared to the baseline case.</div><small>Glazing</small><h6>{glazing_pct:.1f}%</h6></div>
                    <div class="card"><span class="info-icon">ⓘ</span><div class="tooltip">Maximum percentage reduction in annual energy consumption achieved through window-to-wall ratio variations, compared to the baseline case.</div><small>WWR</small><h6>{wwr_pct:.1f}%</h6></div>
                    <div class="card"><span class="info-icon">ⓘ</span><div class="tooltip">Maximum percentage reduction in annual energy consumption achieved through lighting efficiency variations, compared to the baseline case.</div><small>Lighting</small><h6>{light_pct:.1f}%</h6></div>
                    <div class="card"><span class="info-icon">ⓘ</span><div class="tooltip">Maximum percentage reduction in annual energy consumption achieved through equipment efficiency variations, compared to the baseline case.</div><small>Equipment</small><h6>{equip_pct:.1f}%</h6></div>
                    <div class="card"><span class="info-icon">ⓘ</span><div class="tooltip">Difference between the maximum and minimum annual energy consumption observed across all design parameter variations.</div><small>Energy Range (kWh)</small><h6>{range_energy:,.0f}</h6></div>
                    <div class="card"><span class="info-icon">ⓘ</span><div class="tooltip">Difference between the maximum and minimum peak electrical demand observed across all design parameter variations.</div><small>Demand Range (kW)</small><h6>{range_demand:,.1f}</h6></div>
                </div>
                """, 
            unsafe_allow_html=True)

        import plotly.graph_objects as go
        from plotly.subplots import make_subplots
        baseline_energy = combined_Data_expanded.loc[0, "Energy_Outcome(KWH)"]
        combined_Data_expanded["Energy_Saving_%"] = (
            (baseline_energy - combined_Data_expanded["Energy_Outcome(KWH)"])
            / baseline_energy
        ) * 100
        baseline = combined_Data_expanded.loc[[0]]
        wall_param  = combined_Data_expanded.iloc[1:]
        roof_param  = combined_Data_expanded.iloc[1:]
        light_param = light_df.iloc[1:]
        equip_param = equip_df.iloc[1:]
        glazing_param = glazing_df.iloc[1:]
        glazingr_param = glazingr_df.iloc[1:]
        # st.write(glazingr_df)
        # extract number after first '_'
        roof_df["DensityIndex"] = (
            roof_df["FileName"]
            .str.split("_")
            .str[1]
            .astype(int)
        )
        wall_df["DensityIndex"] = (
            wall_df["FileName"]
            .str.split("_")
            .str[0]
            .astype(int)
        )
        glazingr_df["DensityIndex"] = (
            glazingr_df["FileName"]
            .str.split("_")
            .str[2]
            .astype(int)
        )

        # classify density
        roof_df["DensityType"] = roof_df["DensityIndex"].apply(lambda x: "Low Density" if x > 12 else "High Density")
        wall_df["DensityType"] = wall_df["DensityIndex"].apply(lambda x: "Low Density" if x > 12 else "High Density")
        glazingr_df["DensityType"] = glazingr_df["DensityIndex"].apply(lambda x: "Low Density" if x > 12 else "High Density")
        # st.write(roof_df)
        all_figs = []   # <-- collect all charts here
        fig_wall = energy_param_plot_wall(wall_df,baseline,"R-VAL-W","Wall R-Value (h·ft²·°F/Btu)","Energy Use vs Wall R-Value",show_legend=True)
        fig_roof = energy_param_plot(roof_df,baseline,"R-VAL-R","Roof R-Value (h·ft²·°F/Btu)","Energy Use vs Roof R-Value",show_legend=True)
        # all_figs.extend([fig_wall, fig_roof])
        st.session_state.all_figs.extend([fig_wall, fig_roof])
        st.markdown(f"""<br>""", unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1:
            st.plotly_chart(fig_wall, use_container_width=True)
        with col2:
            st.plotly_chart(fig_roof, use_container_width=True)
        fig_light = make_single_plot(light_df,x_col="Light(W/Sqft)",title="Energy Use vs Lighting Power",x_label="Lighting Power Density (W/ft²)")
        fig_equip = make_single_plot(equip_df,x_col="Equip(W/Sqft)",title="Energy Use vs Equipment Power",x_label="Equipment Power Density (W/ft²)")
        all_figs.extend([fig_light, fig_equip])
        st.session_state.all_figs.extend([fig_light, fig_equip])
        col1, col2 = st.columns(2)
        with col1:
            st.plotly_chart(fig_light, use_container_width=True)
        with col2:
            st.plotly_chart(fig_equip, use_container_width=True)
        
        fig_SC_glazing = make_single_plot(glazing_df,x_col="SHGC",title="Energy Use vs SHGC",x_label="SHGC")
        fig_R_glazing = energy_param_plot(glazingr_df,baseline,"R-VAL-Wind","Window R-Value (h·ft²·°F/Btu)","Energy Use vs Window R-Value",show_legend=True)
        # fig_R_glazing = make_single_plot(glazingr_df,x_col="R-VAL-Wind",title="Energy Use vs R-Value",x_label="R-Value (HR·ft²·°F / Btu)")
        st.session_state.all_figs.extend([fig_SC_glazing, fig_R_glazing])
        col1, col2 = st.columns(2)
        with col1:
            st.plotly_chart(fig_SC_glazing, use_container_width=True)
        with col2:
            st.plotly_chart(fig_R_glazing, use_container_width=True)
        st.divider()

        # ---- SAVE IMMEDIATELY ----
        st.session_state.area_ft2 = area_ft2
        st.session_state.conditioned_pct = conditioned_pct
        st.session_state.wwr_pct = wwr_pct
        st.session_state.wall_floor_ratio = wall_floor_ratio
        st.session_state.above_grade_pct = above_grade_pct
        st.session_state.roof_floor_ratio = roof_floor_ratio
        st.session_state.usage_factor = usage_factor

        st.session_state.wall_pct = wall_pct
        st.session_state.roof_pct = roof_pct
        st.session_state.light_pct = light_pct
        st.session_state.equip_pct = equip_pct
        st.session_state.glazing_pct = glazing_pct

        st.session_state.range_energy = range_energy
        st.session_state.range_demand = range_demand
    
    if st.button("Generate Report"):
        if not st.session_state.all_figs:
            st.error("Please run Simulate first.")
            st.stop()

        with st.spinner("Generating Report..."):
            img_paths = save_figs_as_images(st.session_state.all_figs)
            cz = ashrae_zone
            if isinstance(cz, str) and "Ext." in cz:
                cz = cz.replace("Ext.", "Extremely")
            cz = f"{cz}"
            project_info = {
                "project_name": project_name or "—",
                "country": selected_country,
                "city": user_input.title() if user_input else "—",
                "climate_zone": cz,
                "typology": selected_typology,
                "weather": weather_path,
            }
            values = {
                "area": st.session_state.get("area_ft2"),
                "condArea": st.session_state.get("conditioned_pct"),
                "wwr": st.session_state.get("wwr_pct"),
                "wfr": st.session_state.get("wall_floor_ratio"),
                "climate": cz,
                "agArea": st.session_state.get("above_grade_pct"),
                "envelopeFloorArea": st.session_state.get("roof_floor_ratio"),
                "estimateHrsUse": st.session_state.get("usage_factor"),

                "wallpct": st.session_state.get("wall_pct"),
                "roofpct": st.session_state.get("roof_pct"),
                "lightpct": st.session_state.get("light_pct"),
                "equippct": st.session_state.get("equip_pct"),
                "wwrpct": st.session_state.get("wwr_pct"),
                "glazingpct": st.session_state.get("glazing_pct"),

                "EnergyRange": st.session_state.get("range_energy"),
                "rangeDemand": st.session_state.get("range_demand"),
            }

            create_pdf(image_paths=img_paths, project_info=project_info,values=values)
        with open("Energy_Parametric_Report.pdf", "rb") as f:
            st.download_button(
                "⬇️ Download Report",
                data=f,
                file_name="Energy_Parametric_Report.pdf",
                mime="application/pdf"
            )

        st.success("Report Generated!")

st.markdown('<hr style="border:1px solid red">', unsafe_allow_html=True)
st.image("images/image123456.png", width=2000) 
st.markdown(
        """
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.4/css/all.min.css">
    <style>
        .footer {
            background-color: #f8f9fa;
            padding: 20px 0;
            color: #333;
            display: flex;
            justify-content: space-between;
            align-items: center;
            text-align: center;
        }
        .footer .logo {
            flex: 1;
        }
        .footer .logo img {
            max-width: 150px;
            height: auto;
        }
        .footer .social-media {
            flex: 2;
        }
        .footer .social-media p {
            margin: 0;
            font-size: 16px;
        }
        .footer .icons {
            margin-top: 10px;
        }
        .footer .icons a {
            margin: 0 10px;
            color: #666;
            text-decoration: none;
            transition: color 0.3s ease;
        }
        .footer .icons a:hover {
            color: #0077b5; /* LinkedIn color as default */
        }
        .footer .icons a .fab {
            font-size: 28px;
        }
        .footer .additional-content {
            margin-top: 10px;
        }
        .footer .additional-content h4 {
            margin: 0;
            font-size: 18px;
            color: #007bff;
        }
        .footer .additional-content p {
            margin: 5px 0;
            font-size: 16px;
        }
    </style>
    <div style="text-align:center; font-size:14px;">
        Email: <a href="mailto:info@edsglobal.com">info@edsglobal.com</a>&nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;&nbsp;
        Phone: +91 . 11 . 4056 8633&nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;&nbsp;
        <a href="https://twitter.com/edsglobal?lang=en" target="_blank"><i class="fab fa-twitter" style="color:#1DA1F2; margin:0 6px;"></i></a>
        <a href="https://www.facebook.com/Environmental.Design.Solutions/" target="_blank"><i class="fab fa-facebook" style="color:#4267B2; margin:0 6px;"></i></a>
        <a href="https://www.instagram.com/eds_global/?hl=en" target="_blank"><i class="fab fa-instagram" style="color:#E1306C; margin:0 6px;"></i></a>
        <a href="https://www.linkedin.com/company/environmental-design-solutions/" target="_blank"><i class="fab fa-linkedin" style="color:#0077b5; margin:0 6px;"></i></a>
    </div>
    """,
    unsafe_allow_html=True
)