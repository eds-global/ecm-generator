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

def u_to_r(u_w_m2k):
    """Convert U (W/m²K) to R (h·ft²·°F/Btu)"""
    return 5.678 / u_w_m2k

def sqft_to_m2(area_sqft):
    return area_sqft * 0.092903  # exact conversion


def match_area_condition(area_m2, condition):
    if condition == "ALL":
        return True

    if condition.startswith("<"):
        limit = float(condition.replace("<", "").strip())
        return area_m2 < limit

    if condition.startswith(">"):
        limit = float(condition.replace(">", "").strip())
        return area_m2 > limit

    return False


def match_building_category(rule_category, typology):
    if rule_category == "All":
        return True

    # handle multi-category like: "No-Star Hotel, Business"
    categories = [c.strip().lower() for c in rule_category.split(",")]
    return typology.lower() in categories


def get_wall_u_limits(data, ruleset_data, climate_zone, typology):
    """
    data → dataframe row or dict containing area column
    ruleset_data → JSON
    climate_zone → "Composite", "Hot-Dry", etc
    typology → "School", "Business", "Hospital", etc
    """
    limits = {}

    # Wall area in SQFT → convert to m2
    wall_area_sqft = float(data["Wall-Total-Above-Grade(SQFT)"].iloc[0])
    wall_area_m2 = sqft_to_m2(wall_area_sqft)
    for rule in ruleset_data["rules"]:
        # element filter
        if rule["element"] != "Wall":
            continue

        # typology filter
        if not match_building_category(rule["building_category"], typology):
            continue

        # area filter
        if not match_area_condition(wall_area_m2, rule["area_condition"]):
            continue

        # extract limits
        code = rule["code_level"]
        u_val = rule["u_factor_limits"].get(climate_zone)

        if u_val is not None:
            limits[code] = u_val

    return limits

def get_roof_u_limits(data, ruleset_data, climate_zone, typology):
    limits = {}

    # Roof area in SQFT → convert to m2
    roof_area_sqft = float(data["ROOF-AREA(SQFT)"].iloc[0])
    roof_area_m2 = sqft_to_m2(roof_area_sqft)

    for rule in ruleset_data["rules"]:
        # element filter
        if rule["element"] != "Roof":
            continue

        # typology filter
        if not match_building_category(rule["building_category"], typology):
            continue

        # area filter
        if not match_area_condition(roof_area_m2, rule["area_condition"]):
            continue

        # extract limits
        code = rule["code_level"]
        u_val = rule["u_factor_limits"].get(climate_zone)

        if u_val is not None:
            limits[code] = u_val

    return limits

def energy_param_plot_wallGains(
    x_param_df,
    baseline_df,
    ecsbc_climate, typologies,
    x_col,
    x_label,
    title,   # kept for compatibility, not used
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
            lambda x: "High Density" if x >= 13 else "Low Density"
        )

    fig = go.Figure()

    # Force x column
    x_col = "R-VAL-W"
    baseline_x = baseline_df[x_col].iloc[0]

    # =========================
    # DataFrames
    # =========================
    low_df = (
        df[(df["DensityType"] == "Low Density") & (df[x_col] != baseline_x)]
        .sort_values(by=x_col)
    )

    high_df = (
        df[(df["DensityType"] == "High Density") & (df[x_col] != baseline_x)]
        .sort_values(by=x_col)
    )

    # =========================
    # Y-axis ranges (GLOBAL min-max)
    # =========================
    y_gain_vals = pd.concat([
        low_df["WALL CONDUCTION"],
        high_df["WALL CONDUCTION"],
        baseline_df["WALL CONDUCTION"]
    ])

    y_loss_vals = pd.concat([
        low_df["WALL CONDUCTION_loss"],
        high_df["WALL CONDUCTION_loss"],
        baseline_df["WALL CONDUCTION_loss"]
    ])

    y_gain_min, y_gain_max = y_gain_vals.min(), y_gain_vals.max()
    y_loss_min, y_loss_max = y_loss_vals.min(), y_loss_vals.max()

    gain_pad = (y_gain_max - y_gain_min) * 0.05
    loss_pad = (y_loss_max - y_loss_min) * 0.05

    # =========================
    # Low Density
    # =========================
    fig.add_trace(go.Scatter(
        x=low_df[x_col],
        y=low_df["WALL CONDUCTION"],
        mode="lines",
        line=dict(color="#B5B5B5", width=2, shape="spline", smoothing=1.2),
        name="Low Density",
        hovertemplate="<b>Low Density</b><br>R-Value: %{x:.2f}<br>Gain: %{y:.2f} kW<extra></extra>",
        yaxis="y"
    ))

    # =========================
    # High Density
    # =========================
    fig.add_trace(go.Scatter(
        x=high_df[x_col],
        y=high_df["WALL CONDUCTION"],
        mode="lines",
        line=dict(color="#5A5A5A", width=2, shape="spline", smoothing=1.2),
        name="High Density",
        hovertemplate="<b>High Density</b><br>R-Value: %{x:.2f}<br>Gain: %{y:.2f} kW<extra></extra>",
        yaxis="y"
    ))

    # =========================
    # As Designed
    # =========================
    fig.add_trace(go.Scatter(
        x=baseline_df[x_col],
        y=baseline_df["WALL CONDUCTION"],
        mode="markers",
        marker=dict(size=14, color="red"),
        name="As Designed",
        hovertemplate="<b>As Designed</b><br>R-Value: %{x:.2f}<br>Gain: %{y:.2f} kW<extra></extra>",
        yaxis="y"
    ))

    # =========================
    # Layout
    # =========================
    fig.update_layout(
        template="plotly_white",
        height=500,
        margin=dict(l=80, r=40, t=25, b=65),   # 🔥 reduced bottom space
        font=dict(size=13),

        # ✅ Legend: tight to plot bottom, no box
        legend=dict(
            orientation="h",
            yanchor="top",
            y=-0.008,     # 🔥 ultra-tight
            xanchor="center",
            x=0.5,
            font=dict(size=12)
        ),

        # Top (Gains)
        yaxis=dict(
            title="Gain (kW)",
            domain=[0.52, 1.0],
            range=[y_gain_min - gain_pad, y_gain_max + gain_pad],
            showgrid=True,
            zeroline=False
        ),

        showlegend=show_legend
    )

    # =========================
    # X-axis
    # =========================
    fig.update_xaxes(
        title=x_label,
        anchor="y",
        showgrid=True
    )

    return fig

def energy_param_plot_wallLoss(
    x_param_df,
    baseline_df,
    ecsbc_climate, typologies,
    x_col,
    x_label,
    title,
    show_legend=True):

    import pandas as pd
    import plotly.graph_objects as go

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
            lambda x: "High Density" if x >= 13 else "Low Density"
        )

    fig = go.Figure()

    # Force column
    x_col = "R-VAL-W"
    baseline_x = baseline_df[x_col].iloc[0]

    # =========================
    # DataFrames
    # =========================
    low_df = (
        df[(df["DensityType"] == "Low Density") & (df[x_col] != baseline_x)]
        .sort_values(by=x_col)
    )

    high_df = (
        df[(df["DensityType"] == "High Density") & (df[x_col] != baseline_x)]
        .sort_values(by=x_col)
    )

    # =========================
    # Y-axis ranges (GLOBAL min-max)
    # =========================
    y_loss_vals = pd.concat([
        low_df["WALL CONDUCTION_loss"],
        high_df["WALL CONDUCTION_loss"],
        baseline_df["WALL CONDUCTION_loss"]
    ])

    y_loss_min, y_loss_max = y_loss_vals.min(), y_loss_vals.max()
    loss_pad = (y_loss_max - y_loss_min) * 0.08

    # =========================
    # Low Density
    # =========================
    fig.add_trace(go.Scatter(
        x=low_df[x_col],
        y=low_df["WALL CONDUCTION_loss"],
        mode="lines",
        line=dict(color="#B5B5B5", width=2, shape="spline", smoothing=1.2),
        name="Low Density Loss",
        showlegend=True,
        hovertemplate="<b>Low Density</b><br>R-Value: %{x:.2f}<br>Loss: %{y:.2f} kW<extra></extra>",
        yaxis="y2"
    ))

    # =========================
    # High Density
    # =========================
    fig.add_trace(go.Scatter(
        x=high_df[x_col],
        y=high_df["WALL CONDUCTION_loss"],
        mode="lines",
        line=dict(color="#5A5A5A", width=2, shape="spline", smoothing=1.2),
        name="High Density Loss",
        showlegend=True,
        hovertemplate="<b>High Density</b><br>R-Value: %{x:.2f}<br>Loss: %{y:.2f} kW<extra></extra>",
        yaxis="y2"
    ))

    # =========================
    # As Designed
    # =========================
    fig.add_trace(go.Scatter(
        x=baseline_df[x_col],
        y=baseline_df["WALL CONDUCTION_loss"],
        mode="markers",
        marker=dict(size=14, color="red"),
        name="As Designed Loss",
        showlegend=True,
        hovertemplate="<b>As Designed</b><br>R-Value: %{x:.2f}<br>Loss: %{y:.2f} kW<extra></extra>",
        yaxis="y2"
    ))

    # =========================
    # Layout
    # =========================
    fig.update_layout(
        template="plotly_white",
        height=520,
        title="",
        margin=dict(l=80, r=40, t=60, b=90),
        font=dict(size=13),

        legend=dict(
            orientation="h",
            yanchor="top",
            y=-0.22,
            xanchor="center",
            x=0.5
        ),

        # Bottom (Loss axis)
        yaxis2=dict(
            domain=[0.0, 0.48],
            range=[y_loss_min - loss_pad, y_loss_max + loss_pad],
            showgrid=True,
            zeroline=False,
            title="Loss (kW)"
        ),

        showlegend=show_legend
    )

    # X axis
    fig.update_xaxes(
        title=x_label,
        anchor="y2",
        showgrid=True
    )

    return fig

def energy_param_plot_wall(
    x_param_df,
    baseline_df,
    ecsbc_climate, typologies,
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

    # =========================
    # ECSBC Vertical Lines
    # =========================
    json_path = "rulesets/ecsbc_envelope_ruleset.json"
    if os.path.exists(json_path):
        with open(json_path, "r", encoding="utf-8") as f:
            ruleset_data = json.load(f)
        # change climate zone as needed
        climate_zone = ecsbc_climate
        wall_limits = get_wall_u_limits(x_param_df, ruleset_data, climate_zone, typologies)
        line_styles = {
            "ECSBC": dict(color="#F36608", dash="dash"),      # Dark orange
            "ECSBC+": dict(color="#E08B45", dash="dash"),     # Medium-light orange
            "SUPER ECSBC": dict(color="#F7C59F", dash="dash") # Very light orange
        }

        from collections import defaultdict

        tolerance = 0.02   # adjust if needed
        grouped = defaultdict(list)

        # -------- Group close R-values --------
        for code_level, u_val in wall_limits.items():
            r_val = round(u_to_r(u_val), 3)
            placed = False
            for k in grouped:
                if abs(k - r_val) <= tolerance:
                    grouped[k].append((code_level, r_val))
                    placed = True
                    break
            if not placed:
                grouped[r_val].append((code_level, r_val))

        # -------- Plot with offsets --------
        offset_step = 0.015

        for base_r, items in grouped.items():
            n = len(items)
            for i, (code_level, r_val) in enumerate(items):
                offset = (i - (n - 1) / 2) * offset_step
                x_pos = base_r + offset

                fig.add_vline(
                    x=x_pos,
                    line_width=2,
                    line_dash=line_styles[code_level]["dash"],
                    line_color=line_styles[code_level]["color"],
                    annotation_text=f"{code_level}<br>R={r_val:.2f}",
                    annotation_position="top",
                    annotation_font_size=11,
                    annotation_yshift=12 * i
                )

    return fig

def energy_param_plot_roof_loss(
    x_param_df,
    baseline_df,
    ecsbc_climate, typologies,
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
            y=mid_df["ROOF CONDUCTION_loss"],
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
                "<b>Roof Conduction</b>: %{y:,.0f} kW<br>"
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
            y=high_df["ROOF CONDUCTION_loss"],
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
                "<b>Roof Conduction</b>: %{y:,.0f} kW<br>"
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
            y=baseline_df["ROOF CONDUCTION_loss"],
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
                "<b>Roof Conduction</b>: %{y:,.0f} kW<br>"
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
    fig.update_yaxes(title="Losses (kW)", tickformat=".4s")

    return fig

def energy_param_plot_roof_gains(
    x_param_df,
    baseline_df,
    ecsbc_climate, typologies,
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
            y=mid_df["ROOF CONDUCTION"],
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
                "<b>Roof Conduction</b>: %{y:,.0f} kW<br>"
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
            y=high_df["ROOF CONDUCTION"],
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
                "<b>Roof Conduction</b>: %{y:,.0f} kW<br>"
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
            y=baseline_df["ROOF CONDUCTION"],
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
                "<b>Roof Conduction</b>: %{y:,.0f} kW<br>"
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
    fig.update_yaxes(title="Gains (kW)", tickformat=".4s")

    return fig

def energy_param_plot_wall_gains(
    x_param_df,
    baseline_df,
    ecsbc_climate, typologies,
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
            y=mid_df["WALL CONDUCTION"],
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
                "<b>Wall Conduction</b>: %{y:,.0f} kW<br>"
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
            y=high_df["WALL CONDUCTION"],
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
                "<b>Wall Conduction</b>: %{y:,.0f} kW<br>"
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
            y=baseline_df["WALL CONDUCTION"],
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
                "<b>Wall Conduction</b>: %{y:,.0f} kW<br>"
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
    fig.update_yaxes(title="Gains (kW)", tickformat=".4s")

    return fig

def energy_param_plot_solar_loss(
    x_param_df,
    baseline_df,
    ecsbc_climate, typologies,
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
            y=mid_df["WINDOW GLASS+FRM COND_loss"],
            mode="lines",
            line=dict(
                color="#B0B0B0",
                width=2,
                shape="spline",
                smoothing=1.2
            ),
            name="Parametric",
            hovertemplate=(
                f"<b>{x_label}</b>: %{{x:.2f}}<br>"
                "<b>Window Solar</b>: %{y:,.0f} kW<br>"
                "<b>Type</b>: Solar Parameter<extra></extra>"
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
            y=high_df["WINDOW GLASS+FRM COND_loss"],
            mode="lines",
            line=dict(
                color="#6B6B6B",
                width=2,
                shape="spline",
                smoothing=1.2
            ),
            name="Parametric",
            hovertemplate=(
                f"<b>{x_label}</b>: %{{x:.2f}}<br>"
                "<b>Window Solar</b>: %{y:,.0f} kW<br>"
                "<b>Type</b>: Solar Parameter<extra></extra>"
            )
        )
    )

    # =========================
    # As Designed (only point kept)
    # =========================
    fig.add_trace(
        go.Scatter(
            x=baseline_df[x_col],
            y=baseline_df["WINDOW GLASS+FRM COND_loss"],
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
                "<b>Window Solar</b>: %{y:,.0f} kW<br>"
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
    fig.update_yaxes(title="Losses (kW)", tickformat=".4s")

    return fig

def energy_param_plot_window_loss(
    x_param_df,
    baseline_df,
    ecsbc_climate, typologies,
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
            y=mid_df["WINDOW GLASS+FRM COND_loss"],
            mode="lines",
            line=dict(
                color="#B0B0B0",
                width=2,
                shape="spline",
                smoothing=1.2
            ),
            name="Parametric",
            hovertemplate=(
                f"<b>{x_label}</b>: %{{x:.2f}}<br>"
                "<b>Window Conduction</b>: %{y:,.0f} kW<br>"
                "<b>Type</b>: Window Parameter<extra></extra>"
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
            y=high_df["WINDOW GLASS+FRM COND"],
            mode="lines",
            line=dict(
                color="#6B6B6B",
                width=2,
                shape="spline",
                smoothing=1.2
            ),
            name="Parametric",
            hovertemplate=(
                f"<b>{x_label}</b>: %{{x:.2f}}<br>"
                "<b>Window Conduction</b>: %{y:,.0f} kW<br>"
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
            y=baseline_df["WINDOW GLASS+FRM COND"],
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
                "<b>Window Conduction</b>: %{y:,.0f} kW<br>"
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
    fig.update_yaxes(title="Losses (kW)", tickformat=".4s")

    return fig

def energy_param_plot_window_gains(
    x_param_df,
    baseline_df,
    ecsbc_climate, typologies,
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
            y=mid_df["WINDOW GLASS+FRM COND"],
            mode="lines",
            line=dict(
                color="#B0B0B0",
                width=2,
                shape="spline",
                smoothing=1.2
            ),
            name="Parametric",
            hovertemplate=(
                f"<b>{x_label}</b>: %{{x:.2f}}<br>"
                "<b>Window Conduction</b>: %{y:,.0f} kW<br>"
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
            y=high_df["WINDOW GLASS+FRM COND"],
            mode="lines",
            line=dict(
                color="#6B6B6B",
                width=2,
                shape="spline",
                smoothing=1.2
            ),
            name="Parametric",
            hovertemplate=(
                f"<b>{x_label}</b>: %{{x:.2f}}<br>"
                "<b>Window Conduction</b>: %{y:,.0f} kW<br>"
                "<b>Type</b>: Window Paramter<extra></extra>"
            )
        )
    )

    # =========================
    # As Designed (only point kept)
    # =========================
    fig.add_trace(
        go.Scatter(
            x=baseline_df[x_col],
            y=baseline_df["WINDOW GLASS+FRM COND"],
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
                "<b>Window Conduction</b>: %{y:,.0f} kW<br>"
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
    fig.update_yaxes(title="Gains (kW)", tickformat=".4s")

    return fig

def energy_param_plot_solar_gains(
    x_param_df,
    baseline_df,
    ecsbc_climate, typologies,
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
            y=mid_df["WINDOW GLASS SOLAR"],
            mode="lines",
            line=dict(
                color="#B0B0B0",
                width=2,
                shape="spline",
                smoothing=1.2
            ),
            name="Parametric",
            hovertemplate=(
                f"<b>{x_label}</b>: %{{x:.2f}}<br>"
                "<b>Window Solar</b>: %{y:,.0f} kW<br>"
                "<b>Type</b>: Solar Parameter<extra></extra>"
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
            y=high_df["WINDOW GLASS SOLAR"],
            mode="lines",
            line=dict(
                color="#6B6B6B",
                width=2,
                shape="spline",
                smoothing=1.2
            ),
            name="Parametric",
            hovertemplate=(
                f"<b>{x_label}</b>: %{{x:.2f}}<br>"
                "<b>Window Solar</b>: %{y:,.0f} kW<br>"
                "<b>Type</b>: Solar Parameter<extra></extra>"
            )
        )
    )

    # =========================
    # As Designed (only point kept)
    # =========================
    fig.add_trace(
        go.Scatter(
            x=baseline_df[x_col],
            y=baseline_df["WINDOW GLASS SOLAR"],
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
                "<b>Window Solar</b>: %{y:,.0f} kW<br>"
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
    fig.update_yaxes(title="Gains (kW)", tickformat=".4s")

    return fig

def energy_param_plot_wall_loss(
    x_param_df,
    baseline_df,
    ecsbc_climate, typologies,
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
            y=mid_df["WALL CONDUCTION_loss"],
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
                "<b>Wall Conduction</b>: %{y:,.0f} kW<br>"
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
            y=high_df["WALL CONDUCTION_loss"],
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
                "<b>Wall Conduction</b>: %{y:,.0f} kW<br>"
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
            y=baseline_df["WALL CONDUCTION_loss"],
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
                "<b>Wall Conduction</b>: %{y:,.0f} kW<br>"
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
    fig.update_yaxes(title="Losses (kW)", tickformat=".4s")

    return fig

# ------------ Graph ----------- #
def energy_param_plot(
    x_param_df,
    baseline_df,
    ecsbc_climate, typologies,
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

    # =========================
    # ECSBC Vertical Lines
    # =========================
    json_path = "rulesets/ecsbc_envelope_ruleset.json"
    if os.path.exists(json_path):
        with open(json_path, "r", encoding="utf-8") as f:
            ruleset_data = json.load(f)

        # change climate zone as needed
        climate_zone = ecsbc_climate
        roof_limits = get_roof_u_limits(x_param_df, ruleset_data, climate_zone, typologies)

        line_styles = {
            "ECSBC": dict(color="#F36608", dash="dash"),      # Dark orange
            "ECSBC+": dict(color="#E08B45", dash="dash"),     # Medium-light orange
            "SUPER ECSBC": dict(color="#F7C59F", dash="dash") # Very light orange
        }

        for code_level, u_val in roof_limits.items():
            r_val = u_to_r(u_val)
            fig.add_vline(
                x=r_val,
                line_width=2,
                line_dash=line_styles[code_level]["dash"],
                line_color=line_styles[code_level]["color"],
                annotation_text=f"{code_level}<br>R={r_val:.2f}",
                annotation_position="top",
                annotation_font_size=11
            )

    return fig

def energy_param_plot_roofGains(
    x_param_df,
    baseline_df,
    ecsbc_climate, typologies,
    x_col,
    x_label,
    title,   # kept for compatibility but not used
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
            lambda x: "High Density" if x >= 13 else "Low Density"
        )

    fig = go.Figure()

    # Force x column
    x_col = "R-VAL-R"
    baseline_x = baseline_df[x_col].iloc[0]

    # =========================
    # DataFrames
    # =========================
    low_df = (
        df[(df["DensityType"] == "Low Density") & (df[x_col] != baseline_x)]
        .sort_values(by=x_col)
    )

    high_df = (
        df[(df["DensityType"] == "High Density") & (df[x_col] != baseline_x)]
        .sort_values(by=x_col)
    )

    # =========================
    # Y-axis ranges (GLOBAL min-max)
    # =========================
    y_gain_vals = pd.concat([
        low_df["ROOF CONDUCTION"],
        high_df["ROOF CONDUCTION"],
        baseline_df["ROOF CONDUCTION"]
    ])

    y_loss_vals = pd.concat([
        low_df["ROOF CONDUCTION_loss"],
        high_df["ROOF CONDUCTION_loss"],
        baseline_df["ROOF CONDUCTION_loss"]
    ])

    y_gain_min, y_gain_max = y_gain_vals.min(), y_gain_vals.max()
    y_loss_min, y_loss_max = y_loss_vals.min(), y_loss_vals.max()

    gain_pad = (y_gain_max - y_gain_min) * 0.05
    loss_pad = (y_loss_max - y_loss_min) * 0.05

    # =========================
    # Low Density
    # =========================
    fig.add_trace(go.Scatter(
        x=low_df[x_col],
        y=low_df["ROOF CONDUCTION"],
        mode="lines",
        line=dict(color="#B5B5B5", width=2, shape="spline", smoothing=1.2),
        name="Low Density",
        hovertemplate="<b>Low Density</b><br>R-Value: %{x:.2f}<br>Gain: %{y:.2f} kW<extra></extra>",
        yaxis="y"
    ))

    # =========================
    # High Density
    # =========================
    fig.add_trace(go.Scatter(
        x=high_df[x_col],
        y=high_df["ROOF CONDUCTION"],
        mode="lines",
        line=dict(color="#5A5A5A", width=2, shape="spline", smoothing=1.2),
        name="High Density",
        hovertemplate="<b>High Density</b><br>R-Value: %{x:.2f}<br>Gain: %{y:.2f} kW<extra></extra>",
        yaxis="y"
    ))

    # =========================
    # As Designed
    # =========================
    fig.add_trace(go.Scatter(
        x=baseline_df[x_col],
        y=baseline_df["ROOF CONDUCTION"],
        mode="markers",
        marker=dict(size=14, color="red"),
        name="As Designed",
        hovertemplate="<b>As Designed</b><br>R-Value: %{x:.2f}<br>Gain: %{y:.2f} kW<extra></extra>",
        yaxis="y"
    ))

    # =========================
    # Layout
    # =========================
    fig.update_layout(
        template="plotly_white",
        height=500,
        margin=dict(l=80, r=40, t=25, b=65),   # 🔥 tight bottom margin
        font=dict(size=13),

        # ✅ Legend: very close to plot
        legend=dict(
            orientation="h",
            yanchor="top",
            y=-0.15,      # 🔥 ultra-tight
            xanchor="center",
            x=0.5,
            font=dict(size=12)
        ),

        # Gains axis
        yaxis=dict(
            title="Gain (kW)",
            domain=[0.52, 1.0],
            range=[y_gain_min - gain_pad, y_gain_max + gain_pad],
            showgrid=True,
            zeroline=False
        ),

        showlegend=show_legend
    )

    # =========================
    # X-axis
    # =========================
    fig.update_xaxes(
        title=x_label,
        anchor="y",
        showgrid=True
    )

    return fig

def energy_param_plot_roofLoss(
    x_param_df,
    baseline_df,
    ecsbc_climate, typologies,
    x_col,
    x_label,
    title,
    show_legend=True):

    import pandas as pd
    import plotly.graph_objects as go

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
            lambda x: "High Density" if x >= 13 else "Low Density"
        )

    fig = go.Figure()

    # Force column
    x_col = "R-VAL-R"
    baseline_x = baseline_df[x_col].iloc[0]

    # =========================
    # DataFrames
    # =========================
    low_df = (
        df[(df["DensityType"] == "Low Density") & (df[x_col] != baseline_x)]
        .sort_values(by=x_col)
    )

    high_df = (
        df[(df["DensityType"] == "High Density") & (df[x_col] != baseline_x)]
        .sort_values(by=x_col)
    )

    # =========================
    # Y-axis ranges (GLOBAL min-max)
    # =========================
    y_loss_vals = pd.concat([
        low_df["ROOF CONDUCTION_loss"],
        high_df["ROOF CONDUCTION_loss"],
        baseline_df["ROOF CONDUCTION_loss"]
    ])

    y_loss_min, y_loss_max = y_loss_vals.min(), y_loss_vals.max()
    loss_pad = (y_loss_max - y_loss_min) * 0.08

    # =========================
    # Low Density
    # =========================
    fig.add_trace(go.Scatter(
        x=low_df[x_col],
        y=low_df["ROOF CONDUCTION_loss"],
        mode="lines",
        line=dict(color="#B5B5B5", width=2, shape="spline", smoothing=1.2),
        name="Low Density Loss",
        showlegend=True,
        hovertemplate="<b>Low Density</b><br>R-Value: %{x:.2f}<br>Loss: %{y:.2f} kW<extra></extra>",
        yaxis="y2"
    ))

    # =========================
    # High Density
    # =========================
    fig.add_trace(go.Scatter(
        x=high_df[x_col],
        y=high_df["ROOF CONDUCTION_loss"],
        mode="lines",
        line=dict(color="#5A5A5A", width=2, shape="spline", smoothing=1.2),
        name="High Density Loss",
        showlegend=True,
        hovertemplate="<b>High Density</b><br>R-Value: %{x:.2f}<br>Loss: %{y:.2f} kW<extra></extra>",
        yaxis="y2"
    ))

    # =========================
    # As Designed
    # =========================
    fig.add_trace(go.Scatter(
        x=baseline_df[x_col],
        y=baseline_df["ROOF CONDUCTION_loss"],
        mode="markers",
        marker=dict(size=14, color="red"),
        name="As Designed Loss",
        showlegend=True,
        hovertemplate="<b>As Designed</b><br>R-Value: %{x:.2f}<br>Loss: %{y:.2f} kW<extra></extra>",
        yaxis="y2"
    ))

    # =========================
    # Layout
    # =========================
    fig.update_layout(
        template="plotly_white",
        height=520,
        title="",
        margin=dict(l=80, r=40, t=60, b=90),
        font=dict(size=13),

        legend=dict(
            orientation="h",
            yanchor="top",
            y=-0.22,
            xanchor="center",
            x=0.5
        ),

        # Bottom (Loss axis)
        yaxis2=dict(
            title="Loss (kW)",
            domain=[0.0, 0.48],
            range=[y_loss_min - loss_pad, y_loss_max + loss_pad],
            showgrid=True,
            zeroline=False
        ),

        showlegend=show_legend
    )

    # X axis
    fig.update_xaxes(
        title=x_label,
        anchor="y2",
        showgrid=True
    )

    return fig

def energy_param_plot_wind(
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

def get_bar_color(row):
    # Rule 1: FileName starts with 0 → red
    if str(row["FileName"]).startswith("0_"):
        return "#D32F2F"   # red
    
    # Rule 2: DensityType based
    if row["DensityType"] == "High Density":
        return "#6B6B6B"
    elif row["DensityType"] == "Low Density":
        return "#B0B0B0"
    
    # fallback
    return "#999999"

def get_bar_color_roof(row):
    # Rule 1: FileName starts with 0 → red
    if str(row["FileName"]).startswith("0_0_"):
        return "#D32F2F"   # red
    
    # Rule 2: DensityType based
    if row["DensityType"] == "High Density":
        return "#6B6B6B"
    elif row["DensityType"] == "Low Density":
        return "#B0B0B0"
    
    # fallback
    return "#999999"

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
        xaxis_title=x_label,
        yaxis_title="Energy Use (kWh)",
        yaxis=dict(tickformat=".1s"),
        legend=dict(orientation="h", y=-0.3, x=0.5, xanchor="center"),
        height=450
    )
    fig.update_yaxes(title="Energy Use (kWh)", tickformat=".4s")
    return fig

def make_single_plot_shgc(df, x_col, title, x_label):
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
            f"<b>{x_label}</b>: %{{x:.2f}}<br>"
            "<b>Energy</b>: %{customdata[0]}<extra></extra>"
        )
    ))

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
            f"<b>{x_label}</b>: %{{x:.2f}}<br>"
            "<b>Energy</b>: %{customdata[0]}<extra></extra>"
        )
    ))

    fig.update_layout(
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

def save_figs_as_images_2(figs, folder="pdf_charts"):
    import os
    import shutil
    import plotly.graph_objects as go

    if os.path.exists(folder):
        shutil.rmtree(folder)
    os.makedirs(folder)

    paths = []

    for i, fig in enumerate(figs):

        # Clone the figure so Streamlit version is untouched
        fig_pdf = go.Figure(fig)

        fig_pdf.update_layout(
            width=1400,
            height=750,
            margin=dict(l=60, r=60, t=80, b=160),
            template="none"
        )

        path = f"{folder}/chart_{i+1}.png"

        fig_pdf.write_image(path, scale=4)

        paths.append(path)

    return paths

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

from reportlab.platypus import PageBreak, ListFlowable, ListItem, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle

def add_design_parameter_tables(elements, styles):
    from reportlab.platypus import Table, TableStyle, Paragraph, Spacer
    from reportlab.lib import colors
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import inch

    TOTAL_WIDTH = 7.4 * inch   # 🔒 fixed page width

    title_style = ParagraphStyle(
        name="ParamTitle",
        fontName="Helvetica-Bold",
        fontSize=11,
        spaceAfter=8
    )

    bold = ParagraphStyle(
        name="ParamBold",
        fontName="Helvetica-Bold",
        fontSize=9,
        leading=12
    )

    # -------------------------------
    # SECTION TITLE
    # -------------------------------
    elements.append(Paragraph("Design Parameter Ranges", title_style))
    elements.append(Spacer(1, 0.15 * inch))

    # ==========================================================
    # WALL TABLE
    # ==========================================================
    wall_data = [
        ["Category", "Type", "Layer Sequence (External → Internal)", "Insulation", "R-Values"],
        ["Wall", "High Density",
         "Cement Plaster → XPS → Brick → Cement Plaster",
         "XPS",
         "2.5, 5, 7.5, 10, 15, 20, 25, 30"],
        ["Wall", "Low Density",
         "Cement Board → XPS → Gypsum Board",
         "XPS",
         "2.5, 5, 7.5, 10, 15, 20, 25, 30"],
    ]

    wall_colwidths = [
        0.7*inch,  # Category
        0.9*inch,  # Type
        3.3*inch,  # Layers
        0.7*inch,  # Insulation
        1.8*inch,  # R-values
    ]  # = 7.4 inch

    wall_table = Table(wall_data, colWidths=wall_colwidths)
    wall_table.setStyle([
        ("BOX", (0,0), (-1,-1), 0.5, colors.grey),
        ("INNERGRID", (0,0), (-1,-1), 0.25, colors.lightgrey),
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#F2F4F7")),   # header
        ("BACKGROUND", (0,1), (-1,-1), colors.HexColor("#FAFBFD")), # body
        ("FONT", (0,0), (-1,0), "Helvetica-Bold", 9),
        ("FONT", (0,1), (-1,-1), "Helvetica", 9),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("LEFTPADDING", (0,0), (-1,-1), 6),
        ("RIGHTPADDING", (0,0), (-1,-1), 6),
        ("TOPPADDING", (0,0), (-1,-1), 5),
        ("BOTTOMPADDING", (0,0), (-1,-1), 5),
    ])

    elements.append(Paragraph("Wall Assemblies", bold))
    elements.append(wall_table)
    elements.append(Spacer(1, 0.22 * inch))

    # ==========================================================
    # ROOF TABLE
    # ==========================================================
    roof_data = [
        ["Category", "Type", "Layer Sequence (External → Internal)", "Insulation", "R-Values"],
        ["Roof", "High Density",
         "Cement Plaster → XPS → Brick → Concrete",
         "XPS", "2.5, 5, 7.5, 10, 15, 20, 25, 30"],
        ["Roof", "Low Density",
         "Metal Deck → XPS",
         "XPS",
         "2.5, 5, 7.5, 10, 15, 20, 25, 30"],
    ]
    roof_colwidths = [
        0.7*inch,  # Category
        0.9*inch,  # Type
        3.3*inch,  # Layers
        0.7*inch,  # Insulation
        1.8*inch,  # R-values
    ]

    roof_table = Table(roof_data, colWidths=roof_colwidths)
    roof_table.setStyle([
        ("BOX", (0,0), (-1,-1), 0.5, colors.grey),
        ("INNERGRID", (0,0), (-1,-1), 0.25, colors.lightgrey),
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#F1F7F3")),
        ("BACKGROUND", (0,1), (-1,-1), colors.HexColor("#FBFDFC")),
        ("FONT", (0,0), (-1,0), "Helvetica-Bold", 9),
        ("FONT", (0,1), (-1,-1), "Helvetica", 9),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("LEFTPADDING", (0,0), (-1,-1), 6),
        ("RIGHTPADDING", (0,0), (-1,-1), 6),
        ("TOPPADDING", (0,0), (-1,-1), 5),
        ("BOTTOMPADDING", (0,0), (-1,-1), 5),
    ])

    elements.append(Paragraph("Roof Assemblies", bold))
    elements.append(roof_table)
    elements.append(Spacer(1, 0.22 * inch))

    # ==========================================================
    # LOADS TABLE
    # ==========================================================
    loads_data = [
        ["Category", "Parameter", "Full Form", "Range", "Step"],
        ["Internal Load", "LPD", "Lighting Power Density", "0.3 → 1 W/ft²", "+0.1 W/ft²"],
        ["Internal Load", "EPD", "Equipment Power Density", "1 → 2 W/ft²", "+0.1 W/ft²"]
    ]

    loads_colwidths = [
        1.3*inch,  # Category
        0.9*inch,  # Param
        2.4*inch,  # Full form
        1.4*inch,  # Range
        1.4*inch   # Step
    ]  # = 7.4 inch

    loads_table = Table(loads_data, colWidths=loads_colwidths)
    loads_table.setStyle([
        ("BOX", (0,0), (-1,-1), 0.5, colors.grey),
        ("INNERGRID", (0,0), (-1,-1), 0.25, colors.lightgrey),
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#F8F6EE")),
        ("BACKGROUND", (0,1), (-1,-1), colors.HexColor("#FDFCF7")),
        ("FONT", (0,0), (-1,0), "Helvetica-Bold", 9),
        ("FONT", (0,1), (-1,-1), "Helvetica", 9),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("LEFTPADDING", (0,0), (-1,-1), 6),
        ("RIGHTPADDING", (0,0), (-1,-1), 6),
        ("TOPPADDING", (0,0), (-1,-1), 5),
        ("BOTTOMPADDING", (0,0), (-1,-1), 5),
    ])

    elements.append(Paragraph("Internal Loads", bold))
    elements.append(loads_table)
    elements.append(Spacer(1, 0.22 * inch))

    # ==========================================================
    # GLAZING TABLE
    # ==========================================================
    glazing_data = [
        ["Parameter", "Range", "Step"],
        # ["U-Value", "0.8 → 3.5 W/m²K", "Stepwise"],
        ["SHGC", "0.2 → 0.9", "+0.1"],
        ["WWR", "20% → 80%", "+10%"],
    ]

    glazing_colwidths = [
        1.8*inch,
        3.8*inch,
        1.8*inch
    ]  # = 7.4 inch

    glazing_table = Table(glazing_data, colWidths=glazing_colwidths)
    glazing_table.setStyle([
        ("BOX", (0,0), (-1,-1), 0.5, colors.grey),
        ("INNERGRID", (0,0), (-1,-1), 0.25, colors.lightgrey),
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#F4F1F7")),
        ("BACKGROUND", (0,1), (-1,-1), colors.HexColor("#FBFAFD")),
        ("FONT", (0,0), (-1,0), "Helvetica-Bold", 9),
        ("FONT", (0,1), (-1,-1), "Helvetica", 9),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("LEFTPADDING", (0,0), (-1,-1), 6),
        ("RIGHTPADDING", (0,0), (-1,-1), 6),
        ("TOPPADDING", (0,0), (-1,-1), 5),
        ("BOTTOMPADDING", (0,0), (-1,-1), 5),
    ])

    elements.append(Paragraph("Glazing Parameters", bold))
    elements.append(glazing_table)
    elements.append(Spacer(1, 0.25 * inch))

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
        ["Climate Zone",f"{project_info.get('climate_zone', '—')} (ASHRAE), {values.get('ecsbcZone')} (NBC)"],
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
        ("Area (ft²)", f"{values.get('area', 0):,.0f}" if values.get("area") else 0),
        ("Conditioned Area (%)", f"{values.get('condArea', 0):.1f}" if values.get("condArea") else 0),
        ("Window-to-Wall Ratio (%)", f"{values.get('wwr', 0):.1f}" if values.get("wwr") else 0),
        ("Wall-to-Floor Ratio (%)", f"{values.get('wfr', 0):.1f}" if values.get("wfr") else 0),
        ("Above-Grade Area (%)", f"{values.get('agArea', 0):,.1f}" if values.get("agArea") else 0),
        ("Envelope-to-Floor Ratio (%)", f"{values.get('envelopeFloorArea', 0):.1f}" if values.get("envelopeFloorArea") else 0),
        ("Estimated Hours of Use", f"{values.get('estimateHrsUse', 0):,.0f}" if values.get("estimateHrsUse") else 0),
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

    elements.append(
        Paragraph(
            "<b><font color='black'>Energy Use</font></b>",
            ParagraphStyle(
                name="EnergyUse",
                fontSize=11,
                spaceAfter=8
            )
        )
    )

    # -------------------------------
    # CHART GRID (2 PER ROW)
    # -------------------------------
    row = []
    table_data = []
    countImage = 0
    heading_inserted = False

    for img_path in image_paths:
        img = Image(img_path, width=3.8 * inch, height=2.4 * inch)
        row.append(img)
        countImage += 1

        # 2 images per row
        if len(row) == 2:
            table_data.append(row)
            row = []

        # When 3 rows (6 images) completed
        if countImage == 6 and not heading_inserted:
            # flush current images
            if table_data:
                elements.append(Table(table_data, colWidths=[3.7 * inch, 3.7 * inch]))
                elements.append(Spacer(1, 0.35 * inch))
                table_data = []

            # insert heading
            elements.append(
                Paragraph(
                    "<b><font color='black'>Gains and Losses</font></b>",
                    ParagraphStyle(
                        name="AboutTitle",
                        fontSize=11,
                        spaceAfter=8
                    )
                )
            )
            heading_inserted = True  # prevent repeat insert
            heading_inserted = False
        
    if row:
        table_data.append(row)

    if table_data:
        elements.append(Table(table_data, colWidths=[3.7 * inch, 3.7 * inch]))
    elements.append(Spacer(1, 0.35 * inch))
    elements.append(PageBreak())
    add_design_parameter_tables(elements, styles)
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
            "through systematic investigations in just a few clicks. "
            "These investigations are accessible through interactive charts and "
            "downloadable reports. Please feel free to share your feedback with us at "
            "<u>info@edsglobal.com</u>.",
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
            "The team focuses on climate change mitigation, low-carbon design, building simulation, performance audits, and capacity building. "
            "EDS has synthesized its years of experience in these domains by developing IT applications to support these endeavours. EDS continues to contribute to the buildings community with useful tools through its IT services. ",
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
            "AWESIM (referred to as the “Application” hereafter) is an outcome of the best "
            "efforts of building simulation experts and IT developers at "
            "<b>Environmental Design Solutions Pvt. Ltd. (EDS).</b>"
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

    pdf.build(
        elements,
        onFirstPage=lambda c, d: add_header_footer(c, d, title, logo_path, awesim_logo),
        onLaterPages=add_later_pages
    )

def create_com_pdf(image_paths, project_info, values, pdf_name="Energy_Parametric_Report.pdf"):
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
        ["Climate Zone",f"{project_info.get('climate_zone', '—')} (ASHRAE), {values.get('ecsbcZone')} (NBC)"],
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

    parameters = [
        ("Wall R-Value(ft²·°F·h/BTU)", "r_wall"),
        ("Roof R-Value(ft²·°F·h/BTU)", "r_roof"),
        ("Window U-Value(BTU/ft²·°F·h)", "r_wind"),
        ("Shading Coefficient", "shgc"),
        ("Window-to-Wall Ratio (%)", "wwr"),
        ("Lighting (W/ft²)", "light"),
        ("Equipment (W/ft²)", "equip"),
        ("Energy Outcome (kWh)", "ener"),
    ]

    table_data = [
        [
            Paragraph("<b>Parameters</b>", kpi_bold),
            Paragraph("<b>As Designed</b>", kpi_bold),
            Paragraph("<b>User Selected</b>", kpi_bold),
        ]
    ]

    for label, key in parameters:
        v0 = values.get(f"{key}_0", 0)
        v1 = values.get(f"{key}_1", 0)

        if key == "ener":
            v0 = f"{v0:,.0f}" if v0 else 0
            v1 = f"{v1:,.0f}" if v1 else 0
        elif key in ["r_wall", "r_roof", "r_wind"]:
            v0 = f"{v0:.2f}" if v0 else 0
            v1 = f"{v1:.2f}" if v1 else 0
        else:
            v0 = f"{v0:.1f}" if v0 else 0
            v1 = f"{v1:.1f}" if v1 else 0

        table_data.append([
            Paragraph(label, kpi_bold),
            Paragraph(str(v0), kpi_normal),
            Paragraph(str(v1), kpi_normal),
        ])

    kpi_table = Table(
        table_data,
        colWidths=[
            3.4 * inch,
            2 * inch,
            2 * inch
        ]
    )

    kpi_table.setStyle([
        ("BOX", (0, 0), (-1, -1), 0.5, colors.grey),
        ("INNERGRID", (0, 0), (-1, -1), 0.25, colors.lightgrey),
        ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ])

    elements.append(kpi_table)
    elements.append(Spacer(1, 0.35 * inch))

    # -------------------------------
    # Energy Use Distribution
    # -------------------------------
    elements.append(
        Paragraph(
            "<b>Energy Use Distribution</b>",
            ParagraphStyle(name="KpiTitle", fontSize=10, spaceAfter=8)
        )
    )

    img = Image(image_paths[0], width=7.4 * inch, height=4.2 * inch)
    elements.append(img)
    elements.append(Spacer(1, 0.35 * inch))
    elements.append(PageBreak())

    # -------------------------------
    # Energy Use Comparison
    # -------------------------------
    elements.append(
        Paragraph(
            "<b>Energy Use Comparison</b>",
            ParagraphStyle(name="KpiTitle", fontSize=10, spaceAfter=8)
        )
    )

    img1 = Image(image_paths[1], width=3.65 * inch, height=2.8 * inch)
    img2 = Image(image_paths[2], width=3.65 * inch, height=2.8 * inch)

    table = Table([[img1, img2]], colWidths=[3.7 * inch, 3.7 * inch])
    elements.append(table)
    elements.append(Spacer(1, 0.35 * inch))

    # -------------------------------
    # Heating Gains Summary
    # -------------------------------
    elements.append(
        Paragraph(
            "<b>Gains Summary</b>",
            ParagraphStyle(name="KpiTitle", fontSize=10, spaceAfter=8)
        )
    )

    img = Image(image_paths[3], width=7.4 * inch, height=4.2* inch)
    elements.append(img)
    elements.append(Spacer(1, 0.35 * inch))
    elements.append(PageBreak())

    # -------------------------------
    # Losses Summary
    # -------------------------------
    elements.append(
        Paragraph(
            "<b>Losses Summary</b>",
            ParagraphStyle(name="KpiTitle", fontSize=10, spaceAfter=8)
        )
    )

    img = Image(image_paths[4], width=7.4 * inch, height=4.2 * inch)
    elements.append(img)
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
            "through systematic investigations in just a few clicks. "
            "These investigations are accessible through interactive charts and "
            "downloadable reports. Please feel free to share your feedback with us at "
            "<u>info@edsglobal.com</u>.",
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
            "The team focuses on climate change mitigation, low-carbon design, building simulation, performance audits, and capacity building. "
            "EDS has synthesized its years of experience in these domains by developing IT applications to support these endeavours. EDS continues to contribute to the buildings community with useful tools through its IT services. ",
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
            "AWESIM (referred to as the “Application” hereafter) is an outcome of the best "
            "efforts of building simulation experts and IT developers at "
            "<b>Environmental Design Solutions Pvt. Ltd. (EDS).</b>"
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

    pdf.build(
        elements,
        onFirstPage=lambda c, d: add_header_footer(c, d, title, logo_path, awesim_logo),
        onLaterPages=add_later_pages
    )

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
# Reset dropdown BEFORE widget is created
if st.session_state.reset_tools:
    st.session_state.tools_dropdown = "Select"
    st.session_state.reset_tools = False

col1, col2, col3, col5 = st.columns([0.16,0.16,3.4,0.6])
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
        st.write("""AWESIM enables an energy simulation expert to discover optimal design parameters through systematic investigations in just a few clicks. Please share your feedback with us.
        These investigations are accessible through interactive charts and downloadable reports. Please feel free to share your feedback with us at **info@edsglobal.com**
        """)
    with col2:
        st.markdown('<span style="font-size:24px; font-weight:600;">Parametric Analysis</span>',unsafe_allow_html=True)
        st.write("""AWESIM enables systematic parametric investigations with ease. Interactive charts and downloadable reports provide clear insights. 
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
        The team focuses on climate change mitigation, low-carbon design, building simulation, performance audits, and capacity building.
        "EDS has synthesized its years of experience in these domains by developing IT applications to support these endeavours. EDS continues to contribute to the buildings community with useful tools through its IT services.""")
        st.markdown("</div>", unsafe_allow_html=True)
    # --------- Card 2: PARSim ----------
    with col2:
        st.markdown('<span style="font-size:24px; font-weight:600;">This Version</span>',unsafe_allow_html=True)
        st.write("""
        **Build v2.0.0.1**

        **Fixes**-
        - **PDF Report Generation:** A new automated report generation feature has been added.
        - Analysis for Wall and Roof parametrics now includes Low Density constructions as well. 
        """)
        st.markdown("</div>", unsafe_allow_html=True)
    # --------- Card 3: COMSim ----------
    with col3:
        st.markdown('<span style="font-size:24px; font-weight:600;">Disclaimer & Acknowledgement</span>',unsafe_allow_html=True)
        st.markdown("""
            **AWECSIM** (referred to as the *“Application”* hereafter) is an outcome of the best efforts of building simulation experts and IT developers at **Environmental Design Solutions Pvt. Ltd. (EDS)**.""")
        # if st.button("Read more"):
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
            # if os.path.exists(project_folder):
            #     st.warning("⚠️ Project name already exists! Please select another name.")
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
        main_typologies = ["Business", "Retail", "Hospital", "Hotel", "Residential", "School", "Assembly"]
        st.write("🌆 Select Typology")
        selected_typology = st.selectbox("", main_typologies, label_visibility="collapsed", key="typology")

    bin_name = ""
    # Only show City dropdown if not "Other"
    if selected_country != "Custom Weather":
        with col2:
            st.write("🌎 Select City")
            user_input = st.selectbox("", filtered_locations, label_visibility="collapsed").lower()
            selected_ = location[location["Sim_location"].str.lower() == user_input.lower()]
            # st.write(user_input)
            # st.write(selected_)
            # Extract Ashrae Climate Zone
            if not selected_.empty:
                ashrae_zone = selected_["Ashrae Climate Zone"].iloc[0]
                ecsbc_zone = selected_["NBC Climate"].iloc[0]
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
        # st.cache_data.clear()
        # st.cache_resource.clear()
        # if uploaded_file is None:
        #     st.warning("⚠️ Please Upload .INP File!")
        #     st.stop()
        # if bin_name is None:
        #     st.warning("⚠️ Please Upload .BIN File!")
        #     st.stop()
        # if not project_name_clean:
        #     st.warning("⚠️ Please enter a project name.")
        #     st.stop()
        # with st.spinner("⚡ Processing... This may take a few minutes."):
        #     os.makedirs(output_inp_folder, exist_ok=True)
        #     new_batch_id = f"{int(time.time())}"  # unique ID

        #     selected_rows = updated_df[updated_df['Batch_ID'] == run_cnt]
        #     batch_output_folder = os.path.join(output_inp_folder, f"{user_nm}")
        #     os.makedirs(batch_output_folder, exist_ok=True)

        #     num = 1
        #     modified_files = []
        #     for _, row in selected_rows.iterrows():
        #         selected_inp = uploaded_file.name
        #         new_inp_name = f"{row['Wall']}_{row['Roof']}_{row['Glazing']}_{row['GlazingR']}_{row['Light']}_{row['WWR']}_{row['Equip']}_{selected_inp}"
        #         new_inp_path = os.path.join(batch_output_folder, new_inp_name)

        #         inp_file_path = os.path.join(inp_folder, selected_inp)
        #         if not os.path.exists(inp_file_path):
        #             st.error(f"File {inp_file_path} not found. Skipping.")
        #             continue

        #         # st.info(f"Modifying INP file {num}: {selected_inp} -> {new_inp_name}")
        #         num += 1

        #         # Apply modifications
        #         inp_content = wwr.process_window_insertion_workflow(inp_file_path, row["WWR"])
        #         # inp_content = orient.updateOrientation(inp_content, row["Orient"])
        #         inp_content = lighting.updateLPD(inp_content, row['Light'])
        #         # inp_content = insertWall.update_Material_Layers_Construction(inp_content, row["Wall"])
        #         # inp_content = insertRoof.update_Material_Layers_Construction(inp_content, row["Roof"])
        #         # inp_content = insertRoof.removeDuplicates(inp_content)
        #         inp_content = equip.updateEquipment(inp_content, row['Equip'])
        #         inp_content = windows.insert_glass_types_multiple_outputs(inp_content, row['Glazing'])
        #         inp_content = windows.insert_glass_UVal(inp_content, row['GlazingR'])
        #         inp_content =remove_utility(inp_content)
        #         # inp_content = windows.readSCUVal(inp_content, row['Glazing'])
        #         if row['Light'] > 0:
        #             inp_content = remove_betweenLightEquip(inp_content)
        #         count = ModifyWallRoof.count_exterior_walls(inp_content)
        #         if count > 1:
        #             inp_content = ModifyWallRoof.fix_walls(inp_content, row["Wall"])
        #             inp_content = ModifyWallRoof.fix_roofs(inp_content, row["Roof"])
        #             # inp_content = insertRoof.removeDuplicates(inp_content)
        #             with open(new_inp_path, 'w') as file:
        #                 file.writelines(inp_content)
        #             modified_files.append(new_inp_name)
        #         else:
        #             st.write("No Exterior-Wall Exists!")
        
        #     simulate_files = []
        #     if uploaded_file is None:
        #         st.error("Please upload an INP file before starting the simulation.")
        #     else:
        #         # st.markdown(f"<span style='color:green;'>✅ Updating DAYLIGHTING from YES to NO!</span>", unsafe_allow_html=True)
        #         script_dir = os.path.dirname(os.path.abspath(__file__))
        #         shutil.copy(os.path.join(script_dir, "script.bat"), batch_output_folder)
        #         inp_files = [f for f in os.listdir(batch_output_folder) if f.lower().endswith(".inp")]
        #         for inp_file in inp_files:
        #             file_path = os.path.join(batch_output_folder, os.path.splitext(inp_file)[0])
        #             subprocess.call(
        #                 [os.path.join(batch_output_folder, "script.bat"), file_path, weather_path],
        #                 shell=True
        #             )
        #             simulate_files.append(inp_file)
            
        #         subprocess.call([os.path.join(batch_output_folder, "script.bat"), batch_output_folder, weather_path], shell=True)
        #         required_sections = ['BEPS', 'BEPU', 'LS-C', 'LV-B', 'LV-D', 'PS-E', 'SV-A']
        #         log_file_path = check_missing_sections(batch_output_folder, required_sections, new_batch_id, user_nm)
        #         get_failed_simulation_data(batch_output_folder, log_file_path)
        #         clean_folder(batch_output_folder)
        #         combined_Data = get_files_for_data_extraction(batch_output_folder, log_file_path, new_batch_id, location_id, user_nm, user_input, selected_typology)
        #         combined_Data = combined_Data.reset_index(drop=True)
                
        st.session_state.all_figs = []   # 🔥 MUST RESET HERE
        exportCSV = resource_path(os.path.join("2026-02-18T08-49_export.csv"))
        combined_Data = pd.read_csv(exportCSV)
        # st.write(combined_Data)
        ashrae_zone = "Very Cold"
        # combined_Data["Equip(W/Sqft)"] = combined_Data["Equipment-Total(W)"] / combined_Data["Floor-Total-Above-Grade(SQFT)"]
        # combined_Data["Light(W/Sqft)"] = combined_Data["Power Lighting Total(W)"] / combined_Data["Floor-Total-Above-Grade(SQFT)"]
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
        usage_factor = round(row0["EFLH"], 0) if "EFLH" in row0 else None
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
                    <div class="card"><span class="info-icon">ⓘ</span><div class="tooltip">Roof area relative to floor area</div><small>Envelope-to-Floor Ratio</small><h6>{roof_floor_ratio:.1f}</h6></div>
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
        combined_Data_expanded["Energy_Saving_%"] = ((baseline_energy - combined_Data_expanded["Energy_Outcome(KWH)"])/ baseline_energy) * 100
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
        json_path = "rulesets/ecsbc_envelope_ruleset.json"
        if os.path.exists(json_path):
            try:
                with open(json_path, "r", encoding="utf-8") as f:
                    ruleset_data = json.load(f)
            except Exception as e:
                st.error(f"Error reading JSON file: {e}")
        else:
            st.error("Ruleset file not found ❌")
        # st.write(combined_Data)
        cols = [
            "WALL CONDUCTION",
            "ROOF CONDUCTION",
            "WINDOW GLASS+FRM COND",
            "WINDOW GLASS SOLAR",
            "DOOR CONDUCTION",
            "INTERNAL SURFACE COND",
            "UNDERGROUND SURF COND",
            "OCCUPANTS TO SPACE",
            "LIGHT TO SPACE",
            "EQUIPMENT TO SPACE",
            "PROCESS TO SPACE",
            "INFILTRATION"
        ]
        wall_df["Total Load(kW)"] = wall_df[cols].sum(axis=1)
        roof_df["Total Load(kW)"] = roof_df[cols].sum(axis=1)
        wall_df["Final_R_Wall"] = np.where(wall_df["R-Value-Wall"].isna() | (wall_df["R-Value-Wall"] == ""),wall_df["R-VAL-W"],wall_df["R-Value-Wall"])
        roof_df["Final_R_Roof"] = np.where(roof_df["R-Value-Roof"].isna() | (roof_df["R-Value-Roof"] == ""),roof_df["R-VAL-R"],roof_df["R-Value-Roof"])

        # Create color column
        wall_df["bar_color"] = np.where(wall_df["R-Value-Wall"].isna() | (wall_df["R-Value-Wall"] == ""),"red", "#6B6B6B")
        roof_df["bar_color"] = np.where(roof_df["R-Value-Roof"].isna() | (roof_df["R-Value-Roof"] == ""),"red", "#6B6B6B")

        # st.write(wall_df)
        # st.plotly_chart(fig, use_container_width=True)
        all_figs = []   # <-- collect all charts here
        st.markdown(f"""<br>""", unsafe_allow_html=True)
        st.markdown("<h5 style='color:#dc2626; text-align:left;'>Energy Use</h5>", unsafe_allow_html=True)
        fig_wall = energy_param_plot_wall(wall_df, baseline, ecsbc_zone, selected_typology, "R-VAL-W","Wall R-Value (h·ft²·°F/Btu)","Energy Use vs Wall R-Value",show_legend=True)
        fig_roof = energy_param_plot(roof_df, baseline, ecsbc_zone, selected_typology, "R-VAL-R","Roof R-Value (h·ft²·°F/Btu)","Energy Use vs Roof R-Value",show_legend=True)
        # all_figs.extend([fig_wall, fig_roof])
        st.session_state.all_figs.extend([fig_wall, fig_roof])
        st.markdown(f"""<br>""", unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1:
            st.plotly_chart(fig_wall, use_container_width=True)
        with col2:
            st.plotly_chart(fig_roof, use_container_width=True)
        # st.write(light_df)
        # st.write(equip_df)
        # st.write(glazing_df)
        fig_light = make_single_plot(light_df,x_col="Light(W/Sqft)",title="Energy Use vs Lighting Power",x_label="Lighting Power Density (W/ft²)")
        fig_equip = make_single_plot(equip_df,x_col="Equip(W/Sqft)",title="Energy Use vs Equipment Power",x_label="Equipment Power Density (W/ft²)")
        all_figs.extend([fig_light, fig_equip])
        st.session_state.all_figs.extend([fig_light, fig_equip])
        col1, col2 = st.columns(2)
        with col1:
            st.plotly_chart(fig_light, use_container_width=True)
        with col2:
            st.plotly_chart(fig_equip, use_container_width=True)
        
        fig_SC_glazing = make_single_plot_shgc(glazing_df,x_col="SHGC",title="Energy Use vs SHGC",x_label="SHGC")
        fig_R_glazing = energy_param_plot_wind(glazingr_df,baseline,"R-VAL-Wind","Window R-Value (h·ft²·°F/Btu)","Energy Use vs Window R-Value",show_legend=True)
        # fig_R_glazing = make_single_plot(glazingr_df,x_col="R-VAL-Wind",title="Energy Use vs R-Value",x_label="R-Value (HR·ft²·°F / Btu)")
        st.session_state.all_figs.extend([fig_SC_glazing, fig_R_glazing])
        col1, col2 = st.columns(2)
        with col1:
            st.plotly_chart(fig_SC_glazing, use_container_width=True)
        with col2:
            st.plotly_chart(fig_R_glazing, use_container_width=True)
        # st.divider()
        st.markdown("<h5 style='color:#dc2626; text-align:left;'>Gains and Losses</h5>", unsafe_allow_html=True)

        # Sort for clean structure
        wall_df = wall_df.sort_values("Final_R_Wall")
        roof_df = roof_df.sort_values("Final_R_Roof")
        # st.write(wall_df)
        # st.write(baseline)
        # st.write(glazingr_df)
        # fig_wallGains = energy_param_plot_wallGains(wall_df, baseline, ecsbc_zone, selected_typology, "Final_R_Wall","Wall R-Value (h·ft²·°F/Btu)","Energy Use vs Wall R-Value",show_legend=True)
        # fig_roofGains = energy_param_plot_roofGains(roof_df, baseline, ecsbc_zone, selected_typology, "Final_R_Roof","Roof R-Value (h·ft²·°F/Btu)","Energy Use vs Roof R-Value",show_legend=True)
        # fig_wallLoss = energy_param_plot_wallLoss(wall_df, baseline, ecsbc_zone, selected_typology, "Final_R_Wall","Wall R-Value (h·ft²·°F/Btu)","Energy Use vs Wall R-Value",show_legend=True)
        # fig_roofLoss = energy_param_plot_roofLoss(roof_df, baseline, ecsbc_zone, selected_typology, "Final_R_Roof","Roof R-Value (h·ft²·°F/Btu)","Energy Use vs Roof R-Value",show_legend=True)
        fig_wallGains = energy_param_plot_wall_gains(wall_df, baseline, ecsbc_zone, selected_typology, "R-VAL-W","Wall R-Value (h·ft²·°F/Btu)","Energy Use vs Wall R-Value",show_legend=True)
        fig_roofGains = energy_param_plot_roof_gains(roof_df, baseline, ecsbc_zone, selected_typology, "R-VAL-R","Roof R-Value (h·ft²·°F/Btu)","Energy Use vs Roof R-Value",show_legend=True)
        fig_wallLoss = energy_param_plot_wall_loss(wall_df, baseline, ecsbc_zone, selected_typology, "R-VAL-W","Wall R-Value (h·ft²·°F/Btu)","Energy Use vs Wall R-Value",show_legend=True)
        fig_roofLoss = energy_param_plot_roof_loss(roof_df, baseline, ecsbc_zone, selected_typology, "R-VAL-R","Roof R-Value (h·ft²·°F/Btu)","Energy Use vs Roof R-Value",show_legend=True)
        fig_windowGains = energy_param_plot_window_gains(glazingr_df, baseline, ecsbc_zone, selected_typology, "R-VAL-Wind","Glazing R-Value (h·ft²·°F/Btu)","Energy Use vs Glazing R-Value",show_legend=True)
        fig_windowLoss = energy_param_plot_window_loss(glazingr_df, baseline, ecsbc_zone, selected_typology, "R-VAL-Wind","Glazing R-Value (h·ft²·°F/Btu)","Energy Use vs Glazing R-Value",show_legend=True)
        fig_solarGains = energy_param_plot_solar_gains(glazing_df, baseline, ecsbc_zone, selected_typology, "SHGC","Glazing SHGC","Energy Use vs Glazing SHGC",show_legend=True)
        fig_solarLoss = energy_param_plot_solar_loss(glazing_df, baseline, ecsbc_zone, selected_typology, "SHGC","Glazing SHGC","Energy Use vs Glazing SHGC",show_legend=True)
        
        col1, col2 = st.columns(2)
        with col1:
            st.plotly_chart(fig_wallGains, use_container_width=True, key="wallGains")
        with col2:
            st.plotly_chart(fig_wallLoss, use_container_width=True, key="wallLoss")
        col1, col2 = st.columns(2)
        with col1:
            st.plotly_chart(fig_roofGains, use_container_width=True, key="roofGains")
        with col2:
            st.plotly_chart(fig_roofLoss, use_container_width=True, key="roofLoss")
        col1, col2 = st.columns(2)
        with col1:
            st.plotly_chart(fig_windowGains, use_container_width=True, key="windGain")
        with col2:
            st.plotly_chart(fig_windowLoss, use_container_width=True, key="windLoss")
        col1, col2 = st.columns(2)
        with col1:
            st.plotly_chart(fig_solarGains, use_container_width=True, key="solarGain")
        with col2:
            st.plotly_chart(fig_solarLoss, use_container_width=True, key="solarLoss")
        st.session_state.all_figs.extend([fig_wallGains, fig_roofGains, fig_wallLoss, fig_roofLoss])
        st.session_state.all_figs.extend([fig_windowGains, fig_windowLoss, fig_solarGains, fig_solarLoss])
        
        components = [
            "WALL CONDUCTION",
            "ROOF CONDUCTION",
            "WINDOW GLASS+FRM COND",
            "WINDOW GLASS SOLAR",
            "DOOR CONDUCTION",
            "INTERNAL SURFACE COND",
            "UNDERGROUND SURF COND",
            "OCCUPANTS TO SPACE",
            "LIGHT TO SPACE",
            "EQUIPMENT TO SPACE",
            "PROCESS TO SPACE",
            "INFILTRATION"
        ]
        density_color_map = {
            "High Density": "#6B6B6B",  # dark grey
            "Low Density": "#B0B0B0",   # light grey
        }

        col1, col2 = st.columns(2)
        wall_df["final_color"] = wall_df.apply(get_bar_color, axis=1)
        roof_df["final_color"] = roof_df.apply(get_bar_color_roof, axis=1)
        # st.write(wall_df)
        # with col1:
        #     wall_df = wall_df[wall_df["Final_R_Wall"] != 30].copy()
        #     wall_df = wall_df.sort_values("Final_R_Wall")

        #     wall_df["density_color"] = wall_df["DensityType"].map(density_color_map)

        #     fig = go.Figure()

        #     # ---- Existing Conduction ----
        #     fig.add_trace(go.Bar(
        #         x=wall_df["Final_R_Wall"],
        #         y=wall_df["WALL CONDUCTION"],
        #         name="Wall Conduction",
        #         marker_color=wall_df["final_color"],
        #         width=0.6,
        #         hovertemplate=(
        #             "<b>Density:</b> %{customdata[1]}<br>"
        #             "<b>R-Value:</b> %{x:.2f} (h·ft²·°F/Btu)<br>"
        #             "<b>Wall Conduction:</b> %{y:.2f} kW<br>"
        #             "<extra></extra>"
        #         ),
        #         customdata=wall_df[["FileName", "DensityType"]].values
        #     ))

        #     # ---- NEW: Conduction Loss ----
        #     fig.add_trace(go.Bar(
        #         x=wall_df["Final_R_Wall"],
        #         y=wall_df["WALL CONDUCTION_loss"],
        #         name="Wall Conduction Loss",
        #         marker_color=wall_df["final_color"],
        #         width=0.6,
        #         hovertemplate=(
        #             "<b>Density:</b> %{customdata[1]}<br>"
        #             "<b>R-Value:</b> %{x:.2f} (h·ft²·°F/Btu)<br>"
        #             "<b>Wall Loss:</b> %{y:.2f} kW<br>"
        #             "<extra></extra>"
        #         ),
        #         customdata=wall_df[["FileName", "DensityType"]].values
        #     ))

        #     fig.update_layout(
        #         barmode="overlay",   # ✅ both bars on same x
        #         title_text="",
        #         xaxis_title="Wall R-Value (h·ft²·°F/Btu)",
        #         yaxis_title="Wall Conduction (kW)",
        #         template="plotly_white",
        #         height=560,
        #         bargap=0.35,
        #         xaxis=dict(type="linear"),
        #         showlegend=False
        #     )

        #     st.plotly_chart(fig, use_container_width=True, key="wall_rvalue_charts")

        # with col2:
        #     roof_df = roof_df[roof_df["Final_R_Roof"] != 30].copy()
        #     roof_df = roof_df.sort_values("Final_R_Roof")

        #     roof_df["density_color"] = roof_df["DensityType"].map(density_color_map)

        #     fig = go.Figure()

        #     # ---- Existing Conduction ----
        #     fig.add_trace(go.Bar(
        #         x=roof_df["Final_R_Roof"],
        #         y=roof_df["ROOF CONDUCTION"],
        #         name="Roof Conduction",
        #         marker_color=roof_df["final_color"],
        #         width=0.6,
        #         hovertemplate=(
        #             "<b>Density:</b> %{customdata[1]}<br>"
        #             "<b>R-Value:</b> %{x:.2f} (h·ft²·°F/Btu)<br>"
        #             "<b>Roof Conduction:</b> %{y:.2f} kW<br>"
        #             "<extra></extra>"
        #         ),
        #         customdata=roof_df[["FileName", "DensityType"]].values
        #     ))

        #     # ---- NEW: Conduction Loss ----
        #     fig.add_trace(go.Bar(
        #         x=roof_df["Final_R_Roof"],
        #         y=roof_df["ROOF CONDUCTION_loss"],
        #         name="Roof Conduction Loss",
        #         marker_color=roof_df["final_color"],
        #         width=0.6,
        #         hovertemplate=(
        #             "<b>Density:</b> %{customdata[1]}<br>"
        #             "<b>R-Value:</b> %{x:.2f} (h·ft²·°F/Btu)<br>"
        #             "<b>Roof Loss:</b> %{y:.2f} kW<br>"
        #             "<extra></extra>"
        #         ),
        #         customdata=roof_df[["FileName", "DensityType"]].values
        #     ))

        #     fig.update_layout(
        #         barmode="overlay",   # ✅ same graph overlay
        #         title_text="",
        #         xaxis_title="Roof R-Value (h·ft²·°F/Btu)",
        #         yaxis_title="Roof Conduction (kW)",
        #         template="plotly_white",
        #         height=560,
        #         bargap=0.35,
        #         xaxis=dict(type="linear"),
        #         showlegend=False
        #     )
        #     st.plotly_chart(fig, use_container_width=True, key="roof_rvalue_charts")

        # ---- SAVE IMMEDIATELY ----
        st.session_state.area_ft2 = area_ft2
        st.session_state.conditioned_pct = conditioned_pct
        st.session_state.wwr = round(wwr*100,1)
        st.session_state.wall_floor_ratio = round(wall_floor_ratio*100,1)
        st.session_state.above_grade_pct = above_grade_pct
        st.session_state.roof_floor_ratio = round(roof_floor_ratio*100,1)
        st.session_state.usage_factor = usage_factor
        st.session_state.wall_pct = wall_pct
        st.session_state.roof_pct = roof_pct
        st.session_state.light_pct = light_pct
        st.session_state.equip_pct = equip_pct
        st.session_state.glazing_pct = glazing_pct
        st.session_state.range_energy = range_energy
        st.session_state.range_demand = range_demand
        st.session_state.ecsbc_zone = ecsbc_zone
        st.divider()

        # ---------- SESSION STATE ----------


        # col1, col2, col3, col4 = st.columns(4, gap="large")
        # # -------- GRID 1 -------- #
        # with col1:
        #     st.markdown(f"""<h5 class="section-title" style="color:#dc2626">Wall</h5>""", unsafe_allow_html=True)
        #     st.markdown("""
        #     **Materials:**  
        #     Cement Plaster · Brick · Cement Board · Gypsum Board  

        #     **Insulation:**  
        #     XPS = Extruded Polystyrene 

        #     **R-Value:**  
        #     R 2.5 → R 30
        #     """)

        # # -------- GRID 2 -------- #
        # with col2:
        #     st.markdown(f"""<h5 class="section-title" style="color:#dc2626">Roof</h5>""", unsafe_allow_html=True)
        #     st.markdown("""
        #     **Materials:**  
        #     Cement Plaster · Brick · Concrete · Metal Deck  

        #     **Insulation:**  
        #     XPS = Extruded Polystyrene

        #     **R-Value:**  
        #     R 2.5 → R 30
        #     """)

        # # -------- GRID 3 -------- #
        # with col3:
        #     st.markdown(f"""<h5 class="section-title" style="color:#dc2626">Loads</h5>""", unsafe_allow_html=True)
        #     st.markdown("""
        #     **Equipment:**  
        #     1.0 → 2.0 W/ft²

        #     **Lighting:**  
        #     0.3 → 1.0 W/ft²
        #     """)

        # # -------- GRID 4 -------- #
        # with col4:
        #     st.markdown(f"""<h5 class="section-title" style="color:#dc2626">Glazing</h5>""", unsafe_allow_html=True)
        #     st.markdown("""
        #     **Shading Coefficient:**  
        #     0.25 · 0.30 · 0.40 ... 0.90

        #     **R-Value:**  
        #     R 2.5 → R 30
        #     """)

        # # -------- DATA MODEL (CSV DOWNLOAD) -------- #
        # data = [
        #     ["Wall", "Material", "CP, Brick (High Density), Concrete, Cement Board, Gypsum Board (Low Density), Metal Deck"],
        #     ["Wall", "Insulation", "XPS"],
        #     ["Wall", "R-Value", "R 2.5 → R 30"],
        #     ["Roof", "Material", "CP, Brick, Concrete, Metal Deck"],
        #     ["Roof", "Insulation", "XPS"],
        #     ["Roof", "R-Value", "R 2.5 → R 30"],
        #     ["Load", "Equipment Density", "1.0 → 2.0 W/sqft"],
        #     ["Load", "Lighting Density", "0.3 → 1.0 W/sqft"],
        #     ["Glazing", "Shading Coefficient", "0.25, 0.30, 0.40, 0.50, 0.60, 0.70, 0.80, 0.90"],
        #     ["Glazing", "R-Value", "R 2.5 → R 30"],
        # ]

        # df = pd.DataFrame(data, columns=["Category", "Parameter", "Range/Options"])
        # csv = df.to_csv(index=False).encode("utf-8")

        # st.download_button(
        #     "⬇️ Download Parameters Database",
        #     data=csv,
        #     file_name="parametric_model_database.csv",
        #     mime="text/csv"
        # )
    
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
                "wwr": st.session_state.get("wwr"),
                "wfr": st.session_state.get("wall_floor_ratio"),
                "climate": cz,
                "agArea": st.session_state.get("above_grade_pct"),
                "envelopeFloorArea": st.session_state.get("roof_floor_ratio"),
                "estimateHrsUse": st.session_state.get("usage_factor"),
                "wallpct": st.session_state.get("wall_pct"),
                "roofpct": st.session_state.get("roof_pct"),
                "lightpct": st.session_state.get("light_pct"),
                "equippct": st.session_state.get("equip_pct"),
                "wwr": st.session_state.get("wwr"),
                "glazingpct": st.session_state.get("glazing_pct"),
                "EnergyRange": st.session_state.get("range_energy"),
                "rangeDemand": st.session_state.get("range_demand"),
                "ecsbcZone":ecsbc_zone
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

if st.session_state.script_choice == "tool2":
    def resource_path(relative_path):
        return os.path.join(os.getcwd(), relative_path)
    csv_path1 = resource_path(os.path.join("database", "Simulation_locations.csv"))
    weather_df1 = pd.read_csv(csv_path1)
    db_path1 = resource_path(os.path.join("database", "AllData.xlsx"))
    location_path1 = resource_path(os.path.join("database", "Simulation_locations.csv"))
    location1 = pd.read_csv(location_path1)
    locations1 = sorted(location1['Sim_location'].unique())
    # --- Read Excel Sheets ---
    exportxlsx = resource_path(os.path.join("database", "AllData.xlsx"))
    wallDB = pd.read_excel(exportxlsx, sheet_name="Wall")
    roofDB = pd.read_excel(exportxlsx, sheet_name="Roof")
    windDB = pd.read_excel(exportxlsx, sheet_name="Glazing")
    wwrDB = pd.read_excel(exportxlsx, sheet_name="WWR")
    lightDB = pd.read_excel(exportxlsx, sheet_name="Light")
    equipDB = pd.read_excel(exportxlsx, sheet_name="Equip")

    if "Country" in location1.columns:
        # If CSV already has Country column
        countries1 = sorted(location1["Country"].unique().tolist())
    else:
        # If your current CSV has only Indian cities — default to India
        countries1 = ["India"]
    col1, col2, col3 = st.columns([1,1,1])
    with col1:
        st.write("📝 Project Name")
        project_name1 = st.text_input("", placeholder="Enter project name", label_visibility="collapsed")
        
        # project_name = st.text_input("📝 Project Name", placeholder="Enter project name")
        project_name_clean1 = project_name1.replace(" ", "")
        user_nm1 = project_name_clean1
        if project_name_clean1:
            parent_dir1 = os.path.dirname(os.getcwd())
            batch_outputs_dir1 = os.path.join(parent_dir1, "Batch_Outputs")
            project_folder1 = os.path.join(batch_outputs_dir1, project_name_clean1)
        countries1.append("Custom Weather")
        st.write("🌎 Select Country")
        selected_country1 = st.selectbox("", countries1, label_visibility="collapsed")

        # Filter and sort locations for the selected country (if not "Other")
        if selected_country1 != "Custom Weather":
            if "Country" in location1.columns:
                filtered_locations1 = (location1[location1["Country"] == selected_country1]["Sim_location"].dropna().unique().tolist())
                filtered_locations1 = sorted(filtered_locations1)
            else:
                filtered_locations1 = sorted(location1["Sim_location"].dropna().tolist())
        else:
            filtered_locations1 = []  # No locations for "Other"
    with col2:
        # Main typologies
        main_typologies1 = ["Business", "Retail", "Hospital", "Hotel", "Residential", "School", "Assembly"]
        st.write("🌆 Select Typology")
        selected_typology1 = st.selectbox("", main_typologies1, label_visibility="collapsed", key="typology")

    bin_name1 = ""
    # Only show City dropdown if not "Other"
    if selected_country1 != "Custom Weather":
        with col2:
            st.write("🌎 Select City")
            user_input1 = st.selectbox("", filtered_locations1, label_visibility="collapsed").lower()
            selected_1 = location1[location1["Sim_location"].str.lower() == user_input1.lower()]
            # st.write(user_input)
            # st.write(selected_)
            # Extract Ashrae Climate Zone
            if not selected_1.empty:
                ashrae_zone1 = selected_1["Ashrae Climate Zone"].iloc[0]
                ecsbc_zone1 = selected_1["NBC Climate"].iloc[0]
    else:
        with col2:
            user_input1 = "Other-City"
            # When "Other" is selected, show .bin upload option
            st.write("📤 Upload .bin file")
            uploaded_bin1 = st.file_uploader("", type=["bin"], label_visibility="collapsed")
            if uploaded_bin1 is not None:
                save_folder1 = r"C:\doe22\weather"
                os.makedirs(save_folder1, exist_ok=True)
                # Always rename uploaded file to 1.bin
                save_path1 = os.path.join(save_folder1, "1.bin")
                # Save uploaded file as 1.bin
                with open(save_path1, "wb") as f:
                    f.write(uploaded_bin1.getbuffer())
                bin_name1 = "1"   # without extension
            
    with col3:
        st.write("📤 Upload eQUEST INP file")
        uploaded_file1 = st.file_uploader("", type=["inp"], label_visibility="collapsed")
        if uploaded_file1:
            # st.write(uploaded_file)
            if uploaded_file1.name != 'F1.inp':
                uploaded_file1.name = user_nm1 + '.inp'

            # Go one step outside current working directory
            parent_dir1 = os.path.dirname(os.getcwd())
            batch_outputs_dir1 = os.path.join(parent_dir1, "Batch_Outputs_Com")
            os.makedirs(batch_outputs_dir1, exist_ok=True)
            uploaded_inp_path1 = os.path.join(batch_outputs_dir1, uploaded_file1.name)
            with open(uploaded_inp_path1, "wb") as f:
                f.write(uploaded_file1.getbuffer())
            output_inp_folder1 = os.path.dirname(uploaded_inp_path1)
            inp_folder1 = output_inp_folder1
            project_folder1 = os.path.join(batch_outputs_dir1, project_name_clean1)
    run_cnt1 = 1
    location_id1, weather_path1 = "", ""
    if user_input1 != "Other-City":
        matched_row1 = weather_df1[weather_df1['Sim_location'].str.lower().str.contains(user_input1)]
    else:
        matched_row1 = pd.DataFrame()
    if not matched_row1.empty:
        location_id1 = matched_row1.iloc[0]['Location_ID']
        weather_path1 = matched_row1.iloc[0]['Weather_file']
    elif user_input1:
        weather_path1 = bin_name1
    
    if "all_figs1" not in st.session_state:
        st.session_state.all_figs1 = []

    col1, col2, col3, col4, col5, col6, col7 = st.columns(7)
    # --- Wall ---
    with col1:
        wall_options = ["As Designed"] + [f"R{r}" for r in [2.5, 5, 7.5, 10, 12.5, 15, 17.5, 20, 22.5, 25, 27.5, 30]]
        wall_choice = st.selectbox("Wall R-Value", wall_options)
    # --- Roof ---
    with col2:
        roof_options = ["As Designed"] + [f"R{r}" for r in [2.5, 5, 7.5, 10, 12.5, 15, 17.5, 20, 22.5, 25, 27.5, 30]]
        roof_choice = st.selectbox("Roof R-Value", roof_options)
    # --- Glazing ---
    with col3:
        glazing_scoptions = ["As Designed"] + [0.25, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9]
        glazing_scchoice = st.selectbox("Glazing SC", glazing_scoptions)
    with col4:
        glazing_roptions = ["As Designed"] + [1.00, 1.52, 2.00, 2.50, 3.03, 3.45, 4.00, 4.55, 4.76]
        glazing_rchoice = st.selectbox("Glazing U-Value", glazing_roptions)
    # --- WWR ---
    with col5:
        wwr_values = wwrDB["WWR"].iloc[0:].dropna().unique()
        wwr_values_pct = ["As Designed"] + [f"{v*10}%" for v in wwr_values if v != 0]
        wwr_choice = st.selectbox("WWR", wwr_values_pct)
    # --- Light ---
    with col6:
        light_values = lightDB["Percent"].iloc[0:].dropna().unique()
        light_values_pct = ["As Designed"] + [0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 1]
        light_choice = st.selectbox("Lighting (W/ft²)", light_values_pct)
    # --- Equip ---
    with col7:
        equip_values = equipDB["Percent"].iloc[0:].dropna().unique()
        equip_values_pct = ["As Designed"] + [1, 1.1, 1.2, 1.3, 1.4, 1.5, 1.6, 1.7, 1.8, 1.9, 2.0]
        equip_choice = st.selectbox("Equipment (W/ft²)", equip_values_pct)

    # --- Note ---
    # st.markdown("""**Note:** *Options in Glazings are coded as: <b>GlazingType, U-Value, Shading Coefficient, Lighting-Transmittence</b>.*""",unsafe_allow_html=True)
    # --- Convert selections to numeric codes ---
    def get_index(choice, options):
        """Return numeric code (0 for 'As Designed', else 1, 2, ...)"""
        return options.index(choice)

    wall_code = get_index(wall_choice, wall_options)
    roof_code = get_index(roof_choice, roof_options)
    glazing_code = get_index(glazing_scchoice, glazing_scoptions)
    glazing_rcode = get_index(glazing_rchoice, glazing_roptions)
    wwr_code = get_index(wwr_choice, wwr_values_pct)
    light_code = get_index(light_choice, light_values_pct)
    equip_code = get_index(equip_choice, equip_values_pct)

    # --- Create single-row dataframe ---
    output_df_New = pd.DataFrame([{"Batch_ID": 1, "Run_ID": 1, "Wall": wall_code, "Roof": roof_code, "Glazing": glazing_code, "WWR": wwr_code, "GlazingR": glazing_rcode, "Light": light_code, "Equip": equip_code}])
    asdes_df = pd.DataFrame([{"Batch_ID": 1, "Run_ID": 1, "Wall": 0, "Roof": 0, "Glazing": 0, "WWR": 0, "GlazingR": 0, "Light": 0, "Equip": 0}])
    output_df_New = pd.concat([output_df_New, asdes_df], ignore_index=True)

    from plotly.subplots import make_subplots  
    import plotly.graph_objects as go
    run_cnt1 = 1
    # st.write(output_df_New)
    if "all_figs_com" not in st.session_state:
        st.session_state.all_figs_com = []
    if st.button("Simulate 🚀") and uploaded_file1 is not None:
        if (output_df_New.iloc[0] == output_df_New.iloc[1]).all():
            st.warning("⚠️ Please select at least one parameter from the dropdown to compare.")
            st.stop()
        with st.spinner("⚡ Processing... This may take a few minutes."):
            # st.write(output_df_New)
            os.makedirs(output_inp_folder1, exist_ok=True)
            new_batch_id = f"{int(time.time())}"

            selected_rows = output_df_New[output_df_New['Batch_ID'] == run_cnt1]
            batch_output_folder1 = os.path.join(output_inp_folder1, f"{user_nm1}")
            os.makedirs(batch_output_folder1, exist_ok=True)
            # st.write(batch_output_folder1)

            num1 = 1
            modified_files1 = []
            for _, row in selected_rows.iterrows():
                selected_inp = uploaded_file1.name
                new_inp_name1 = f"{row['Wall']}_{row['Roof']}_{row['Glazing']}_{row['GlazingR']}_{row['Light']}_{row['WWR']}_{row['Equip']}_{selected_inp}"
                new_inp_path = os.path.join(batch_output_folder1, new_inp_name1)

                inp_file_path1 = os.path.join(inp_folder1, selected_inp)
                if not os.path.exists(inp_file_path1):
                    st.error(f"File {inp_file_path1} not found. Skipping.")
                    continue
                # st.info(f"Modifying INP file {num}: {selected_inp} -> {new_inp_name}")
                num1 += 1
                # st.write(new_inp_name1)
                # st.write(new_inp_path)

                # Apply modifications
                inp_content = wwr.process_window_insertion_workflow(inp_file_path1, row["WWR"])
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
                    inp_content = remove_betweenLightEquip(inp_content)
                count = ModifyWallRoof.count_exterior_walls(inp_content)
                if count > 1:
                    if row["Wall"] > 0 and row["Roof"] == 0:
                        inp_content = ModifyWallRoof.fix_walls(inp_content, row["Wall"])
                    if row["Wall"] == 0 and row["Roof"] > 0:
                        inp_content = ModifyWallRoof.fix_roofs(inp_content, row["Roof"])
                    if row["Wall"] > 0 and row["Roof"] > 0:
                        inp_content = ModifyWallRoof.fix_walls(inp_content, row["Wall"])
                        inp_content = ModifyWallRoof.fix_roofs(inp_content, row["Roof"])
                        inp_content = ModifyWallRoof.fix_walls_roofs(inp_content, row["Wall"], row["Roof"])
                        # inp_content = insertRoof.removeDuplicates(inp_content)
                    with open(new_inp_path, 'w') as file:
                        file.writelines(inp_content)
                    modified_files1.append(new_inp_name1)
                else:
                    st.write("No Exterior-Wall Exists!")
        
            # st.write(modified_files1)
            simulate_files1 = []
            if uploaded_file1 is None:
                st.error("Please upload an INP file before starting the simulation.")
            else:
                script_dir1 = os.path.dirname(os.path.abspath(__file__))
                shutil.copy(os.path.join(script_dir1, "script.bat"), batch_output_folder1)
                inp_files1 = [f for f in os.listdir(batch_output_folder1) if f.lower().endswith(".inp")]
                # st.write(inp_files1)
                for inp_file in inp_files1:
                    file_path1 = os.path.join(batch_output_folder1, os.path.splitext(inp_file)[0])
                    subprocess.call([os.path.join(batch_output_folder1, "script.bat"), file_path1, weather_path1], shell=True)
                    simulate_files1.append(inp_file)
                # st.write(simulate_files1)
            
                subprocess.call([os.path.join(batch_output_folder1, "script.bat"), batch_output_folder1, weather_path1], shell=True)
                required_sections = ['BEPS', 'BEPU', 'LS-C', 'LV-B', 'LV-D', 'PS-E', 'SV-A']
                log_file_path1 = check_missing_sections(batch_output_folder1, required_sections, new_batch_id, user_nm1)
                get_failed_simulation_data(batch_output_folder1, log_file_path1)
                clean_folder(batch_output_folder1)
                combined_Data1 = get_files_for_data_extraction(batch_output_folder1, log_file_path1, new_batch_id, location_id1, user_nm1, user_input1, selected_typology1)
                combined_Data1 = combined_Data1.reset_index(drop=True)

            # exportCSV = resource_path(os.path.join("2026-03-09T12-54_export.csv"))
            # combined_Data1 = pd.read_csv(exportCSV)
            # st.write(combined_Data1)

            ############################################
            ############################################
            ############################################
            row0 = combined_Data1.iloc[0]
            row1 = combined_Data1.iloc[1]
            wwr_0 = (row0.get("WWR", 0))
            wwr_1 = (row1.get("WWR", 0))
            shgc_0 = (row0.get("SHGC", 0))
            shgc_1 = (row1.get("SHGC", 0))
            r_wall_0 = (row0.get("R-VAL-W", 0))
            r_wall_1 = (row1.get("R-VAL-W", 0))
            r_roof_0 = (row0.get("R-VAL-R", 0))
            r_roof_1 = (row1.get("R-VAL-R", 0))
            r_wind_0 = (row0.get("R-VAL-Wind", 0))
            r_wind_1 = (row1.get("R-VAL-Wind", 0))
            ener_0 = (row0.get("Energy_Outcome(KWH)", 0))
            ener_1 = (row1.get("Energy_Outcome(KWH)", 0))
            light_0 = (row0.get("Light(W/Sqft)", 0))
            light_1 = (row1.get("Light(W/Sqft)", 0))
            equip_0 = (row0.get("Equip(W/Sqft)", 0))
            equip_1 = (row1.get("Equip(W/Sqft)", 0))

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
                    <h5 class="section-title" style="color:#dc2626">As Desined</h5>
                    <div class="analysis-grid">
                        <div class="card"><small>Wall R-Value (ft²·°F·h/BTU)</small><h6>{r_wall_0:.2f}</h6></div>
                        <div class="card"><small>Roof R-Value (ft²·°F·h/BTU)</small><h6>{r_roof_0:.2f}</h6></div>
                        <div class="card"><small>Window U-Value(BTU/ft²·°F·h)</small><h6>{1/r_wind_0:.2f}</h6></div>
                        <div class="card"><small>Shading Coefficient</small><h6>{shgc_0*0.87:.2f}</h6></div>
                        <div class="card"><small>Window-to-Wall Ratio</small><h6>{wwr_0*100:.1f}%</h6></div>
                        <div class="card"><small>Lighting (W/ft²)</small><h6>{light_0:.1f}</h6></div>
                        <div class="card"><small>Equipment (W/ft²)</small><h6>{equip_0:.1f}</h6></div>
                        <div class="card"><small>Energy Outcome (kWh)</small><h6>{ener_0:,.0f}</h6></div>
                    </div>
                """, 
                unsafe_allow_html=True)
            with col_line:
                st.markdown("<div style='border-left:3px solid red; height:245px'></div>",unsafe_allow_html=True)
                wall_val = float(r_wall_0 if wall_choice == "As Designed" else wall_choice.replace("R", ""))
                roof_val = float(r_roof_0 if roof_choice == "As Designed" else roof_choice.replace("R", ""))
                window_val = float(1/r_wind_0 if glazing_rchoice == "As Designed" else glazing_rchoice)
                shading_val = float(shgc_0*0.87 if glazing_scchoice == "As Designed" else glazing_scchoice)

                # Handle WWR (strip % if needed)
                if wwr_choice == "As Designed":
                    wwr_val = wwr_0*100
                else:
                    wwr_val = float(wwr_choice.strip('%'))

                # Lighting, Equipment, Energy
                light_val = float(light_0 if light_choice == "As Designed" else light_choice)
                equip_val = float(equip_0 if equip_choice == "As Designed" else equip_choice)
                ener_val = float(ener_0 if str(ener_1) == "As Designed" else ener_1)

            with col2:
                st.markdown(f"""
                    <h5 class="section-title" style="color:#dc2626">User Selected</h5>
                    <div class="analysis-grid">
                        <div class="card"><small>Wall R-Value (ft²·°F·h/BTU)</small><h6>{wall_val:.2f}</h6></div>
                        <div class="card"><small>Roof R-Value (ft²·°F·h/BTU)</small><h6>{roof_val:.2f}</h6></div>
                        <div class="card"><small>Window U-Value(BTU/ft²·°F·h)</small><h6>{window_val:.2f}</h6></div>
                        <div class="card"><small>Shading Coefficient</small><h6>{shading_val:.2f}</h6></div>
                        <div class="card"><small>Window-to-Wall Ratio</small><h6>{wwr_val:.1f}%</h6></div>
                        <div class="card"><small>Lighting (W/ft²)</small><h6>{light_val:.1f}</h6></div>
                        <div class="card"><small>Equipment (W/ft²)</small><h6>{equip_val:.1f}</h6></div>
                        <div class="card"><small>Energy Outcome (kWh)</small><h6>{ener_1:,.0f}</h6></div>
                    </div><br>
                """, unsafe_allow_html=True)
            
            st.session_state.wwr_0 = wwr_0*100
            st.session_state.wwr_1 = wwr_val
            st.session_state.shgc_0 = shgc_0*0.87
            st.session_state.shgc_1 = shading_val
            st.session_state.r_wall_0 = r_wall_0
            st.session_state.r_wall_1 = wall_val
            st.session_state.r_roof_0 = r_roof_0
            st.session_state.r_roof_1 = roof_val
            st.session_state.r_wind_0 = 1/r_wind_0
            st.session_state.r_wind_1 = window_val
            st.session_state.ener_0 = ener_0
            st.session_state.ener_1 = ener_1
            st.session_state.light_0  = light_0
            st.session_state.light_1  = light_val
            st.session_state.equip_0  = equip_0
            st.session_state.equip_1  = equip_val
            
            ############################################
            ################ Calculation ###############
            ############################################

            generated_names = [os.path.splitext(f)[0] for f in modified_files1]
            merged_df = combined_Data1[combined_Data1['FileName'].isin(generated_names)]
            # st.write(merged_df)
            merged_df.index = ["As Designed", "User Selected"]
            energy_cols = [
                "WALL CONDUCTION", "ROOF CONDUCTION", "WINDOW GLASS+FRM COND", "WINDOW GLASS SOLAR",
                "DOOR CONDUCTION", "INTERNAL SURFACE COND", "UNDERGROUND SURF COND",
                "OCCUPANTS TO SPACE", "LIGHT TO SPACE", "EQUIPMENT TO SPACE", "PROCESS TO SPACE", "INFILTRATION"
            ]
            energy_loss_cols = [
                "WALL CONDUCTION_loss", "ROOF CONDUCTION_loss", "WINDOW GLASS+FRM COND_loss", "WINDOW GLASS SOLAR_loss",
                "DOOR CONDUCTION_loss", "INTERNAL SURFACE COND_loss", "UNDERGROUND SURF COND_loss",
                "OCCUPANTS TO SPACE_loss", "LIGHT TO SPACE_loss", "EQUIPMENT TO SPACE_loss", "PROCESS TO SPACE_loss", "INFILTRATION_loss"
            ]
            as_designed = merged_df.loc["As Designed", energy_cols]
            as_designed_loss = merged_df.loc["As Designed", energy_loss_cols]

            ecm = merged_df.loc["User Selected", energy_cols]
            ecm_loss = merged_df.loc["User Selected", energy_loss_cols]

            m_df_all = pd.concat([as_designed, ecm], axis=1)
            m_df_all_loss = pd.concat([as_designed_loss, ecm_loss], axis=1)

            # parameters having negative values
            negative_params = m_df_all[(m_df_all < 0).any(axis=1)]
            positive_params = m_df_all_loss[(m_df_all_loss > 0).any(axis=1)]

            as_designed = as_designed[as_designed > 0]
            ecm = ecm[ecm > 0]
            # as_designed_loss = as_designed_loss[as_designed_loss > 0]
            # ecm_loss = ecm_loss[ecm_loss > 0]

            as_designed_loss = as_designed_loss[as_designed_loss < 0].abs()
            ecm_loss = ecm_loss[ecm_loss < 0].abs()

            m_df = pd.concat([as_designed, ecm], axis=1)
            m_df_loss = pd.concat([as_designed_loss, ecm_loss], axis=1)

            m_df = m_df.rename(columns={'As Designed': 'As Designed (kW)', "User Selected": 'User Selected (kW)'})
            m_df_all = m_df_all.rename(columns={'As Designed': 'As Designed (kW)', "User Selected": 'User Selected (kW)'})
            m_df_all_loss = m_df_all_loss.rename(columns={'As Designed': 'As Designed (kW)', "User Selected": 'User Selected (kW)'})
            m_df.insert(0, 'Parameters', m_df.index)
            m_df_all.insert(0, 'Parameters', m_df_all.index)
            m_df_all_loss.insert(0, 'Parameters', m_df_all_loss.index)
            m_df = m_df.loc[:, m_df.columns.notna()]          # drop NaN column names
            m_df = m_df.loc[:, m_df.columns != ''] 
            m_df['As Designed (kW)'] = pd.to_numeric(m_df['As Designed (kW)'], errors='coerce')
            m_df_all['As Designed (kW)'] = pd.to_numeric(m_df_all['As Designed (kW)'], errors='coerce')
            m_df_all_loss['As Designed (kW)'] = pd.to_numeric(m_df_all_loss['As Designed (kW)'], errors='coerce')
            m_df['User Selected (kW)'] = pd.to_numeric(m_df['User Selected (kW)'], errors='coerce')
            m_df_all['User Selected (kW)'] = pd.to_numeric(m_df_all['User Selected (kW)'], errors='coerce')
            m_df_all_loss['User Selected (kW)'] = pd.to_numeric(m_df_all_loss['User Selected (kW)'], errors='coerce')
            m_df['% Saving'] = ((m_df['As Designed (kW)'] - m_df['User Selected (kW)']) / m_df['As Designed (kW)']) * 100
            m_df_all['% Saving'] = ((m_df_all['As Designed (kW)'] - m_df_all['User Selected (kW)']) / m_df_all['As Designed (kW)']) * 100
            m_df_all_loss['% Saving'] = ((m_df_all_loss['As Designed (kW)'] - m_df_all_loss['User Selected (kW)']) / m_df_all_loss['As Designed (kW)']) * 100
            m_df['% Saving'] = m_df['% Saving'].round(1)
            m_df_all['% Saving'] = m_df_all['% Saving'].round(1)
            m_df_all_loss['% Saving'] = m_df_all_loss['% Saving'].round(1)
            m_df = m_df.reset_index(drop=True)
            m_df_all = m_df_all.reset_index(drop=True)
            m_df_all_loss = m_df_all_loss.reset_index(drop=True)

            ##########################################################
            ################## Charts Calculation ####################
            ##########################################################

            dfsi = []
            for file in os.listdir(batch_output_folder1):
                if file.lower().endswith("_bepu.csv"):
                    file_path = os.path.join(batch_output_folder1, file)
                    df = pd.read_csv(file_path)
                    dfsi.append(df)
            
            if dfsi:  # check if list is not empty
                dfs = dfsi[0]  # first BEPU CSV
                i = 1  # or any valid index less than len(dfsi)
                df = dfsi[i]
            else:
                st.warning("No _bepu.csv files found.")
            # st.write(dfs) # as desined
            # st.write(df) # user selected
            # df = df.drop(index=[0, 1])
            # dfs = dfs.drop(index=[0, 1])
            numeric_cols = df.columns[3:]
            numeric_cols1 = dfs.columns[3:]
            df[numeric_cols] = df[numeric_cols].apply(pd.to_numeric, errors='coerce')
            dfs[numeric_cols1] = dfs[numeric_cols1].apply(pd.to_numeric, errors='coerce')
            # st.write(numeric_cols)
            # st.write(numeric_cols1)
            # 3️⃣ Group by 'Units' and sum numeric columns
            unit_sum_df = df.groupby("BEPU-UNIT")[numeric_cols].sum().reset_index()
            unit_sum_dfs = dfs.groupby("BEPU-UNIT")[numeric_cols1].sum().reset_index()
            # Filter only KWH rows
            df_asdes_kwh = unit_sum_dfs[unit_sum_dfs["BEPU-UNIT"] == "KWH"]
            df_ecm_kwh = unit_sum_df[unit_sum_df["BEPU-UNIT"] == "KWH"]
            param_cols = [col for col in unit_sum_dfs.columns if col not in ['Unnamed: 0', 'BEPU-UNIT', 'TOTAL-BEPU']]

            # Create a combined DataFrame
            combined_df = pd.DataFrame({"Parameters": param_cols, "As Designed (kWh)": df_asdes_kwh[param_cols].iloc[0].values, "User Selected (kWh)": df_ecm_kwh[param_cols].iloc[0].values})
            combined_df["% Saving"] = ((combined_df["As Designed (kWh)"] - combined_df["User Selected (kWh)"]) / combined_df["As Designed (kWh)"] * 100).round(2)
            combined_df = combined_df[(combined_df["As Designed (kWh)"] > 0) | (combined_df["User Selected (kWh)"] > 0)]
            # st.write(combined_df)
            from plotly.subplots import make_subplots
            import plotly.graph_objects as go
            unit_sum_df = unit_sum_df.loc[:, (unit_sum_df != 0).any(axis=0)]
            unit_sum_dfs = unit_sum_dfs.loc[:, (unit_sum_dfs != 0).any(axis=0)]
            component_columns_1 = [col for col in unit_sum_df.columns if col != "TOTAL-BEPU"]
            component_columns_2 = [col for col in unit_sum_dfs.columns if col != "TOTAL-BEPU"]
            contribution2 = unit_sum_df[component_columns_1].sum().reset_index()
            contribution2.columns = ["Component", "Value"]
            contribution1 = unit_sum_dfs[component_columns_2].sum().reset_index()
            contribution1.columns = ["Component", "Value"]
        
            color_map = {
                "MISQ-EQUIP": "#1f77b4",
                "LIGHTS": "#aec7e8",
                "SPACE-COOLING": "#ff4d4d",
                "VENT FANS": "#f4a3a3",
                "EXT-USAGE": "#2ca58d",
                "PUMPS & AUX": "#7bd389",
                "HEAT-REJECT": "#ff9f1c",
                "SPACE-HEATING": "#ffd166",
                "HT-PUMP-SUPPLEMENT": "#9b5de5",
                "BEPU-UNIT": "#9467bd"
            }

            figEUI = make_subplots(
                rows=1, cols=2,
                specs=[[{'type':'domain'}, {'type':'domain'}]],
                subplot_titles=('As Designed', 'User Selected'),
                horizontal_spacing=0.05
            )

            figEUI.add_trace(
                go.Pie(
                    labels=contribution1["Component"],
                    values=contribution1["Value"],
                    marker=dict(colors=[color_map.get(i, "#808080") for i in contribution1["Component"]]),
                    # marker=dict(colors=[color_map[i] for i in contribution1["Component"]]),
                    name="As Designed"
                ),1,1
            )

            figEUI.add_trace(
                go.Pie(
                    labels=contribution2["Component"],
                    values=contribution2["Value"],
                    marker=dict(colors=[color_map.get(i, "#808080") for i in contribution2["Component"]]),
                    name="User Selected"
                ),1,2
            )

            figEUI.update_traces(
                hole=0.35,
                textinfo="percent+label",
                hoverinfo="label+value+percent",
                textfont=dict(size=12)
            )

            figEUI.update_layout(
                showlegend=True,
                legend=dict(
                    orientation="h",
                    yanchor="bottom",
                    y=-0.15,
                    xanchor="center",
                    x=0.5
                ),
                margin=dict(t=50, b=60, l=30, r=30),
                height=500
            )
            import plotly.io as pio
            pio.templates.default = "none"
            figEUI.update_traces(textposition="inside")

            custom_colors = ["#FF4B4B", "#1E90FF"]
            temp_df = combined_df.rename(columns={"As Designed (kWh)": "As Designed", "User Selected (kWh)": "User Selected"})

            figA = px.bar(
                temp_df,
                x="Parameters",
                y=["As Designed", "User Selected"],
                barmode="group",
                labels={"value": "Energy (kWh)", "Parameters": ""},  # Remove x-axis label
                color_discrete_sequence=custom_colors
            )

            figA.update_layout(
                    legend=dict(
                        title_text="",           # no legend title
                        orientation="h",          # horizontal legend
                        yanchor="bottom",
                        y=-0.4,                   # below chart
                        xanchor="center",
                        x=0.5,
                        font=dict(size=12)
                    ),
                    xaxis_tickangle=-45,
                    margin=dict(t=20, b=80)
                )
            
            
            # --- Chart 2: Stacked Bar Chart ---
            combined_df = combined_df.drop('% Saving', axis=1)
            # --- Build Horizontal Stacked Bar Chart manually ---
            figB = go.Figure()
            for param in combined_df["Parameters"]:
                values = combined_df.loc[
                    combined_df["Parameters"] == param,
                    ["As Designed (kWh)", "User Selected (kWh)"]
                ].values[0]

                figB.add_trace(go.Bar(
                    name=param,
                    y=["As Designed", "User Selected"],
                    x=values,
                    text=[f"{v:,}" for v in values],   # comma separated inside bars
                    texttemplate='%{x:.2~s}', # Use the same format specifier for data labels
                    textposition='inside',
                    orientation='h'
                ))
            # fig.update_traces(
            #     texttemplate='%{x:.2~s}', # Use the same format specifier for data labels
            #     textposition='outside'
            # )

            figB.update_layout(
                barmode='stack',
                legend=dict(
                    orientation="h",
                    yanchor="top",
                    y=-0.25,
                    xanchor="center",
                    x=0.5
                ),
                xaxis=dict(
                    title="Energy (kWh)",
                    # tickformat=",",   # comma separated on axis
                ),
                yaxis_title=""
            )

            ###############################################
            #################### EUI ######################
            ###############################################

            st.markdown("<h5 style='text-align:left; color:red; font-weight:600;'>Energy Use Distribution</h5>", unsafe_allow_html=True)
            st.plotly_chart(figEUI, use_container_width=True, key="EnergyUseDistribyion")
            st.markdown("<h5 style='text-align:left; color:red; font-weight:600;'>Energy Use Comparison</h5>", unsafe_allow_html=True)
            col1, col2 = st.columns(2)
            with col1:
                st.plotly_chart(figA, use_container_width=True)
            with col2:
                st.plotly_chart(figB, use_container_width=True)

            st.markdown("<h5 style='text-align:left; color:red; font-weight:600;'>Energy Use Summary by End Use</h5>", unsafe_allow_html=True)
            combined_df["% Saving"] = ((combined_df["As Designed (kWh)"] - combined_df["User Selected (kWh)"]) / combined_df["As Designed (kWh)"] * 100).round(2)
            # st.dataframe(combined_df.style.format("{:,}"))
            numeric_cols = combined_df.select_dtypes(include='number').columns
            styled_df = combined_df.style.format({col: "{:,}" for col in numeric_cols})
            st.dataframe(styled_df)
            st.markdown('<hr style="border:1px solid red">', unsafe_allow_html=True)

            ###############################################
            ############## Gains and Losses ###############
            ###############################################

            all_labels = sorted(set(as_designed.index).union(set(ecm.index)))
            all_labels_loss = sorted(set(as_designed_loss.index).union(set(ecm_loss.index)))
            color_map = {
                "WALL CONDUCTION": "#dfe183",     # Blue
                "ROOF CONDUCTION": "#d893c1",     # Red
                "WINDOW GLASS+FRM COND": "#87ceeb",  # Light Blue
                "WINDOW GLASS SOLAR": "#ff6347",  # Light Red / Tomato
                "DOOR CONDUCTION": "#2ecc71",     # Green
                "INTERNAL SURFACE COND": "#90ee90",  # Light Green
                "UNDERGROUND SURF COND": "#4169e1",  # Royal Blue
                "OCCUPANTS TO SPACE": "#ff7f7f",  # Soft Red
                "LIGHT TO SPACE": "#3cb371",      # Medium Sea Green
                "EQUIPMENT TO SPACE": "#add8e6",  # Light Blue
                "PROCESS TO SPACE": "#66cdaa",    # Medium Aquamarine
                "INFILTRATION": "#ff4500"         # Orange-Red
            }
            color_map_loss = {
                "WALL CONDUCTION_loss": "#dfe183",     # Blue
                "ROOF CONDUCTION_loss": "#d893c1",     # Red
                "WINDOW GLASS+FRM COND_loss": "#87ceeb",  # Light Blue
                "WINDOW GLASS SOLAR_loss": "#ff6347",  # Light Red / Tomato
                "DOOR CONDUCTION_loss": "#2ecc71",     # Green
                "INTERNAL SURFACE COND_loss": "#90ee90",  # Light Green
                "UNDERGROUND SURF COND_loss": "#4169e1",  # Royal Blue
                "OCCUPANTS TO SPACE_loss": "#ff7f7f",  # Soft Red
                "LIGHT TO SPACE_loss": "#3cb371",      # Medium Sea Green
                "EQUIPMENT TO SPACE_loss": "#add8e6",  # Light Blue
                "PROCESS TO SPACE_loss": "#66cdaa",    # Medium Aquamarine
                "INFILTRATION_loss": "#ff4500"         # Orange-Red
            }

            # --- Helper to get colors in correct order ---
            def get_colors(labels):
                return [color_map.get(l, "#cccccc") for l in labels]
            def get_colors_loss(labels):
                return [color_map_loss.get(l, "#cccccc") for l in labels]
            def clean_labels(labels):
                # Ensure output is a list of strings
                return [str(l).replace("_loss", "") for l in labels]

            # --- Create subplots ---
            fig = make_subplots(
                rows=1,
                cols=2,
                specs=[[{'type': 'domain'}, {'type': 'domain'}]],
                subplot_titles=('As Designed', 'User Selected'),
                horizontal_spacing=0.12
            )

            figLoss = make_subplots(
                rows=1,
                cols=2,
                specs=[[{'type': 'domain'}, {'type': 'domain'}]],
                subplot_titles=('As Designed', 'User Selected'),
                horizontal_spacing=0.12
            )
            fig.add_trace(go.Pie(
                labels=as_designed.index,
                values=as_designed.values,
                # name="As Designed",
                hole=0.45,  # donut hole
                textinfo='label+percent',  # show label and %
                textposition='inside',   # <-- force text inside
                textfont=dict(size=12, color="black"),
                insidetextorientation='radial',
                hovertemplate='<b>%{label}</b><br>Value: %{value}<br>Percent: %{percent}',
                marker=dict(colors=get_colors(as_designed.index), line=dict(color='white', width=2))
            ), 1, 1)
            figLoss.add_trace(go.Pie(
                labels = clean_labels(as_designed_loss.index.tolist()),
                values=as_designed_loss.values,
                # name="As Designed",
                hole=0.45,  # donut hole
                textinfo='label+percent',  # show label and %
                textposition='inside',   # <-- force text inside
                textfont=dict(size=12, color="black"),
                insidetextorientation='radial',
                hovertemplate='<b>%{label}</b><br>Value: %{value}<br>Percent: %{percent}',
                marker=dict(colors=get_colors_loss(as_designed_loss.index), line=dict(color='white', width=2))
            ), 1, 1)

            # --- Add second pie (ECM) ---
            fig.add_trace(go.Pie(
                labels=ecm.index,
                values=ecm.values,
                # name="ECM",
                hole=0.45,  # donut hole
                textinfo='label+percent',
                textposition='inside',   # <-- force text inside
                textfont=dict(size=12, color="black"),
                insidetextorientation='radial',
                hovertemplate='<b>%{label}</b><br>Value: %{value}<br>Percent: %{percent}',
                marker=dict(colors=get_colors(ecm.index), line=dict(color='white', width=2))
            ), 1, 2)
            figLoss.add_trace(go.Pie(
                labels=clean_labels(ecm_loss.index.tolist()),
                values=ecm_loss.values,
                hole=0.45,  # donut hole
                textinfo='label+percent',
                textposition='inside',   # <-- force text inside
                textfont=dict(size=12, color="black"),
                insidetextorientation='radial',
                hovertemplate='<b>%{label}</b><br>Value: %{value}<br>Percent: %{percent}',
                marker=dict(
                    colors=get_colors_loss(ecm_loss.index),
                    line=dict(color='white', width=2)
                )
            ), 1, 2)
            st.markdown("<h5 style='text-align:left; color:red; font-weight:600;'>Gains Summary</h5>", unsafe_allow_html=True)
            annotations = []
            annotations1 = []
            # Left chart time
            annotations_gain = [
                dict(
                    text="Peak Time: Jun 16, 4 PM",
                    x=0.20,
                    y=-0.15,
                    xref="paper",
                    yref="paper",
                    showarrow=False,
                    font=dict(size=12, color="black")
                ),
                dict(
                    text="Peak Time: Jun 16, 4 PM",
                    x=0.80,
                    y=-0.15,
                    xref="paper",
                    yref="paper",
                    showarrow=False,
                    font=dict(size=12, color="black")
                )
            ]
            annotations_loss = [
                dict(
                    text="Peak Time: Jan 18, 6 AM",
                    x=0.20,
                    y=-0.15,
                    xref="paper",
                    yref="paper",
                    showarrow=False,
                    font=dict(size=12, color="black")
                ),
                dict(
                    text="Peak Time: Jan 18, 6 AM",
                    x=0.80,
                    y=-0.15,
                    xref="paper",
                    yref="paper",
                    showarrow=False,
                    font=dict(size=12, color="black")
                )
            ]
            fig.update_layout(
                showlegend=True,
                legend=dict(
                    orientation="h",
                    yanchor="top",
                    y=-0.18,
                    xanchor="center",
                    x=0.5
                ),
                height=600,
                width=1200,
                margin=dict(t=120, b=160, l=40, r=40),
                annotations=annotations_gain
            )
           
            figLoss.update_layout(
                showlegend=True,
                legend=dict(
                    orientation="h",
                    yanchor="top",
                    y=-0.18,
                    xanchor="center",
                    x=0.5
                ),
                height=600,
                width=1200,
                margin=dict(t=120, b=160, l=40, r=40),
                annotations=annotations_loss
            )
            fig.update_traces(textposition="inside", automargin=True)
            figLoss.update_traces(textposition="inside", automargin=True)
            st.plotly_chart(fig)

            if not negative_params.empty:
                neg_list = negative_params['Parameters'].tolist()
                st.markdown(f""" ⚠️ **Note:** The above charts consider only Cooling Load for Peak Sizing Period. Losses components such as {', '.join(map(str, neg_list))} were ignored.""")
            else:
                st.markdown(f"""⚠️ **Note:** The above charts consider only Cooling Load for Peak Sizing Period..""")
            
            st.markdown("<h5 style='text-align:left; color:red; font-weight:600;'>Losses Summary</h5>", unsafe_allow_html=True)
            st.plotly_chart(figLoss)
            if not positive_params.empty:
                pos_list = [i.replace('_loss','') for i in positive_params.index]
                st.markdown(
                    f"""⚠️ **Note:** The above charts consider only Heating Load for Peak Sizing Period.""")
            else:
                st.markdown(f"""⚠️ **Note:** The above charts consider only only Heating Load for Peak Sizing Period.""")
            
            st.markdown(
                "<h5 style='text-align:left; color:red; font-weight:600;'>Building Peak Load Components of Gains and Losses</h5>", unsafe_allow_html=True)
            m_df_all_loss = m_df_all_loss.rename(columns={
                'Parameters': 'Parameters ECM',
                'As Designed (kW)': 'As Designed ECM (kW)',
                'User Selected (kW)': 'User Selected ECM (kW)',
                '% Saving': '% Saving ECM'
            })
            result = pd.concat([m_df_all, m_df_all_loss], axis=1)
            result = result.drop(columns=['Parameters ECM'])
            result = result.rename(columns={
                'As Designed (kW)': 'As Designed',
                'User Selected (kW)': 'User Selected',
                '% Saving': 'Saving (%)',
                'As Designed ECM (kW)': 'As Designed',
                'User Selected ECM (kW)': 'User Selected',
                '% Saving ECM': 'Saving (%)'
            })
            result.columns = pd.MultiIndex.from_tuples([
                ('', 'Parameters'),
                ('Cooling (kW)', 'As Designed'),
                ('Cooling (kW)', 'User Selected'),
                ('Cooling (kW)', 'Saving (%)'),
                ('Heating (kW)', 'As Designed'),
                ('Heating (kW)', 'User Selected'),
                ('Heating (kW)', 'Saving (%)')
            ])
            st.dataframe(result)

            wwr_0 = (row0.get("WWR", 0))
            wwr_1 = (row1.get("WWR", 0))
            shgc_0 = (row0.get("SHGC", 0))
            shgc_1 = (row1.get("SHGC", 0))
            r_wall_0 = (row0.get("R-VAL-W", 0))
            r_wall_1 = (row1.get("R-VAL-W", 0))
            r_roof_0 = (row0.get("R-VAL-R", 0))
            r_roof_1 = (row1.get("R-VAL-R", 0))
            r_wind_0 = (row0.get("R-VAL-Wind", 0))
            r_wind_1 = (row1.get("R-VAL-Wind", 0))
            ener_0 = (row0.get("Energy_Outcome(KWH)", 0))
            ener_1 = (row1.get("Energy_Outcome(KWH)", 0))
            light_0 = (row0.get("Light(W/Sqft)", 0))
            light_1 = (row1.get("Light(W/Sqft)", 0))
            equip_0 = (row0.get("Equip(W/Sqft)", 0))
            equip_1 = (row1.get("Equip(W/Sqft)", 0))
            
            all_figs_com = []
            st.session_state.all_figs_com.extend([figEUI, figA, figB, fig, figLoss])
    
    if st.button("Generate Report"):
        if not st.session_state.all_figs_com:
            st.error("Please run Simulate first.")
            st.stop()

        with st.spinner("Generating Report..."):
            img_paths = save_figs_as_images_2(st.session_state.all_figs_com)
            cz = ashrae_zone1
            if isinstance(cz, str) and "Ext." in cz:
                cz = cz.replace("Ext.", "Extremely")
            cz = f"{cz}"
            project_info = {
                "project_name": project_name1 or "—",
                "country": selected_country1,
                "city": user_input1.title() if user_input1 else "—",
                "climate_zone": cz,
                "typology": selected_typology1,
                "weather": weather_path1,
            }
            values = {
                "wwr_0": st.session_state.get("wwr_0"),
                "wwr_1": st.session_state.get("wwr_1"),
                "shgc_0": st.session_state.get("shgc_0"),
                "shgc_1": st.session_state.get("shgc_1"),
                "r_wall_0": st.session_state.get("r_wall_0"),
                "r_wall_1": st.session_state.get("r_wall_1"),
                "r_roof_0": st.session_state.get("r_roof_0"),
                "r_roof_1": st.session_state.get("r_roof_1"),
                "r_wind_0": st.session_state.get("r_wind_0"),
                "r_wind_1": st.session_state.get("r_wind_1"),
                "ener_0": st.session_state.get("ener_0"),
                "ener_1": st.session_state.get("ener_1"),
                "light_0": st.session_state.get("light_0"),
                "light_1": st.session_state.get("light_1"),
                "equip_0": st.session_state.get("equip_0"),
                "equip_1": st.session_state.get("equip_1"),
                "climate": cz,
                "ecsbcZone":ecsbc_zone1
            }
            
            create_com_pdf(image_paths=img_paths, project_info=project_info,values=values)
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