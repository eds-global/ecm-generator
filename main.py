import streamlit as st
import subprocess
import os
from src import lpd
import pandas as pd
from streamlit_lottie import st_lottie
from streamlit_card import card
from PIL import Image as PILImage
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.animation import FuncAnimation
import json
import streamlit.components.v1 as components

# Set the page configuration with additional options layout='wide',
st.set_page_config(
    page_title="ECM Generator",
    page_icon="🌟",
    layout='wide',  # Only 'centered' or 'wide' are valid options
    menu_items={                          
        'Get Help': 'https://www.example.com/help',
        'Report a bug': 'https://www.example.com/bug',
        'About': '# This is an **eQuest Utilities** application!'
    }
)

def load_lottieurl(url: str):
    r = requests.get(url)
    if r.status_code != 200:
        return None
    return r.json()
    
def set_dark_theme():
    """
    Function to set a dark theme using CSS.
    """
    # Define the HTML code with CSS for a dark theme
    html_code = """
    <style>
    .stApp {
        background-color: black;  /* Set background color to black */
        color: white;  /* Set text color to white */
    }
    .stMarkdown, .stImage, .stDataFrame, .stTable, .stTextInput, .stButton, .stSidebar {
        background-color: transparent !important; /* Make elements' background transparent */
        color: white !important;  /* Ensure text color within these elements is white */
    }
    .stButton > button {
        background-color: #333; /* Dark background for buttons */
        color: white;  /* White text for buttons */
    }
    .stSidebar {
        background-color: #222; /* Slightly lighter background for sidebar */
    }
    .stTextInput > div > input {
        background-color: #444; /* Dark background for text input */
        color: white;  /* White text for text input */
    }
    </style>
    """
    # Inject the HTML code in the Streamlit app
    st.markdown(html_code, unsafe_allow_html=True)
    
def confetti_animation():
    st.markdown(
        """
        <style>
        @keyframes confetti {
            0% { transform: translateY(0) rotate(0deg); }
            100% { transform: translateY(-100vh) rotate(360deg); }
        }
        .confetti {
            position: absolute;
            width: 10px;
            height: 10px;
            background-color: #f00;
            background-image: linear-gradient(135deg, transparent 10%, #f00 10%, #f00 20%, transparent 20%, transparent 30%, #0f0 30%, #0f0 40%, transparent 40%, transparent 50%, #00f 50%, #00f 60%, transparent 60%, transparent 70%);
            background-size: 10px 10px;
            animation: confetti 5s linear infinite;
            opacity: 0.7;
        }
        </style>
        """
    )
    st.markdown('<div class="confetti"></div>', unsafe_allow_html=True)

# Render the button with the defined style
# st.markdown(button_style, unsafe_allow_html=True)

# Define CSS style with text-shadow effect for the heading
heading_style = """
    <style>
    .heading-with-shadow {
        text-align: left;
        color: red;
        text-shadow: 0px 8px 4px rgba(255, 255, 255, 0.4);
        background-color: white;
    }
</style>
"""
st.markdown(heading_style, unsafe_allow_html=True)
def main(): 
    card_button_style = """
        <style>
        .card-button {
            width: 100%;
            padding: 20px;
            background-color: white;
            border: none;
            border-radius: 10px;
            box-shadow: 0 2px 2px rgba(0,0,0,0.2);
            transition: box-shadow 0.3s ease;
            text-align: center;
            font-size: 16px;
            cursor: pointer;
        }
        .card-button:hover {
            box-shadow: 0 8px 16px rgba(0,0,0,0.3);
        }
        </style>
    """

    st.markdown(
        """
        <style>
        body {
            background-color: #bfe1ff;  /* Set your desired background color here */
            animation: changeColor 5s infinite;
        }
        .css-18e3th9 {
            padding-top: 0rem;  /* Adjust the padding at the top */
        }
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        .viewerBadge_container__1QSob {visibility: hidden;}
        .stActionButton {margin: 5px;} /* Optional: Adjust button spacing */
        header .stApp [title="View source on GitHub"] {
            display: none;
        }
        .stApp header, .stApp footer {visibility: hidden;}
        </style>
        """,
        unsafe_allow_html=True
    )

    # Initialize session state for script_choice if it does not exist
    if 'script_choice' not in st.session_state:
        st.session_state.script_choice = "ecm"  # Set default to "about"
            
    logo_url = "https://ecm-generator-edsglobal.streamlit.app/"
    logo_image_path = "images/eQcb_142.gif"
    # col1, col2, col3 = st.columns([1,1,0.5])
    # with col1:
    #     st.image(logo_image_path, width=80)
    # with col2:
        # st.markdown("<h1 class='heading-with-shadow'>eQUEST Utilities</h1>", unsafe_allow_html=True)
    st.markdown("# :rainbow[ECM Generator]")

    on = st.toggle("Select Theme")
    if on:
        set_dark_theme()
        pass  # Do nothing
        background_image_url = "https://i.pinimg.com/originals/cf/04/e9/cf04e9530f25312133dc7f93586591ff.gif"
    # with col3:
    #     st.image("images/EDSlogo.jpg", width=120)

    st.markdown('<hr style="border:1px solid black">', unsafe_allow_html=True)
    st.markdown("""
        <style>
        .stButton button {
            height: 30px;
            width: 166px;
        }
        </style>
    """, unsafe_allow_html=True)
    
    # # Create two rows of columns with equal widths
    # col2, col3 = st.columns(2) 
    
    # # Second row of buttons
    # with col2:
    #     if st.button("About EDS", key="eds"): 
    #         st.session_state.script_choice = "eds"
    # with col3:
    #     if st.button("ECM Generator", key="ecm"):
    #         st.session_state.script_choice = "ecm"

    if st.session_state.script_choice == "ecm":
        st.markdown("""
        <h4 style="color:red;">📄 Energy Conservation Measures (ECM)</h4>
        <b>Purpose:</b> The ECM Generator is designed to read and interpret INP files, which are the primary project files used by eQuest. These files contain all the necessary data about a building's energy model, including geometry, materials, systems, and schedules.<br>
        """, unsafe_allow_html=True)
        uploaded_file = st.file_uploader("Upload an INP file", type="inp", accept_multiple_files=False)

        if uploaded_file is not None:
            # Initialize ECM sets with None values if not already in session state
            if "ecm_sets" not in st.session_state:
                st.session_state.ecm_sets = [{"Orient": None, "Wall-Type": None, "Roof-Type": None, 
                                            "Window-Type": None, "WWR": None, "LPD": None, "EPD": None}]

            # Function to add a new ECM set
            def add_new_ecm_set():
                st.session_state.ecm_sets.append({"Orient": None, "Wall-Type": None, "Roof-Type": None, 
                                                "Window-Type": None, "WWR": None, "LPD": None, "EPD": None})

            # Render all ECM sets dynamically
            for index, ecm_set in enumerate(st.session_state.ecm_sets):
                st.markdown(f"""<h7 style="color:red;">🔴 ECM Set {index + 1}</h7>""", unsafe_allow_html=True)

                # ECM Set Inputs
                col1, col2, col3, col4, col5, col6 = st.columns(6)
                
                with col1:
                    ecm_set["Orient"] = st.text_input(
                        f"Orientation (°)", 
                        key=f"orient_{index}", 
                        value="" if ecm_set["Orient"] is None else str(ecm_set["Orient"]),
                        placeholder="Enter Orientation"
                    )
                    ecm_set["Orient"] = float(ecm_set["Orient"]) if ecm_set["Orient"] else None
                
                with col2:
                    ecm_set["Wall-Type"] = st.selectbox(
                        f"Wall-Type", 
                        options=[None, "Solid_Burnt_Brick-230[ENS]", "Solid_Burnt_Brick-230_XPS-5[ENS]", "Solid_Burnt_Brick-230_XPS-10[ENS]",
                        "Solid_Burnt_Brick-230_EPS-25[ENS]", "Solid_Burnt_Brick-230_EPS-50[ENS]", "AAC_Block_Wall-200[ENS]",
                        "AAC_Block_Wall-200_EPS-25[ENS]", "AAC_Block_Wall-200_EPS-50[ENS]", "AAC_Block_Wall-200_PUF-50[ENS]",
                        "Fly_Ash_Brick-230_PUF-25[ENS]", "Fly_Ash_Brick-230_PUF-50[ENS]", "Solid_Concrete_Block-200[ENS]",
                        "Solid_Concrete_Block-200_XPS-5[ENS]", "Solid_Concrete_Block-200_EPS-15[ENS]", "Solid_Concrete_Block-200_EPS-20[ENS]",
                        "Solid_Concrete_Block-200_XPS-25[ENS]", "Solid_Concrete_Block-200_XPS-50[ENS]", "Solid_Concrete_Block-200_EPS-25[ENS]",
                        "Solid_Concrete_Block-200_EPS-50[ENS]", "Reinforce_Concrete_200[ENS]"], 
                        key=f"wall_type_{index}", 
                        format_func=lambda x: "Select Wall Type" if x is None else x
                    )
                
                with col3:
                    ecm_set["Roof-Type"] = st.selectbox(
                        f"Roof-Type", 
                        options=[None, "RF-1", "RF-2", "RF-3", "RF-4"], 
                        key=f"roof_type_{index}",
                        format_func=lambda x: "Select Roof Type" if x is None else x
                    )
                
                with col4:
                    ecm_set["Window-Type"] = st.selectbox(
                        f"Window-Type", 
                        options=[None, "WIN-1", "WIN-2"], 
                        key=f"window_type_{index}",
                        format_func=lambda x: "Select Window Type" if x is None else x
                    )
                    ecm_set["WWR"] = st.text_input(
                        f"WWR (0-1)", 
                        key=f"wwr_{index}", 
                        value="" if ecm_set["WWR"] is None else str(ecm_set["WWR"]),
                        placeholder="Enter WWR"
                    )
                    ecm_set["WWR"] = float(ecm_set["WWR"]) if ecm_set["WWR"] else None
                
                with col5:
                    ecm_set["LPD"] = st.text_input(
                        f"LPD (W/ft²)", 
                        key=f"lpd_{index}", 
                        value="" if ecm_set["LPD"] is None else str(ecm_set["LPD"]),
                        placeholder="Enter LPD"
                    )
                    ecm_set["LPD"] = float(ecm_set["LPD"]) if ecm_set["LPD"] else None
                
                with col6:
                    ecm_set["EPD"] = st.text_input(
                        f"EPD (W/ft²)", 
                        key=f"epd_{index}", 
                        value="" if ecm_set["EPD"] is None else str(ecm_set["EPD"]),
                        placeholder="Enter EPD"
                    )
                    ecm_set["EPD"] = float(ecm_set["EPD"]) if ecm_set["EPD"] else None

            # Button to generate INP file
            if st.button("Generate INP"):
                for index, ecm_set in enumerate(st.session_state.ecm_sets):
                    modified_values = {k: v for k, v in ecm_set.items() if v is not None}
                    lpd.main(uploaded_file, modified_values, index + 1)

            # # Button to add a new ECM set
            # if st.button("Add ECM Set"):
            #     add_new_ecm_set()

    elif st.session_state.script_choice == "eds":
        st.markdown("""
            <h4 style="color:red;">🌐 Overview</h4>
        """, unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("""
            Environmental Design Solutions [EDS] is a sustainability advisory firm focusing on the built environment. Since its inception in 2002,
            EDS has worked on over 800 green building and energy efficiency projects worldwide. The diverse milieu of its team of experts converges on
            climate change mitigation policies, energy efficient building design, building code development, energy efficiency policy development, energy
            simulation and green building certification.<br>
    
            EDS has extensive experience in providing sustainable solutions at both, the macro level of policy advisory and planning, as well as a micro
            level of developing standards and labeling for products and appliances. The scope of EDS projects range from international and national level
            policy and code formulation to building-level integration of energy-efficiency parameters. EDS team has worked on developing the Energy Conservation
            Building Code [ECBC] in India and supporting several other international building energy code development, training, impact assessment, and 
            implementation. EDS has the experience of data collection & analysis, benchmarking, energy savings analysis, GHG impact assessment, and developing
            large scale implementation programs.<br>
    
            EDS’ work supports the global endeavour towards a sustainable environment primarily through the following broad categories:
            - Sustainable Solutions for the Built Environment
            - Strategy Consulting for Policy & Codes, and Research
            - Outreach, Communication, Documentation, and Training
    
            """, unsafe_allow_html=True)
            st.link_button("Know More", "https://edsglobal.com", type="primary")
        with col2:
            st.image("https://images.jdmagicbox.com/comp/delhi/k8/011pxx11.xx11.180809193209.h6k8/catalogue/environmental-design-solutions-vasant-vihar-delhi-environmental-management-consultants-leuub0bjnn.jpg", width=590)
        
    elif st.session_state.script_choice == "INP Parser":
        st.markdown("""
        <h4 style="color:red;">📄 INP Parser</h4>
        <b>Purpose:</b> The INP Parser is designed to read and interpret INP files, which are the primary project files used by eQuest. These files contain all the necessary data about a building's energy model, including geometry, materials, systems, and schedules.<br>
        """, unsafe_allow_html=True)
        uploaded_file = st.file_uploader("Upload an INP file", type="inp", accept_multiple_files=False)
        
        if uploaded_file is not None:
            if st.button("Generate CSV"):
                inp_parserv01.main(uploaded_file)
                    
        path = "Animation_blue_robo.json"
        with open(path, "r") as file:
            url = json.load(file)
        with col2:
            st_lottie(url,
                  reverse=True,
                  height=310,
                  width=400,
                  speed=1,
                  loop=True,
                  quality='high',
                  )
    
        with col3:
            st.markdown("#### :rainbow: :rainbow[]")
        st.markdown("""
            <style>
                .rainbow-text {
                    background: linear-gradient(to right, red, orange, yellow, green, blue, indigo, violet);
                    -webkit-background-clip: text;
                    color: transparent;
                    font-size: 2em;
                    font-weight: bold;
                    text-align: center;
                }
                .testimonial-container {
                    display: flex;
                    flex-wrap: wrap;
                    gap: 15px;
                    justify-content: center;
                    margin: 20px 0;
                }
                .testimonial {
                    border: 1px solid #ddd;
                    padding: 15px;
                    border-radius: 10px;
                    background-color: #f9f9f9;
                    max-width: 300px;
                    width: 100%;
                }
                .testimonial h5 {
                    margin: 0 0 10px;
                    color: green;
                }
                .testimonial h3 {
                    # margin: 0 0 10px;
                    color: green;
                }
                .testimonial p {
                    margin: 0;
                    color: black;
                }
            </style>
            <h4 style="text-align: center;">What People Say About Our Tool & Website</h4>
            <div class="testimonial-container">
                <div class="testimonial">
                    <h5>Robin Jain</h4>
                    <p>This is the best eQUEST utility tool I have ever used. Highly recommended! The automation features are a game-changer.
                    I highly recommend eQuest Utilities for anyone serious about optimizing their eQUEST workflow.</p>
                </div>
                <div class="testimonial">
                    <h5>Yasir Iqbal</h4>
                    <p>Amazing tools that save a lot of time and effort. Kudos to the team! Thanks Rajeev!! </p>
                </div>
                <div class="testimonial">
                    <h5>Fareed Rahi</h4>
                    <p>The user interface is very intuitive and easy to use. Great job!</p>
                </div>
                <div class="testimonial">
                    <h5>Mayank Bhatnagar</h4>
                    <p>Fantastic support and great features. Worth every penny!</p>
                </div>
                <div class="testimonial">
                    <h5>Hisham Ahmad</h4>
                    <p>Efficient and easy to navigate. This tool has made my work much easier. 
                    I love how user-friendly and efficient the eQuest Utilities tools are. They’ve made my job much easier and more productive.</p>
                </div>
                <div class="testimonial">
                    <h5>Ashraf Khan</h4>
                    <p>I love how user-friendly and efficient the eQuest Utilities tools are. They’ve made my job much easier and more productive.</p>
                </div>
                <div class="testimonial">
                    <h5>Mukul Chaudhary</h5>
                    <p>The support and features provided by eQuest Utilities are top-notch. It's a must-have for anyone working with eQUEST. </p>
                </div>
                <div class="testimonial">
                    <h5>Md. Ahsan</h4>
                    <p>Exceptional tools and excellent customer service. eQuest Utilities has definitely exceeded my expectations.</p>
                </div>
            </div>
            """, unsafe_allow_html=True)
        with st.container():
            st.markdown("#### :rainbow[Website Visitors Count]")
            components.html("""
                <p align="center">
                    <a href="https://equest-utilities-edsglobal.streamlit.app/" target="_blank">
                        <img src="https://hitwebcounter.com/counter/counter.php?page=15322595&style=0019&nbdigits=5&type=ip&initCount=70" title="Counter Widget" alt="Visit counter For Websites" border="0" />
                    </a>
                </p>
            """, height=80)
    
    elif st.session_state.script_choice == "baselineAutomation":
        st.markdown("""
        <h4 style="color:red;">🤖 Baseline Automation</h4>
        """, unsafe_allow_html=True)
        st.markdown("""
        <b>Purpose:</b> The Baseline Automation tool assists in modifying INP files based on user-defined criteria to create baseline models for comparison.
        """, unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1:
            uploaded_inp_file = st.file_uploader("Upload an INP file", type="inp", accept_multiple_files=False)
        with col2:
            uploaded_sim_file = st.file_uploader("Upload a SIM file", type="sim", accept_multiple_files=False)
        col1, col2, col3, col4, col5 = st.columns(5)
        with col1:
            input_climate = st.selectbox("Climate Zone", options=[1, 2, 3, 4, 5, 6, 7, 8])
        with col2:
            input_building_type = st.selectbox("Building Type", options=[0, 1], format_func=lambda x: "Residential" if x == 0 else "Non-Residential")
        with col3:
            input_area = st.number_input("Enter Area (Sqft)", min_value=0.0, step=0.1)
        with col4:
            number_floor = st.number_input("Number of Floors", min_value=1, step=1)
        with col5:
            heat_type = st.selectbox("Heating Type", options=[0, 1], format_func=lambda x: "Hybrid/Fossil" if x == 0 else "Electric")
    
        if uploaded_inp_file and uploaded_sim_file:
            if st.button("Automate Baseline"):
                baselineAuto.getInp(
                    uploaded_inp_file,
                    uploaded_sim_file,
                    input_climate,
                    input_building_type,
                    input_area,
                    number_floor,
                    heat_type)
                
if __name__ == "__main__":
    main()
    
st.markdown('<hr style="border:1px solid black">', unsafe_allow_html=True)
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
   <div class="footer">
        <div class="social-media" style="flex: 2;">
            <p>&copy; 2024. All Rights Reserved</p>
            <div class="icons">
                <a href="https://twitter.com/edsglobal?lang=en" target="_blank"><i class="fab fa-twitter" style="color: #1DA1F2;"></i></a>
                <a href="https://www.facebook.com/Environmental.Design.Solutions/" target="_blank"><i class="fab fa-facebook" style="color: #4267B2;"></i></a>
                <a href="https://www.instagram.com/eds_global/?hl=en" target="_blank"><i class="fab fa-instagram" style="color: #E1306C;"></i></a>
                <a href="https://www.linkedin.com/company/environmental-design-solutions/" target="_blank"><i class="fab fa-linkedin" style="color: #0077b5;"></i></a>
            </div>
            <div class="additional-content">
                <h4>Contact Us</h4>
                <p>Email: info@edsglobal.com | Phone: +123 456 7890</p>
                <p>Follow us on social media for the latest updates and news.</p>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True
)
