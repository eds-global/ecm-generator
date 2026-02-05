import streamlit as st
import requests
from bs4 import BeautifulSoup
import os
import zipfile
import pandas as pd
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

st.title("Weather File Generator")

url = st.text_input("Enter webpage URL with .bin files")
country_name = st.text_input("Enter Country Name (e.g., USA)", value="USA")

if st.button("Download & Generate"):
    if not url:
        st.error("Please enter a valid URL")
        st.stop()

    # Ensure correct HTTP scheme
    if url.startswith("https://"):
        url = url.replace("https://", "http://")

    st.info("Fetching webpage...")
    try:
        # Session with retries
        session = requests.Session()
        retry = Retry(connect=5, backoff_factor=1)
        adapter = HTTPAdapter(max_retries=retry)
        session.mount("http://", adapter)
        session.mount("https://", adapter)

        response = session.get(url, timeout=15)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, 'html.parser')

        # Step 1: Find all .bin links
        bin_links = []
        for link in soup.find_all('a', href=True):
            href = link['href']
            if href.endswith('.bin'):
                if not href.startswith("http"):
                    href = url.rstrip('/') + '/' + href.lstrip('/')
                href = href.replace("https://", "http://")  # force http
                bin_links.append(href)

        if not bin_links:
            st.warning("No .bin files found on this webpage.")
            st.stop()

        st.info(f"Found {len(bin_links)} .bin files. Downloading...")

        # Step 2: Download .bin files
        output_folder = "downloaded_bins"
        os.makedirs(output_folder, exist_ok=True)
        downloaded_files = []
        excel_data = []

        for link in bin_links:
            filename_with_ext = link.split('/')[-1]
            filename_no_ext = filename_with_ext.replace('.bin','')
            city = filename_no_ext.split('_', 1)[-1].replace('_', ' ')

            excel_data.append({
                "Country": country_name,
                "City": city,
                "Weather": filename_no_ext,
                "WithExt": filename_with_ext
            })

            # Download each file safely
            file_path = os.path.join(output_folder, filename_with_ext)
            try:
                r = session.get(link, timeout=20)
                with open(file_path, "wb") as f:
                    f.write(r.content)
                downloaded_files.append(file_path)
            except Exception as e:
                st.warning(f"Skipped {filename_with_ext}: {e}")

        # Step 3: Create Excel
        df = pd.DataFrame(excel_data)
        excel_filename = "weather_files.xlsx"
        df.to_excel(excel_filename, index=False)

        st.success(f"Downloaded {len(downloaded_files)} .bin files successfully!")

        # Step 4: Zip all .bin files + Excel
        zip_filename = "bins.zip"
        with zipfile.ZipFile(zip_filename, 'w') as zipf:
            for file in downloaded_files:
                zipf.write(file, os.path.basename(file))
            zipf.write(excel_filename)

        # Step 5: Provide download
        with open(zip_filename, "rb") as f:
            st.download_button(
                label="Download ZIP (with Excel)",
                data=f,
                file_name=zip_filename,
                mime="application/zip"
            )

        st.success("ZIP with Excel is ready!")

    except Exception as e:
        st.error(f"Error: {e}")
