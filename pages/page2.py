import streamlit as st


st.title("Page 2")
st.write("You are now on Page 2!")

# Button to go back to Home
if st.button("Go to Home"):
    st.experimental_set_query_params(page="Home")