import streamlit as st

st.set_page_config(
    page_title="Med-Kick Dashboard App",
    page_icon="👋",
)


st.title("Dashboards")
image_path = "medkicklogo.png"
st.image(image_path, caption="", use_column_width=True)
st.sidebar.success("Select a page above.")
