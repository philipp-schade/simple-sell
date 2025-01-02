import streamlit as st
import datetime
from formatter import format_excel_file
from auth import check_password

# Set page configuration
st.set_page_config(
    page_title="SimpleSell Formatter",
    page_icon="images/favicon.png",
    layout="wide",
)

# Set images and favicon
LOGO = "images/GentleNorth.jpg"

# Add a simple authentication mechanism
check_password()

# Add the image to the app
st.image(LOGO, width=180)

st.title("SimpleSell Formatter")

uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])

if uploaded_file is not None:
    st.write("File uploaded successfully.")
    
    if st.button("Format"):
        # Save the uploaded file to a temporary location
        input_file = "uploaded_file.xlsx"
        with open(input_file, "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        # Get the current date in the format dd-mm-yyyy
        current_date = datetime.datetime.now().strftime("%d-%m-%Y")
        
        # Define the output file name with the date
        output_file = f"formatted_file_{current_date}.xlsx"
        
        # Run the formatter
        format_excel_file(input_file, output_file)
        
        # Read the formatted file
        with open(output_file, "rb") as f:
            formatted_data = f.read()
        
        # Provide a download link for the formatted file
        st.download_button(
            label="Download formatted file",
            data=formatted_data,
            file_name=output_file,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )



# Add custom CSS for the footer
footer_css = """
<style>
footer {
    position: fixed;
    bottom: 0;
    width: 100%;
    text-align: center;
    color: gray;
}
</style>
"""

# Add the footer to the app
footer_html = """<footer>
  <p>Developed by P. Schade</p>
</footer>"""

st.markdown(footer_css + footer_html, unsafe_allow_html=True)