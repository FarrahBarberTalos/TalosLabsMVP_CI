from http.client import responses
import streamlit as st
from openai import OpenAI
import pandas as pd
from docx import Document  # For reading and writing .docx files
from io import BytesIO  # For creating a downloadable Word file
import re

# Access the OpenAI API key securely from Streamlit secrets
client = OpenAI(api_key=st.secrets["openai"]["api_key"])

# Add custom CSS for UI
st.markdown(
    """
    <style>
        body {
            background-color: #D7DEE7;
        }
        .logo-container {
            display: flex;
            justify-content: center;
            align-items: center;
            margin-bottom: 10px;
        }
        .title {
            color: #4E81BD;
            font-size: 32px;
            font-weight: bold;
            text-align: center;
            margin-top: 10px;
            margin-bottom: 20px;
        }
        .normal-text {
            color: black;
            font-size: 16px;
            text-align: center;
            margin-top: 10px;
            margin-bottom: 20px;
        }
        .small-font {
            font-size: 12px;
            color: #555;
            margin-bottom: 10px;
        }
        .file-upload {
            margin-bottom: 20px;
        }
        .button-section {
            margin-top: 30px;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

# Center the logo above the title
with st.container():
    st.markdown("<div class='logo-container'><img src='/workspaces/TalosLabs_MVP_CRE/TalosLogo.png' width='120'></div>", unsafe_allow_html=True)

# Show title and description
st.markdown("<div class='title'>Talos Labs: Your CRE Co-Pilot</div>", unsafe_allow_html=True)
st.markdown(
    "<div class='normal-text'>I'm designed to make CRE process management seamless. Let me know how I can help.</div>",
    unsafe_allow_html=True,
)

# File uploader with instructions
st.markdown(
    "<div class='normal-text'>Please upload your personal financial statements, LP memo, and a document outlining the changes requested.</div>",
    unsafe_allow_html=True,
)
uploaded_files = st.file_uploader(
    "Upload relevant documents",
    accept_multiple_files=True,
    type=("txt", "md", "pdf", "xlsx", "docx"),
    key="file_upload",
)

# Initialize a variable to store content from uploaded files
additional_content = ""

if uploaded_files:
    for uploaded_file in uploaded_files:
        # Display filenames with smaller font size
        st.markdown(f"<div class='small-font'>Filename: {uploaded_file.name}</div>", unsafe_allow_html=True)
        # Read and process each file's contents
        if uploaded_file.type == "text/plain":
            additional_content += uploaded_file.read().decode("utf-8") + "\n"
        elif uploaded_file.type == "application/pdf":
            additional_content += "PDF content parsing needed here" + "\n"  # Placeholder for PDFs
        elif uploaded_file.type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
            doc = Document(uploaded_file)
            for para in doc.paragraphs:
                additional_content += para.text + "\n"
        elif uploaded_file.type == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
            df = pd.read_excel(uploaded_file)
            additional_content += df.to_string(index=False) + "\n"  # Convert Excel to a string
        else:
            additional_content += "Unsupported file type.\n"

# Prompt for user changes
if uploaded_files:
    st.markdown(
        "<div class='normal-text'>Please include the changes required below:</div>",
        unsafe_allow_html=True,
    )
    user_changes = st.text_area(
        "Describe the changes needed",
        placeholder="E.g., include property details, investment summary, changes to net worth, etc.",
        key="text_area",
    )
else:
    user_changes = ""

# Separator for buttons
st.write("---")

# Add buttons for memo generation
col1, col2 = st.columns(2)

with col1:
    generate_non_material = st.button("Generate Non-Material Change Memo", key="non_material_button")
with col2:
    generate_material = st.button("Generate Material Change Memo", key="material_button")

# Function to handle the memo generation
def generate_memo():
    try:
        # Combine template content with user inputs
        document_content = f"Uploaded content:\n{additional_content}\nUser changes:\n{user_changes.strip()}"
        messages = [
            {
                "role": "user",
                "content": (
                    f"{document_content}\n\n---\n\n"
                    "Please generate a memo with the following structure: "
                    "1. **Background Information**: A comprehensive overview from the LP memo in paragraph form. "
                    "2. **Property Information**: A bullet-point summary of relevant property information. "
                    "3. **Investment Summary**: Bullet-point overview of rates, terms, and financials from the LP memo. "
                    "4. **Updated Financial Information**: Previous and updated net worth figures with total asset changes."
                ),
            }
        ]

        # Call OpenAI API
        response = client.chat.completions.create(model="gpt-4", messages=messages)
        response_content = response.choices[0].message.content

        # Display output in UI
        st.subheader("Generated Non-Material Change Memo")
        st.markdown(f"<div class='normal-text'>{response_content}</div>", unsafe_allow_html=True)

        # Save output to Word document
        output_doc = Document()
        output_doc.add_heading("Non-Material Change Memo", level=1)
        output_doc.add_paragraph(response_content)

        buffer = BytesIO()
        output_doc.save(buffer)
        buffer.seek(0)
        st.download_button(
            label="Download Non-Material Change Memo",
            data=buffer,
            file_name="Non_Material_Change_Memo.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    except Exception as e:
        st.error(f"An error occurred: {str(e)}")

# Generate the memo based on button clicks
if generate_non_material:
    st.info("Generating Non-Material Change Memo...")
    generate_memo()
elif generate_material:
    st.info("Material Change Memo generation is not implemented yet.")
