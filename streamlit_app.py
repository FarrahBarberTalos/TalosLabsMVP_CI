from http.client import responses
import streamlit as st
import openai
import pandas as pd
from docx import Document  # For reading and writing .docx files
from io import BytesIO  # For creating a downloadable Word file
import re

# Access the OpenAI API key securely from Streamlit secrets
openai.api_key = st.secrets["openai"]["api_key"]

# Load the template document
def load_template():
    return Document("20241118 Example Non-Material Change Memo.docx")

# Modify the template based on user input
def fill_template_with_content(template_doc, additional_content):
    for para in template_doc.paragraphs:
        if "{{additional_content}}" in para.text:  # Placeholder for added content
            para.text = para.text.replace("{{additional_content}}", additional_content)
    return template_doc

def reformat_financial_section(section_text):
    reformatted = ""
    lines = section_text.split("\n")
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        # Only add bullet points for net worth lines
        if "Net Worth" in line:
            reformatted += f"• {line}\n"
        else:
            reformatted += f"  {line}\n"
    return reformatted

def clean_text(text):
    # Basic cleanup only
    text = text.replace('*', '').replace('_', '')
    
    # Handle the financial section
    if "Updated Financial Information:" in text:
        parts = text.split("Updated Financial Information:")
        before = parts[0].strip()
        financial = parts[1].strip() if len(parts) > 1 else ""
        return f"{before}\n\nUpdated Financial Information:\n\n{reformat_financial_section(financial)}"
    
    return text.strip()

# CSS for smaller font size for filenames
st.markdown(
    """
    <style>
    .small-font {
        font-size: 12px;
        color: #555;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# Show title and description
st.title("Dominik: Your CRE Co-Pilot")
st.write("I'm designed to make CRE process management seamless. Let me know how I can help.")

# File uploader - customized with specific instructions
st.write("Please upload your personal financial statements, LP memo, and a document outlining the changes requested.")
uploaded_files = st.file_uploader(
    "Upload relevant documents", accept_multiple_files=True, type=("txt", "md", "pdf", "xlsx", "docx")
)

# Initialize a variable to store the content of the uploaded documents
additional_content = ""

if uploaded_files:
    for uploaded_file in uploaded_files:
        # Display filenames with smaller font size
        st.markdown(f"<div class='small-font'>filename: {uploaded_file.name}</div>", unsafe_allow_html=True)
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

# Prompt box for user to specify changes
if uploaded_files:
    st.write("Please include the changes required below:")
    user_changes = st.text_area("Describe the changes needed", placeholder="E.g., include property details, investment summary, changes to net worth, etc.")
else:
    user_changes = ""

# Separator for buttons
st.write("---")

# Add buttons for memo generation
col1, col2 = st.columns(2)

generate_non_material = col1.button("Generate Non-Material Change Memo")
generate_material = col2.button("Generate Material Change Memo")

if generate_non_material:
    st.info("Non-Material Change Memo generation started...")
    try:
        # Generate the response content and handle document creation
        filled_template_doc = fill_template_with_content(load_template(), additional_content)
        document_content = "\n".join([para.text for para in filled_template_doc.paragraphs])

        messages = [
            {
                "role": "user",
                "content": (
                    f"Here's a document: {document_content} \n\n---\n\n"
                    f"{user_changes.strip()}\n\n"
                    "Please include a formatted 'Background Information' section - with the title in bold - that provides a comprehensive overview from the LP memo, "
                    "the property, and the deal. Ensure that the 'Property Overview' section is presented in bullet points, similar to the 'Investment Summary' section. "
                    "For the 'Updated Financial Information' section, use proper formatting in line with the formatting used above (ie bold header), spacing, and bullet points. I want it to include: 'Previous net worth: {number from the oldest personal financial statement} on one line, then another line with 'Current net worth: {number from the newest personal financial statement}, then on another line: 'Change in assets between the personal financial statements: {include the difference in assets}"
                    "Please do not include any pleasantries in your response. Only return the required output. Also, please ensure that the numbers included, especially for the financial section, are consistent with normal text. This should include consistent formatting with previous and updated net worth figures."
                ),
            }
        ]

        response = openai.ChatCompletion.create(
            model="gpt-4o-mini",
            messages=messages,
        )
        response_content = clean_text(response['choices'][0]['message']['content'])

        # Display the generated output in the UI
        st.subheader("Generated Non-Material Change Memo")
        st.markdown(f"<div style='font-size:18px; line-height:1.6;'>{response_content}</div>", unsafe_allow_html=True)

        # Create a Word document with the response content
        output_doc = Document()
        output_doc.add_heading("Non-Material Change Memo", level=1)
        output_doc.add_paragraph(response_content)

        # Save and download the document
        buffer = BytesIO()
        output_doc.save(buffer)
        buffer.seek(0)
        st.download_button(
            label="Download Non-Material Change Memo",
            data=buffer,
            file_name="Non_Material_Change_Memo.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    except openai.error.RateLimitError as e:
        st.error("You have exceeded your OpenAI API quota. Please check your OpenAI plan and billing details.")
    except openai.error.InvalidRequestError as e:
        st.error("The specified model does not exist or you do not have access to it. Please check your OpenAI access.")
    except openai.error.OpenAIError as e:
        st.error(f"An OpenAI-related error occurred: {str(e)}")
    except Exception as e:
        st.error(f"An unexpected error occurred: {str(e)}")

elif generate_material:
    st.info("Material Change Memo generation started...")
    try:
        # Placeholder logic for Material Change Memo generation
        st.success("Material Change Memo functionality is not implemented yet.")
    except Exception as e:
        st.error(f"An error occurred: {str(e)}")

elif not user_changes.strip() and uploaded_files:
    st.warning("Please include the changes required in the prompt box above.")
elif not uploaded_files:
    st.warning("Please upload relevant files before generating the document.")