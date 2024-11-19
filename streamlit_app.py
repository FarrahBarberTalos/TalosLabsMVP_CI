import streamlit as st
import openai
import pandas as pd
from openai.error import RateLimitError, InvalidRequestError
from docx import Document  # For reading .docx files

# Show title and description
st.title("Dominik: Your CRE Co-Pilot")
st.write(
    "Dominik is designed to make CRE process management seamless. Let me know how I can help."
    " To use this app, you need to provide an OpenAI API key, which you can get [here](https://platform.openai.com/account/api-keys)."
)

# Ask user for their OpenAI API key via `st.text_input`
openai_api_key = st.text_input("OpenAI API Key", type="password")
if not openai_api_key:
    st.info("Please add your OpenAI API key to continue.")
else:
    # Set the OpenAI API key
    openai.api_key = openai_api_key

    # File uploader - accept text, markdown, PDF, Excel, and Word documents
    uploaded_files = st.file_uploader(
        "Choose relevant files", accept_multiple_files=True, type=("txt", "md", "pdf", "xlsx", "docx")
    )

    # Initialize a variable to store the content of the document
    document_content = ""

    if uploaded_files:
        for uploaded_file in uploaded_files:
            st.write("filename:", uploaded_file.name)
            
            # Read and process each file's contents
            if uploaded_file.type == "text/plain":
                document_content += uploaded_file.read().decode("utf-8") + "\n"
            elif uploaded_file.type == "application/pdf":
                # Placeholder for PDF content extraction
                document_content += "PDF content parsing needed here" + "\n"
            elif uploaded_file.type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
                # Use python-docx to read .docx file contents
                doc = Document(uploaded_file)
                for para in doc.paragraphs:
                    document_content += para.text + "\n"
            elif uploaded_file.type == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
                # Read Excel file content using pandas
                df = pd.read_excel(uploaded_file)
                document_content += df.to_string(index=False) + "\n"  # Convert DataFrame to a string for display
            else:
                document_content += "Unsupported file type for demo" + "\n"
            
        # st.write(document_content)  # Display document text for debugging

    # Ask user for a question
    question = st.text_area(
        "What output would you like me to produce for you? Examples of documents that I can produce are updated non-material change documents, credit memos, etc.",
        placeholder="Please will you generate a template non-material change memo based upon the input provided.",
        disabled=not uploaded_files,
    )

    if uploaded_files and question:
        messages = [
            {"role": "system", "content": "You are a helpful assistant."},
            {
                "role": "user",
                "content": f"Here's a document: {document_content} \n\n---\n\n {question}",
            }
        ]

        # generate an answer using the OpenAI API with gpt-4o-mini
        try:
            response = openai.ChatCompletion.create(
                model="gpt-4o-mini",
                messages=messages,
            )
            # Display the response
            st.write(response['choices'][0]['message']['content'])
        
        except RateLimitError:
            st.error("You have exceeded your OpenAI API quota. Please check your OpenAI plan and billing details.")
        except InvalidRequestError:
            st.error("The specified model does not exist or you do not have access to it. Please check your OpenAI access.")
