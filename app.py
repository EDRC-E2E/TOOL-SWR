import streamlit as st
from docx import Document
from io import BytesIO

st.title("SWR Word Replace Tool")
st.write("Upload a Word (.docx) file and provide replacement values.")

# Upload Word file
uploaded_file = st.file_uploader("Upload Word File", type=["docx"])

# User inputs for replacement
old_text = st.text_input("Text to Replace")
new_text = st.text_input("Replace With")

def replace_everywhere(doc, old, new):
    # Replace in paragraphs
    for p in doc.paragraphs:
        if old in p.text:
            for run in p.runs:
                run.text = run.text.replace(old, new)

    # Replace in tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if old in p.text:
                        for run in p.runs:
                            run.text = run.text.replace(old, new)

# When user clicks replace
if uploaded_file and st.button("Replace Text"):
    doc = Document(uploaded_file)
    replace_everywhere(doc, old_text, new_text)

    # Save edited file to memory
    output = BytesIO()
    doc.save(output)
    output.seek(0)

    st.success("✅ Replacement completed!")

    # Download button
    st.download_button(
        label="Download Updated File",
        data=output,
        file_name="updated.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
