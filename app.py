import streamlit as st
import docx
import openpyxl
from docx.oxml.ns import qn
from io import BytesIO


def get_all_paragraph_elements_part(part_element):
    for node in part_element.iter():
        if node.tag.endswith('}p'):
            yield node

def get_text_nodes_from_p(p_elem):
    t_nodes = []
    for node in p_elem.iter():
        if node.tag.endswith('}t'):
            t_nodes.append(node)
    return t_nodes

def paragraph_has_underline(p_elem):
    for r in p_elem.iter():
        if r.tag.endswith('}r'):
            for rpr in r.iter():
                if rpr.tag.endswith('}rPr'):
                    for u in rpr.iter():
                        if u.tag.endswith('}u'):
                            val = u.get(qn('w:val'))
                            if val is None or val.lower() != 'none':
                                return True
    return False

def replace_in_part_xml(part_element, old, new):
    replaced_count = 0
    for p in get_all_paragraph_elements_part(part_element):
        t_nodes = get_text_nodes_from_p(p)
        if not t_nodes:
            continue

        full = ''.join((t.text or '') for t in t_nodes)
        if old not in full:
            continue

        under = paragraph_has_underline(p)
        replacement = full.replace(old, new.upper()) if under else full.replace(old, new)

        t_nodes[0].text = replacement
        for t in t_nodes[1:]:
            t.text = ""

        replaced_count += 1

    return replaced_count

def replace_everywhere_doc(doc, old, new):
    total = 0
    total += replace_in_part_xml(doc.element.body, old, new)

    for section in doc.sections:
        try:
            total += replace_in_part_xml(section.header._element, old, new)
        except:
            pass
        try:
            total += replace_in_part_xml(section.footer._element, old, new)
        except:
            pass

    return total


# =============== STREAMLIT UI ===============

st.title(" SWR GENERATION TOOl")
st.write("Upload Excel FILE & Word file.")

excel_file = st.file_uploader("📁 Upload Excel (.xlsx)", type=["xlsx"])
word_file = st.file_uploader("📄 Upload Word (.docx)", type=["docx"])

if excel_file and word_file:
    if st.button(" GENERATE SWRD"):
        wb = openpyxl.load_workbook(excel_file, data_only=True)
        ws = wb.active

        doc = docx.Document(word_file)

        replacements = []
        for row in range(2, ws.max_row + 1):
            old = ws.cell(row, 1).value
            new = ws.cell(row, 2).value
            if old and new:
                replacements.append((str(old), str(new)))

        total_changes = 0
        for old, new in replacements:
            total_changes += replace_everywhere_doc(doc, old, new)

        output = BytesIO()
        doc.save(output)
        output.seek(0)

        st.success(f"✅ Completed! Total replacements done: {total_changes}")

        st.download_button(
            label="📥 Download Updated Word File",
            data=output,
            file_name="Updated_Document.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
