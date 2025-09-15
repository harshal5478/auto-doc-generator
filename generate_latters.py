import os
import pandas as pd
from docxtpl import DocxTemplate
from docx2pdf import convert

# ✅ Paths
excel_file = "results.xlsx"
template_file = "template.docx"
output_folder = "output"

# ✅ Fix: If "output" exists as a file, delete it
if os.path.exists(output_folder) and not os.path.isdir(output_folder):
    os.remove(output_folder)

# ✅ Create output folder if not exists
os.makedirs(output_folder, exist_ok=True)

# ✅ Read Excel File
df = pd.read_excel(excel_file)

# ✅ Loop through each row in Excel
for index, row in df.iterrows():
    name = str(row["Name"]).strip()
    content = str(row["Content"]).strip()

    # ✅ Clean the name to make it safe for filenames
    safe_name = "".join(c for c in name if c.isalnum() or c in (" ", "_", "-")).rstrip()

    # ✅ Load Word Template
    doc = DocxTemplate(template_file)

    # ✅ Replace placeholders
    context = {
        "name": name,
        "content": content
    }
    doc.render(context)

    # ✅ Save as Word (temporary)
    temp_docx = os.path.join(output_folder, f"{safe_name}.docx")
    doc.save(temp_docx)
    print(f"✅ Created DOCX for: {name}")

# ✅ Convert all DOCX to PDF
try:
    convert(output_folder)
    print("\n✅ All PDFs generated successfully in 'output' folder!")
except Exception as e:
    print("\n⚠ PDF conversion failed for some files. DOCX files are still generated.")
    print(f"Error: {e}")

    




