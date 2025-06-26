import os
import pdfplumber
import pandas as pd

# Set the input folder to the 'pdfs' directory inside the current working directory
INPUT_FOLDER = os.path.join(os.path.dirname(__file__), 'pdfs')
OUTPUT_FILE = 'master_table.xlsx'

def extract_line_items_from_pdf(pdf_path):
    line_items = []
    pdf_file_name = os.path.basename(pdf_path)
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            if not tables:
                continue
            table = tables[0]
            header_idx = None
            for i, row in enumerate(table):
                if row and 'Fund Source' in str(row[0]):
                    header_idx = i
                    break
            if header_idx is None:
                continue
            headers = table[header_idx]
            col_map = {h.lower().strip(): idx for idx, h in enumerate(headers) if h}
            last_budget_type = ""
            for row in table[header_idx+1:]:
                if not row or len(row) < len(headers):
                    continue
                desc = row[col_map.get('description', 0)] or ""
                # Skip summary lines and reset budget type
                if any(x in desc for x in ["Budget Type Total", "Fund Source Total", "Report Total"]):
                    last_budget_type = ""
                    continue
                # Get budget type, fill down if missing
                budget_type_cell = row[col_map.get('budget type', 1)]
                if budget_type_cell and budget_type_cell.strip():
                    last_budget_type = budget_type_cell.strip()
                budget_type = last_budget_type
                line_items.append({
                    "PDF File Name": pdf_file_name,
                    "Fund Source": row[col_map.get('fund source', 0)] or "",
                    "Budget Type": budget_type,
                    "Funding": row[col_map.get('funding', 2)] or "",
                    "Description": desc,
                    "FTE": row[col_map.get('fte s', 4)] or "",
                    "Amount": row[col_map.get('amount', 5)] or ""
                })
    return line_items

def main():
    all_items = []
    for fname in os.listdir(INPUT_FOLDER):
        if fname.lower().endswith(".pdf"):
            pdf_path = os.path.join(INPUT_FOLDER, fname)
            items = extract_line_items_from_pdf(pdf_path)
            all_items.extend(items)
    if not all_items:
        print("No line items found.")
        return
    df_new = pd.DataFrame(all_items)
    # If master_table.xlsx exists, append; else, create new
    if os.path.exists(OUTPUT_FILE):
        df_master = pd.read_excel(OUTPUT_FILE)
        df_master = pd.concat([df_master, df_new], ignore_index=True)
    else:
        df_master = df_new
    df_master.to_excel(OUTPUT_FILE, index=False)
    print(f"Extracted {len(df_new)} new line items. Master table updated: {OUTPUT_FILE}")

if __name__ == "__main__":
    main() 