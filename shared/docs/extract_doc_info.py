import os
import sys
import win32com.client as win32

sys.stdout.reconfigure(encoding='utf-8')

word = win32.gencache.EnsureDispatch('Word.Application')
word.Visible = False

files = [
    (r'd:\chuan-hoa-the-thuc-workspace\Hướng dẫn 05.docx', 'HD05'),
    (r'd:\chuan-hoa-the-thuc-workspace\Nghị định 30.doc', 'ND30'),
    (r'd:\chuan-hoa-the-thuc-workspace\Phụ lục Nghị định 30.doc', 'PL_ND30')
]

out_dir = r'd:\chuan-hoa-the-thuc-workspace\shared\docs\extracted_text'
os.makedirs(out_dir, exist_ok=True)

try:
    for fpath, prefix in files:
        print(f"Opening: {fpath}")
        doc = word.Documents.Open(fpath, ReadOnly=True)
        text = doc.Content.Text
        
        txt_path = os.path.join(out_dir, f"{prefix}.txt")
        with open(txt_path, 'w', encoding='utf-8') as f:
            f.write(text)
        print(f"Saved {prefix}.txt - Length: {len(text)} chars")
        
        # Also extract tables
        tables_path = os.path.join(out_dir, f"{prefix}_tables.txt")
        with open(tables_path, 'w', encoding='utf-8') as f:
            for t_idx in range(1, doc.Tables.Count + 1):
                t = doc.Tables(t_idx)
                f.write(f"\n--- TABLE {t_idx} (Rows: {t.Rows.Count}, Cols: {t.Columns.Count}) ---\n")
                for r_idx in range(1, t.Rows.Count + 1):
                    row_cells = []
                    try:
                        row = t.Rows(r_idx)
                        for c_idx in range(1, row.Cells.Count + 1):
                            cell_text = row.Cells(c_idx).Range.Text.replace('\r\x07', '').replace('\x07', '').strip()
                            row_cells.append(cell_text)
                        f.write(" | ".join(row_cells) + "\n")
                    except Exception as ex:
                        f.write(f"[Row {r_idx} error: {ex}]\n")
        print(f"Saved {prefix}_tables.txt - {doc.Tables.Count} tables")
        doc.Close(False)
finally:
    word.Quit()

print("All extractions completed!")
