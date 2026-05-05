import os
import uuid
import re
import shutil
import time
import urllib.parse
from fastapi import FastAPI, File, UploadFile
from fastapi.responses import JSONResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from markitdown import MarkItDown
from pptx import Presentation
from openpyxl import load_workbook

app = FastAPI(title="MarkItDown Web UI")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Ensure directories exist
os.makedirs("static/conversions", exist_ok=True)
app.mount("/static", StaticFiles(directory="static"), name="static")

md = MarkItDown()

def _extract_pivot_tables(wb_formula, wb_data):
    """Extract pivot tables by reading displayed cell values from pivot location range."""
    from openpyxl.utils import range_boundaries
    pivots_by_sheet = {}

    for ws_f in wb_formula.worksheets:
        if not ws_f._pivots:
            continue

        ws_d = wb_data[ws_f.title]
        sheet_pivots = []

        for pivot in ws_f._pivots:
            info = {"name": pivot.name or "Pivot Table"}

            # Source range metadata
            cache = pivot.cache
            if cache and cache.cacheSource and cache.cacheSource.worksheetSource:
                src = cache.cacheSource.worksheetSource
                info["source"] = f"{src.sheet}!{src.ref}" if src.sheet else str(src.ref or "")

            # Field names from cache
            if cache and cache.cacheFields:
                info["fields"] = [cf.name for cf in cache.cacheFields if cf.name]

            # Read cell values from pivot's rendered location
            if pivot.location and pivot.location.ref:
                ref = pivot.location.ref
                min_col, min_row, max_col, max_row = range_boundaries(ref)
                rows = []
                for r in range(min_row, max_row + 1):
                    row = []
                    for c in range(min_col, max_col + 1):
                        val = ws_d.cell(r, c).value
                        row.append(str(val) if val is not None else '')
                    rows.append(row)
                info["rows"] = rows

            sheet_pivots.append(info)

        if sheet_pivots:
            pivots_by_sheet[ws_f.title] = sheet_pivots

    return pivots_by_sheet


def _escape_cell(text):
    """Escape pipe and newline for Markdown table cell."""
    return text.replace('|', '\\|').replace('\n', ' ')


def convert_excel_with_formulas(file_path):
    """Convert Excel (.xlsx) showing values, formulas, and pivot tables."""
    wb_data = load_workbook(file_path, data_only=True)
    wb_formula = load_workbook(file_path, data_only=False)
    parts = []

    # Extract pivot tables (needs both workbooks)
    pivots_by_sheet = _extract_pivot_tables(wb_formula, wb_data)

    for name in wb_formula.sheetnames:
        ws_d, ws_f = wb_data[name], wb_formula[name]
        if not ws_f.max_row or not ws_f.max_column:
            continue

        parts.append(f"## {name}\n")
        rows = []
        for r in range(1, ws_f.max_row + 1):
            row = []
            for c in range(1, ws_f.max_column + 1):
                val = ws_d.cell(r, c).value
                raw = ws_f.cell(r, c).value
                if isinstance(raw, str) and raw.startswith('='):
                    text = f"{val if val is not None else ''} (`{raw}`)"
                else:
                    text = str(val) if val is not None else ''
                row.append(_escape_cell(text))
            rows.append(row)

        if not rows:
            continue

        ncols = len(rows[0])
        parts.append('| ' + ' | '.join(rows[0]) + ' |')
        parts.append('| ' + ' | '.join(['---'] * ncols) + ' |')
        for row in rows[1:]:
            parts.append('| ' + ' | '.join((row + [''] * ncols)[:ncols]) + ' |')
        parts.append('')

        # Render pivot tables for this sheet
        if name in pivots_by_sheet:
            for pv in pivots_by_sheet[name]:
                parts.append(f"### Pivot Table: {pv['name']}\n")
                if pv.get("source"):
                    parts.append(f"**Source:** `{pv['source']}`\n")
                if pv.get("fields"):
                    parts.append(f"**Fields:** {', '.join(pv['fields'])}\n")

                pv_rows = pv.get("rows", [])
                if not pv_rows:
                    parts.append("*Pivot table detected but no cached data available.*\n")
                    continue

                ncols_pv = max(len(r) for r in pv_rows)
                # First row as header
                h_row = [_escape_cell(c) for c in (pv_rows[0] + [''] * ncols_pv)[:ncols_pv]]
                parts.append('| ' + ' | '.join(h_row) + ' |')
                parts.append('| ' + ' | '.join(['---'] * ncols_pv) + ' |')
                for row in pv_rows[1:]:
                    cells = [_escape_cell(c) for c in (row + [''] * ncols_pv)[:ncols_pv]]
                    parts.append('| ' + ' | '.join(cells) + ' |')
                parts.append('')

    wb_data.close()
    wb_formula.close()
    return '\n'.join(parts)

def cleanup_old_jobs(directory="static/conversions", max_size_mb=100):
    """Giữ thư mục conversions luôn dưới mức max_size_mb, xóa các job cũ nhất nếu vượt quá."""
    try:
        job_dirs = []
        total_size = 0
        for entry in os.scandir(directory):
            if entry.is_dir():
                dir_size = sum(f.stat().st_size for f in os.scandir(entry.path) if f.is_file())
                job_dirs.append({"path": entry.path, "mtime": entry.stat().st_mtime, "size": dir_size})
                total_size += dir_size
                
        if total_size > max_size_mb * 1024 * 1024:
            job_dirs.sort(key=lambda x: x["mtime"]) # Cũ nhất lên đầu
            for job in job_dirs:
                shutil.rmtree(job["path"], ignore_errors=True)
                total_size -= job["size"]
                # Cắt xuống còn 80% giới hạn để có khoảng trống
                if total_size <= (max_size_mb * 0.8) * 1024 * 1024:
                    break
    except Exception as e:
        print(f"Lỗi khi dọn dẹp thư mục: {e}")

@app.get("/")
def serve_index():
    return FileResponse("static/index.html")

@app.post("/api/convert")
async def convert_file(file: UploadFile = File(...)):
    # Auto clean up before processing new file
    cleanup_old_jobs()
    
    job_id = str(uuid.uuid4())
    job_dir = f"static/conversions/{job_id}"
    os.makedirs(job_dir, exist_ok=True)
    
    file_path = os.path.join(job_dir, file.filename)
    
    try:
        # Save uploaded file
        with open(file_path, "wb") as f:
            shutil.copyfileobj(file.file, f)
            
        # Excel: extract values + formulas via openpyxl
        if file.filename.lower().endswith('.xlsx'):
            markdown_text = convert_excel_with_formulas(file_path)
        else:
            result = md.convert(file_path)
            markdown_text = result.text_content
            markdown_text = re.sub(r'\bNaN\b', '', markdown_text)
            markdown_text = re.sub(r'Unnamed:\s*\d+', '', markdown_text)
        
        # PPTX specific logic to extract images and rewrite markdown
        if file.filename.lower().endswith('.pptx'):
            try:
                prs = Presentation(file_path)
                for slide in prs.slides:
                    for shape in slide.shapes:
                        if hasattr(shape, "image"):
                            safe_name = shape.name.replace(" ", "")
                            ext = shape.image.ext
                            img_filename = f"{safe_name}.{ext}"
                            img_path = os.path.join(job_dir, img_filename)
                            
                            # Save the image blob
                            with open(img_path, "wb") as img_f:
                                img_f.write(shape.image.blob)
                            
                            # URL encode the filename for markdown to handle spaces properly
                            encoded_filename = urllib.parse.quote(img_filename)
                            # Regex replace the image tag in markdown
                            pattern = rf"\]\({safe_name}\.[a-zA-Z]+\)"
                            replacement = f"](/static/conversions/{job_id}/{encoded_filename})"
                            markdown_text = re.sub(pattern, replacement, markdown_text, flags=re.IGNORECASE)
                            
            except Exception as e:
                print(f"Error extracting PPTX images: {e}")

        # Save the final markdown to a file for download
        md_filename = f"{os.path.splitext(file.filename)[0]}.md"
        md_filepath = os.path.join(job_dir, md_filename)
        with open(md_filepath, "w", encoding="utf-8") as f:
            f.write(markdown_text)
            
        return {
            "success": True,
            "markdown": markdown_text,
            "download_url": f"/static/conversions/{job_id}/{md_filename}",
            "job_id": job_id
        }
        
    except Exception as e:
        return JSONResponse(status_code=500, content={"success": False, "error": str(e)})

from typing import List

@app.post("/api/convert_batch")
async def convert_batch(files: List[UploadFile] = File(...)):
    cleanup_old_jobs()
    
    job_id = str(uuid.uuid4())
    job_dir = f"static/conversions/{job_id}"
    os.makedirs(job_dir, exist_ok=True)
    
    try:
        results_data = []
        # Process each file
        for file in files:
            # Recreate relative directory structure if needed, but for simplicity we flatten or use filenames
            # In HTML file inputs with webkitdirectory, the filename is just the basename in FastAPI.
            # So we will just use the basename. If there are duplicates, we could have issues, but let's assume unique names.
            file_path = os.path.join(job_dir, file.filename)
            with open(file_path, "wb") as f:
                shutil.copyfileobj(file.file, f)
            
            try:
                if file.filename.lower().endswith('.xlsx'):
                    markdown_text = convert_excel_with_formulas(file_path)
                else:
                    result = md.convert(file_path)
                    markdown_text = result.text_content
                    markdown_text = re.sub(r'\bNaN\b', '', markdown_text)
                    markdown_text = re.sub(r'Unnamed:\s*\d+', '', markdown_text)
                
                if file.filename.lower().endswith('.pptx'):
                    try:
                        prs = Presentation(file_path)
                        for slide in prs.slides:
                            for shape in slide.shapes:
                                if hasattr(shape, "image"):
                                    safe_name = shape.name.replace(" ", "")
                                    ext = shape.image.ext
                                    img_filename = f"{os.path.splitext(file.filename)[0]}_{safe_name}.{ext}"
                                    img_path = os.path.join(job_dir, img_filename)
                                    
                                    with open(img_path, "wb") as img_f:
                                        img_f.write(shape.image.blob)

                                    # URL encode the filename for markdown to handle spaces properly
                                    encoded_filename = urllib.parse.quote(img_filename)
                                    pattern = rf"\]\({safe_name}\.[a-zA-Z]+\)"
                                    replacement = f"](/static/conversions/{job_id}/{encoded_filename})"
                                    markdown_text = re.sub(pattern, replacement, markdown_text, flags=re.IGNORECASE)
                    except Exception as e:
                        print(f"Error extracting PPTX images in batch: {e}")

                # Save md file
                md_filename = f"{os.path.splitext(file.filename)[0]}.md"
                with open(os.path.join(job_dir, md_filename), "w", encoding="utf-8") as f:
                    f.write(markdown_text)
                    
                results_data.append({
                    "filename": file.filename,
                    "markdown": markdown_text
                })
                    
            except Exception as e:
                print(f"Error converting {file.filename}: {e}")
                
            # Remove original file from zip to save space
            os.remove(file_path)
            
        # Create ZIP archive
        zip_path = f"static/conversions/{job_id}_archive"
        shutil.make_archive(zip_path, 'zip', job_dir)
        
        return {
            "success": True,
            "results": results_data,
            "download_url": f"/static/conversions/{job_id}_archive.zip",
            "job_id": job_id
        }
        
    except Exception as e:
        return JSONResponse(status_code=500, content={"success": False, "error": str(e)})

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("app:app", host="0.0.0.0", port=8000, reload=True)
