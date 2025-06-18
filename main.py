import os
import re
import shutil
import tempfile
import io
import base64
import subprocess
from typing import List, Optional
from fastapi import FastAPI, File, UploadFile, HTTPException, Form
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
import uvicorn
from pypdf import PdfWriter, PdfReader
import docx
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet

app = FastAPI(title="PDF and DOCX Merger")

# Create directories if they don't exist
os.makedirs("uploads", exist_ok=True)
os.makedirs("static", exist_ok=True)

# Mount static files
app.mount("/static", StaticFiles(directory="static"), name="static")

# Helper function to extract part number from filename
def extract_part_number(filename: str) -> int:
    """Extract part number from filename (e.g., 'file_part1.pdf' -> 1)"""
    match = re.search(r'part(\d+)', filename.lower())
    if match:
        return int(match.group(1))
    return 0  # Default to 0 if no part number found

# Helper function to sort files by part number
def sort_files_by_part(files: List[UploadFile]) -> List[UploadFile]:
    """Sort files by their part number in the filename"""
    return sorted(files, key=lambda file: extract_part_number(file.filename))

# Helper function to merge PDF files
def merge_pdf_files(file_paths: List[str], output_path: str) -> None:
    """Merge multiple PDF files into a single PDF"""
    merger = PdfWriter()
    
    for file_path in file_paths:
        reader = PdfReader(file_path)
        for page in reader.pages:
            merger.add_page(page)
    
    with open(output_path, "wb") as output_file:
        merger.write(output_file)

# Helper function to convert DOCX to PDF
def convert_docx_to_pdf(docx_path: str, pdf_path: str) -> None:
    """Convert a DOCX file to PDF using ReportLab"""
    # Load the DOCX document
    doc = docx.Document(docx_path)
    
    # Create a PDF document
    pdf = SimpleDocTemplate(pdf_path, pagesize=letter)
    styles = getSampleStyleSheet()
    flowables = []
    
    # Process each paragraph in the DOCX
    for para in doc.paragraphs:
        if para.text.strip():
            # Determine style based on paragraph style
            style_name = 'Normal'
            if para.style.name.startswith('Heading'):
                style_name = 'Heading1'
            
            # Create a paragraph with the appropriate style
            p = Paragraph(para.text, styles[style_name])
            flowables.append(p)
            flowables.append(Spacer(1, 0.2 * inch))
    
    # Process tables (simplified - tables are complex to convert perfectly)
    for table in doc.tables:
        for row in table.rows:
            row_text = ' | '.join([cell.text for cell in row.cells])
            if row_text.strip():
                p = Paragraph(row_text, styles['Normal'])
                flowables.append(p)
                flowables.append(Spacer(1, 0.2 * inch))
    
    # Build the PDF
    pdf.build(flowables)
    
    # If the PDF is empty (no flowables), create a simple PDF with a message
    if not flowables:
        c = canvas.Canvas(pdf_path, pagesize=letter)
        c.drawString(1 * inch, 10 * inch, "No content found in the document.")
        c.save()

# Helper function to merge DOCX files
def merge_docx_files_custom(file_paths: List[str], output_path: str) -> None:
    """Merge multiple DOCX files into a single DOCX"""
    # Start with the first document as the base
    if not file_paths:
        return
    
    # Use the first document as a base and then append others with page breaks
    first_doc = docx.Document(file_paths[0])
    
    # Process each additional document
    for i, file_path in enumerate(file_paths[1:], 1):  # Start from the second document
        # Add a section break to ensure new document starts on a new page
        first_doc.add_section()
        section = first_doc.sections[-1]
        section.start_type = 2  # New page section break
        
        # Load the document to append
        doc = docx.Document(file_path)
        
        # Append all paragraphs from the document
        for para in doc.paragraphs:
            # Create a new paragraph with the same style
            p = first_doc.add_paragraph()
            p.style = para.style
            
            # Copy all runs with their formatting
            for run in para.runs:
                r = p.add_run(run.text)
                r.bold = run.bold
                r.italic = run.italic
                r.underline = run.underline
                if run.font.size:
                    r.font.size = run.font.size
                if run.font.color.rgb:
                    r.font.color.rgb = run.font.color.rgb
        
        # Copy tables
        for table in doc.tables:
            # Create a new table with the same dimensions
            tbl = first_doc.add_table(rows=len(table.rows), cols=len(table.columns))
            
            # Copy cell contents
            for i, row in enumerate(table.rows):
                for j, cell in enumerate(row.cells):
                    if i < len(tbl.rows) and j < len(tbl.rows[i].cells):
                        # Copy cell text
                        tbl.rows[i].cells[j].text = cell.text
    
    # Save the combined document
    first_doc.save(output_path)

@app.get("/", response_class=HTMLResponse)
async def get_upload_page():
    """Return the HTML upload page"""
    html_content = """
    <!DOCTYPE html>
    <html>
    <head>
        <title>PDF and DOCX Merger</title>
        <style>
            body {
                font-family: Arial, sans-serif;
                max-width: 800px;
                margin: 0 auto;
                padding: 20px;
                line-height: 1.6;
            }
            h1 {
                color: #333;
                text-align: center;
            }
            .container {
                background-color: #f9f9f9;
                border-radius: 8px;
                padding: 20px;
                box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            }
            .form-group {
                margin-bottom: 15px;
            }
            label {
                display: block;
                margin-bottom: 5px;
                font-weight: bold;
            }
            input[type="file"] {
                width: 100%;
                padding: 10px;
                border: 1px solid #ddd;
                border-radius: 4px;
            }
            button {
                background-color: #4CAF50;
                color: white;
                padding: 10px 15px;
                border: none;
                border-radius: 4px;
                cursor: pointer;
                font-size: 16px;
            }
            button:hover {
                background-color: #45a049;
            }
            .note {
                font-size: 0.9em;
                color: #666;
                margin-top: 5px;
            }
            .loading {
                display: none;
                text-align: center;
                margin-top: 20px;
            }
            .spinner {
                border: 4px solid rgba(0, 0, 0, 0.1);
                width: 36px;
                height: 36px;
                border-radius: 50%;
                border-left-color: #09f;
                animation: spin 1s linear infinite;
                display: inline-block;
            }
            @keyframes spin {
                0% { transform: rotate(0deg); }
                100% { transform: rotate(360deg); }
            }
        </style>
    </head>
    <body>
        <h1>PDF and DOCX Merger</h1>
        <div class="container">
            <form action="/merge/" enctype="multipart/form-data" method="post" onsubmit="showLoading()">
                <div class="form-group">
                    <label for="files">Select files to merge:</label>
                    <input type="file" name="files" id="files" multiple required accept=".pdf,.docx,.doc">
                    <p class="note">Files will be merged in order based on "part" numbers in their names (e.g., file_part1.pdf, file_part2.pdf)</p>
                </div>
                <div class="form-group">
                    <label for="output_filename">Output filename (without extension):</label>
                    <input type="text" name="output_filename" id="output_filename" value="merged_document">
                </div>
                <button type="submit">Merge Files</button>
            </form>
            <div id="loading" class="loading">
                <div class="spinner"></div>
                <p>Processing files, please wait...</p>
            </div>
        </div>

        <script>
            function showLoading() {
                document.getElementById('loading').style.display = 'block';
            }
        </script>
    </body>
    </html>
    """
    return html_content

@app.post("/merge/")
async def merge_files(
    files: List[UploadFile] = File(...),
    output_filename: str = Form("merged_document")
):
    """Merge uploaded files in order based on part numbers in their filenames"""
    if not files:
        raise HTTPException(status_code=400, detail="No files uploaded")
    
    # Check if all files are of the same type
    file_extensions = [os.path.splitext(file.filename)[1].lower() for file in files]
    
    if not all(ext in ['.pdf', '.docx', '.doc'] for ext in file_extensions):
        raise HTTPException(status_code=400, detail="Only PDF and DOCX/DOC files are supported")
    
    if len(set(file_extensions)) > 1:
        raise HTTPException(status_code=400, detail="All files must be of the same type (either all PDF or all DOCX/DOC)")
    
    # Sort files by part number
    sorted_files = sort_files_by_part(files)
    
    # Create a temporary directory to store uploaded files
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_file_paths = []
        
        # Save uploaded files to temporary directory
        for file in sorted_files:
            temp_file_path = os.path.join(temp_dir, file.filename)
            with open(temp_file_path, "wb") as buffer:
                shutil.copyfileobj(file.file, buffer)
            temp_file_paths.append(temp_file_path)
        
        # Determine file type and merge accordingly
        file_ext = os.path.splitext(sorted_files[0].filename)[1].lower()
        
        # Always create PDF as the final output
        if file_ext == '.pdf':
            # For PDF files, merge them directly
            pdf_output_path = f"uploads/{output_filename}.pdf"
            merge_pdf_files(temp_file_paths, pdf_output_path)
            output_path = pdf_output_path
        else:  # .docx or .doc
            # For DOCX files, first merge them
            docx_output_path = f"uploads/{output_filename}.docx"
            merge_docx_files_custom(temp_file_paths, docx_output_path)
            
            # Then convert to PDF using reportlab
            pdf_output_path = f"uploads/{output_filename}.pdf"
            try:
                # Convert DOCX to PDF using our custom function
                convert_docx_to_pdf(docx_output_path, pdf_output_path)
                output_path = pdf_output_path
            except Exception as e:
                # If conversion fails, return the DOCX file
                print(f"PDF conversion failed: {str(e)}")
                output_path = docx_output_path
    
    # Determine the correct filename and media type
    final_filename = os.path.basename(output_path)
    media_type = "application/pdf" if output_path.endswith(".pdf") else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    
    return FileResponse(
        path=output_path,
        filename=final_filename,
        media_type=media_type
    )

if __name__ == "__main__":
    # Get port from environment variable or default to 8000
    port = int(os.environ.get("PORT", 8000))
    # Use 0.0.0.0 for production, 127.0.0.1 for local development
    host = os.environ.get("HOST", "127.0.0.1")
    uvicorn.run("main:app", host=host, port=port, reload=True)
