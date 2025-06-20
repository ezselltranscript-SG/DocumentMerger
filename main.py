import os
import re
import shutil
import tempfile
import zipfile
import rarfile
import patoolib
from typing import List, Optional, Tuple
from fastapi import FastAPI, File, UploadFile, HTTPException, Form, Request
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
import uvicorn
from pypdf import PdfWriter, PdfReader
import docx

app = FastAPI(title="PDF and DOCX Merger API")

# Configurar CORS para permitir solicitudes desde cualquier origen
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Create directories if they don't exist
os.makedirs("uploads", exist_ok=True)

# Add CORS middleware to allow cross-origin requests
from fastapi.middleware.cors import CORSMiddleware

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Allows all origins
    allow_credentials=True,
    allow_methods=["*"],  # Allows all methods
    allow_headers=["*"],  # Allows all headers
)

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

# Helper function to extract files from compressed archives
def extract_compressed_file(file_path: str, extract_dir: str) -> List[str]:
    """Extract files from ZIP or RAR archive and return paths to extracted files"""
    extracted_files = []
    try:
        # Primero intentamos con zipfile
        try:
            with zipfile.ZipFile(file_path, 'r') as zip_ref:
                zip_ref.extractall(extract_dir)
                extracted_files = [os.path.join(extract_dir, name) for name in zip_ref.namelist() if not name.endswith('/')]
                return extracted_files
        except zipfile.BadZipFile:
            # No es un archivo ZIP, intentamos con RAR
            pass
        
        # Intentamos con rarfile
        try:
            rarfile.UNRAR_TOOL = 'unrar'
            with rarfile.RarFile(file_path) as rar_ref:
                rar_ref.extractall(extract_dir)
                extracted_files = [os.path.join(extract_dir, name) for name in rar_ref.namelist() if not rar_ref.getinfo(name).isdir()]
                return extracted_files
        except rarfile.NotRarFile:
            # No es un archivo RAR, intentamos con patoolib como último recurso
            pass
        
        # Último recurso: intentamos con patoolib que maneja varios formatos
        patoolib.extract_archive(file_path, outdir=extract_dir)
        for root, _, files in os.walk(extract_dir):
            for file in files:
                extracted_files.append(os.path.join(root, file))
        
        if not extracted_files:
            raise ValueError("No se pudieron extraer archivos del archivo comprimido")
            
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Error extracting archive: {str(e)}")
    
    return extracted_files

# Helper function to filter files by extension
def filter_files_by_extension(file_paths: List[str], extensions: List[str]) -> List[str]:
    """Filter files by their extensions"""
    return [path for path in file_paths if os.path.splitext(path)[1].lower() in extensions]

# Helper function to merge DOCX files
def merge_docx_files_custom(file_paths: List[str], output_path: str) -> None:
    """Merge multiple DOCX files into a single DOCX preservando formato y encabezados"""
    import io
    from docx.oxml.shared import OxmlElement
    from docx.oxml.ns import qn
    from copy import deepcopy
    
    if not file_paths:
        return
    
    # Crear un documento combinado
    combined_doc = docx.Document()
    
    # Procesar cada documento
    for i, file_path in enumerate(file_paths):
        # Extraer el nombre del archivo para preservar información de encabezado
        file_name = os.path.basename(file_path)
        doc_part = f"Part{i+1}" if "Part" not in file_name else os.path.splitext(file_name)[0]
        
        # Si no es el primer documento, agregar un salto de sección
        if i > 0:
            combined_doc.add_section()
            section = combined_doc.sections[-1]
            section.start_type = 2  # New page section break
        
        # Cargar el documento actual
        doc = docx.Document(file_path)
        
        # Preservar los estilos y propiedades del documento
        for style in doc.styles:
            try:
                if style.name not in combined_doc.styles:
                    combined_doc.styles.add_style(style.name, style.type)
            except:
                pass  # Ignorar errores si el estilo ya existe o no se puede copiar
        
        # Copiar todos los párrafos con su formato exacto
        for para in doc.paragraphs:
            # Crear un nuevo párrafo con el mismo estilo
            new_para = combined_doc.add_paragraph()
            
            # Copiar el estilo del párrafo
            if para.style:
                try:
                    new_para.style = para.style
                except:
                    pass  # Si el estilo no se puede aplicar, continuar
            
            # Copiar alineación y otras propiedades del párrafo
            if para.paragraph_format.alignment:
                new_para.paragraph_format.alignment = para.paragraph_format.alignment
            
            # Copiar todos los runs con su formato exacto
            for run in para.runs:
                new_run = new_para.add_run(run.text)
                
                # Copiar formato básico
                new_run.bold = run.bold
                new_run.italic = run.italic
                new_run.underline = run.underline
                new_run.font.name = run.font.name
                
                # Copiar tamaño de fuente
                if run.font.size:
                    new_run.font.size = run.font.size
                
                # Copiar color
                if run.font.color.rgb:
                    new_run.font.color.rgb = run.font.color.rgb
        
        # Copiar tablas con su formato
        for table in doc.tables:
            # Crear una nueva tabla con las mismas dimensiones
            new_table = combined_doc.add_table(rows=len(table.rows), cols=len(table.columns))
            
            # Intentar copiar el estilo de la tabla
            try:
                new_table.style = table.style
            except:
                pass
            
            # Copiar contenido y formato de las celdas
            for i, row in enumerate(table.rows):
                for j, cell in enumerate(row.cells):
                    if i < len(new_table.rows) and j < len(new_table.rows[i].cells):
                        # Copiar el contenido de la celda con formato
                        target_cell = new_table.rows[i].cells[j]
                        
                        # Limpiar cualquier párrafo existente en la celda destino
                        for p in target_cell.paragraphs:
                            p._element.getparent().remove(p._element)
                        
                        # Copiar todos los párrafos de la celda origen a la destino
                        for para in cell.paragraphs:
                            cell_para = target_cell.add_paragraph()
                            
                            # Copiar estilo y formato
                            try:
                                cell_para.style = para.style
                            except:
                                pass
                            
                            # Copiar runs con formato
                            for run in para.runs:
                                cell_run = cell_para.add_run(run.text)
                                cell_run.bold = run.bold
                                cell_run.italic = run.italic
                                cell_run.underline = run.underline
                                if run.font.size:
                                    cell_run.font.size = run.font.size
                                if run.font.color.rgb:
                                    cell_run.font.color.rgb = run.font.color.rgb
    
    # Save the combined document
    combined_doc.save(output_path)

@app.post("/api/merge/")
async def api_merge_files(
    file: Optional[UploadFile] = File(None),
    data: Optional[UploadFile] = File(None),
    archive: Optional[UploadFile] = File(None),
    output_filename: Optional[str] = Form("merged_document"),
    request: Request = None
):
    """API endpoint to merge files from a ZIP or RAR archive"""
    # Get the actual file from any of the possible parameter names
    actual_file = file or data or archive
    
    # Debug information
    debug_info = {
        "file_provided": file is not None,
        "data_provided": data is not None,
        "archive_provided": archive is not None,
        "output_filename": output_filename,
    }
    
    if not actual_file:
        # Try to get the file from form data directly
        form = await request.form()
        for key, value in form.items():
            if isinstance(value, UploadFile):
                actual_file = value
                debug_info["found_in_form"] = key
                break
    
    if not actual_file:
        raise HTTPException(status_code=400, detail=f"No file provided. Debug info: {debug_info}")
    
    # Asumimos que el archivo es un archivo comprimido válido y lo procesamos directamente
    # sin verificar la extensión, ya que n8n puede estar enviando archivos sin extensión
    
    # Create a temporary directory to store and extract files
    with tempfile.TemporaryDirectory() as temp_dir:
        # Save the uploaded archive with a nombre genérico para evitar problemas con extensiones
        temp_file_path = os.path.join(temp_dir, "uploaded_archive")
        with open(temp_file_path, "wb") as buffer:
            shutil.copyfileobj(actual_file.file, buffer)
        
        # Create extraction directory
        extract_dir = os.path.join(temp_dir, "extracted")
        os.makedirs(extract_dir, exist_ok=True)
        
        # Extract files from the archive
        try:
            extracted_files = extract_compressed_file(temp_file_path, extract_dir)
        except Exception as e:
            raise HTTPException(status_code=400, detail=f"Error extracting archive: {str(e)}")
        
        # Filter for PDF and DOCX files only
        valid_extensions = ['.pdf', '.docx', '.doc']
        filtered_files = filter_files_by_extension(extracted_files, valid_extensions)
        
        if not filtered_files:
            raise HTTPException(status_code=400, detail="No PDF or DOCX/DOC files found in the archive")
        
        # Check if all files are of the same type
        file_extensions = [os.path.splitext(path)[1].lower() for path in filtered_files]
        unique_extensions = set(file_extensions)
        
        if len(unique_extensions) > 1:
            raise HTTPException(status_code=400, 
                              detail="Files in the archive must be of the same type (either all PDF or all DOCX/DOC)")
        
        # Sort extracted files by part number in filename
        sorted_paths = sorted(filtered_files, key=lambda path: extract_part_number(os.path.basename(path)))
        
        # Determine file type and set output path
        file_ext = os.path.splitext(sorted_paths[0])[1].lower()
        output_path = f"uploads/{output_filename}{file_ext}"
        
        if file_ext == '.pdf':
            # For PDF files, merge them directly
            merge_pdf_files(sorted_paths, output_path)
        else:  # .docx or .doc
            # For DOCX files, merge them
            merge_docx_files_custom(sorted_paths, output_path)
        
        # Set the appropriate media type based on file extension
        if output_path.endswith(".pdf"):
            media_type = "application/pdf"
        else:  # .docx or .doc
            media_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        
        return FileResponse(
            path=output_path,
            filename=os.path.basename(output_path),
            media_type=media_type
        )

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    uvicorn.run(app, host="0.0.0.0", port=port)
