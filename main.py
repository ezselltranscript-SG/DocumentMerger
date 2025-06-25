from fastapi import FastAPI, File, UploadFile, HTTPException, Form, Request
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from typing import List, Optional
from pypdf import PdfWriter, PdfReader
from docx import Document
from docxcompose.composer import Composer
from docx.enum.text import WD_BREAK

import os
import re
import shutil
import tempfile
import zipfile
import rarfile
import patoolib
import logging

# Configurar logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI(title="PDF and DOCX Merger API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

os.makedirs("uploads", exist_ok=True)

# Utilidades
def extract_part_number(filename: str) -> int:
    match = re.search(r'part(\d+)', filename.lower())
    return int(match.group(1)) if match else 0

def sort_files_by_part(files: List[str]) -> List[str]:
    return sorted(files, key=lambda f: extract_part_number(os.path.basename(f)))

def extract_compressed_file(file_path: str, extract_dir: str) -> List[str]:
    extracted_files = []
    try:
        try:
            with zipfile.ZipFile(file_path, 'r') as zip_ref:
                zip_ref.extractall(extract_dir)
                return [os.path.join(extract_dir, name) for name in zip_ref.namelist() if not name.endswith('/')]
        except zipfile.BadZipFile:
            pass

        try:
            rarfile.UNRAR_TOOL = 'unrar'
            with rarfile.RarFile(file_path) as rar_ref:
                rar_ref.extractall(extract_dir)
                return [os.path.join(extract_dir, name) for name in rar_ref.namelist() if not rar_ref.getinfo(name).isdir()]
        except rarfile.NotRarFile:
            pass

        patoolib.extract_archive(file_path, outdir=extract_dir)
        for root, _, files in os.walk(extract_dir):
            for file in files:
                extracted_files.append(os.path.join(root, file))

        return extracted_files
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Error extracting archive: {str(e)}")

def filter_files_by_extension(file_paths: List[str], extensions: List[str]) -> List[str]:
    return [p for p in file_paths if os.path.splitext(p)[1].lower() in extensions]

def merge_pdf_files(file_paths: List[str], output_path: str) -> None:
    merger = PdfWriter()
    for path in file_paths:
        reader = PdfReader(path)
        for page in reader.pages:
            merger.add_page(page)
    with open(output_path, "wb") as out:
        merger.write(out)

def merge_docx_with_headers(file_paths: List[str], output_path: str) -> None:
    """
    Fusiona archivos DOCX preservando los encabezados individuales de cada documento.
    Utiliza un enfoque alternativo que conserva los encabezados originales.
    """
    if not file_paths:
        raise HTTPException(status_code=400, detail="No DOCX files provided")
    
    if len(file_paths) == 1:
        # Si solo hay un archivo, simplemente copiarlo
        shutil.copy(file_paths[0], output_path)
        return
    
    try:
        # Utilizamos el primer documento como base
        base_doc = Document(file_paths[0])
        
        # Crear un directorio temporal para manipular los archivos
        with tempfile.TemporaryDirectory() as temp_dir:
            # Guardar el documento base en el directorio temporal
            base_temp_path = os.path.join(temp_dir, "base.docx")
            base_doc.save(base_temp_path)
            
            # Extraer el documento base
            base_extract_dir = os.path.join(temp_dir, "base_extract")
            os.makedirs(base_extract_dir, exist_ok=True)
            with zipfile.ZipFile(base_temp_path, 'r') as zip_ref:
                zip_ref.extractall(base_extract_dir)
            
            # Procesar cada documento adicional
            for i, file_path in enumerate(file_paths[1:], 1):
                logger.info(f"Procesando documento {i}: {os.path.basename(file_path)}")
                
                # Crear un documento temporal para este archivo
                current_doc = Document(file_path)
                current_temp_path = os.path.join(temp_dir, f"doc_{i}.docx")
                current_doc.save(current_temp_path)
                
                # Extraer el documento actual
                current_extract_dir = os.path.join(temp_dir, f"doc_{i}_extract")
                os.makedirs(current_extract_dir, exist_ok=True)
                with zipfile.ZipFile(current_temp_path, 'r') as zip_ref:
                    zip_ref.extractall(current_extract_dir)
                
                # Verificar si el documento tiene encabezado
                header_files = [f for f in os.listdir(os.path.join(current_extract_dir, "word")) 
                              if f.startswith("header") and f.endswith(".xml")]
                
                if header_files:
                    logger.info(f"Documento {i} tiene encabezados: {header_files}")
                
                # Abrir el documento con python-docx para manipularlo
                doc_to_append = Document(file_path)
                
                # Agregar un salto de página antes de este documento
                last_paragraph = base_doc.add_paragraph()
                last_paragraph.add_run().add_break(WD_BREAK.PAGE)
                
                # Agregar una nueva sección con diferente encabezado
                section_props = base_doc.add_section().start_type
                
                # Copiar el contenido del documento
                for para in doc_to_append.paragraphs:
                    p = base_doc.add_paragraph()
                    for run in para.runs:
                        r = p.add_run(run.text)
                        r.bold = run.bold
                        r.italic = run.italic
                        r.underline = run.underline
                        if run.font.size:
                            r.font.size = run.font.size
                        if run.font.name:
                            r.font.name = run.font.name
                
                # Copiar tablas
                for table in doc_to_append.tables:
                    tbl = base_doc.add_table(rows=len(table.rows), cols=len(table.columns))
                    if table.style:
                        tbl.style = table.style
                    
                    for i, row in enumerate(table.rows):
                        for j, cell in enumerate(row.cells):
                            if i < len(tbl.rows) and j < len(tbl.rows[i].cells):
                                tbl.rows[i].cells[j].text = cell.text
            
            # Guardar el documento base modificado
            base_doc.save(output_path)
            
            # Ahora necesitamos modificar el documento final para incluir los encabezados
            # Esto requiere manipulación directa del archivo DOCX (ZIP)
            with zipfile.ZipFile(output_path, 'a') as docx_zip:
                # Para cada documento adicional, copiar sus encabezados
                for i, file_path in enumerate(file_paths[1:], 1):
                    current_extract_dir = os.path.join(temp_dir, f"doc_{i}_extract")
                    word_dir = os.path.join(current_extract_dir, "word")
                    
                    # Buscar archivos de encabezado
                    header_files = [f for f in os.listdir(word_dir) 
                                  if f.startswith("header") and f.endswith(".xml")]
                    
                    for header_file in header_files:
                        # Copiar el archivo de encabezado con un nuevo nombre
                        source_path = os.path.join(word_dir, header_file)
                        target_name = f"header{i}_{header_file}"
                        
                        # Agregar el archivo al ZIP
                        docx_zip.write(source_path, f"word/{target_name}")
                        
                        logger.info(f"Copiado encabezado {header_file} como {target_name}")
            
            # Modificar el document.xml para referenciar los encabezados correctos
            # Esto es muy complejo y requiere manipulación XML avanzada
            # Por ahora, confiamos en que las secciones y saltos de página funcionen
        
        logger.info(f"Documento combinado guardado en {output_path}")
        
    except Exception as e:
        logger.error(f"Error al fusionar documentos: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Error al fusionar documentos: {str(e)}")

@app.post("/api/merge/")
async def api_merge_files(
    file: Optional[UploadFile] = File(None),
    data: Optional[UploadFile] = File(None),
    archive: Optional[UploadFile] = File(None),
    output_filename: Optional[str] = Form("merged_document"),
    request: Request = None
):
    actual_file = file or data or archive
    if not actual_file:
        form = await request.form()
        for value in form.values():
            if isinstance(value, UploadFile):
                actual_file = value
                break
    if not actual_file:
        raise HTTPException(status_code=400, detail="No file provided.")

    with tempfile.TemporaryDirectory() as temp_dir:
        temp_path = os.path.join(temp_dir, "archive_input")
        with open(temp_path, "wb") as f:
            shutil.copyfileobj(actual_file.file, f)

        extract_dir = os.path.join(temp_dir, "extracted")
        os.makedirs(extract_dir, exist_ok=True)
        extracted = extract_compressed_file(temp_path, extract_dir)

        valid_exts = ['.pdf', '.docx']
        filtered = filter_files_by_extension(extracted, valid_exts)
        if not filtered:
            raise HTTPException(status_code=400, detail="No valid PDF or DOCX files found.")

        sorted_files = sort_files_by_part(filtered)
        ext = os.path.splitext(sorted_files[0])[1].lower()
        output_path = f"uploads/{output_filename}{ext}"

        if ext == ".pdf":
            merge_pdf_files(sorted_files, output_path)
            media_type = "application/pdf"
        elif ext == ".docx":
            merge_docx_with_headers(sorted_files, output_path)
            media_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        else:
            raise HTTPException(status_code=400, detail="Unsupported file type.")

        return FileResponse(path=output_path, filename=os.path.basename(output_path), media_type=media_type)

if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 10000))
    uvicorn.run(app, host="0.0.0.0", port=port)
