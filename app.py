import streamlit as st
from fpdf import FPDF
from docxtpl import DocxTemplate
import pandas as pd
from datetime import datetime

# Función para generar documentos de Word
def generate_word(template_path, excel_data, output_folder):
    doc = DocxTemplate(template_path)
    fecha = datetime.today().strftime("%d/%m/%y")
    generated_files = []

    for _, row in excel_data.iterrows():
        context = {
            'nombre_alumno': row['Nombre del Alumno'],
            'nota_mat': row['Mat'],
            'nota_fis': row['Fis'],
            'nota_qui': row['Qui'],
            'fecha': fecha
        }
        doc.render(context)
        output_file = f"{output_folder}/Notas_de_{row['Nombre del Alumno']}.docx"
        doc.save(output_file)
        generated_files.append(output_file)
    
    return generated_files

# Clase para crear PDFs
class PDF(FPDF):
    def header(self):
        if hasattr(self, 'document_title'):
            self.set_font('Arial', 'B', 12)
            self.cell(0, 10, self.document_title, 0, 1, 'C')
    
    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Página {self.page_no()}', 0, 0, 'C')
    
    def chapter_title(self, title, font='Arial', size=12):
        self.set_font(font, 'B', size)
        self.multi_cell(0, 10, title, 0, 1, 'L')
        self.ln(10)
    
    def chapter_body(self, body, font='Arial', size=12):
        self.set_font(font, '', size)
        self.multi_cell(0, 10, body)
        self.ln()

def generate_pdf(filename, document_title, author, chapters):
    pdf = PDF()
    pdf.document_title = document_title
    pdf.add_page()
    if author:
        pdf.set_author(author)

    for chapter in chapters:
        title, body, font, size = chapter
        pdf.chapter_title(title, font, size)
        pdf.chapter_body(body, font, size)
    
    pdf.output(filename)
    return filename

# Interfaz principal con Streamlit
def main():
    st.title("EasyFlow: Generador de Documentos y PDFs")
    
    # Cargar datos
    uploaded_excel = st.file_uploader("Sube tu archivo Excel", type=["xlsx"])
    uploaded_template = st.file_uploader("Sube tu plantilla de Word", type=["docx"])

    if uploaded_excel and uploaded_template:
        st.success("Archivos cargados correctamente.")
        df = pd.read_excel(uploaded_excel)

        st.write("Vista previa de datos cargados:")
        st.dataframe(df.head())

        # Configuración de documentos
        document_title = st.text_input("Título del PDF", "Reporte de Notas")
        author = st.text_input("Autor del PDF", "EasyFlow")
        
        generate_word_docs = st.checkbox("Generar documentos Word", value=True)
        generate_pdf_doc = st.checkbox("Generar PDF resumen", value=True)

        if st.button("Generar Documentos"):
            if generate_word_docs:
                word_files = generate_word(uploaded_template, df, ".")
                st.success("Documentos Word generados:")
                for file in word_files:
                    st.write(file)
            
            if generate_pdf_doc:
                chapters = [
                    (f"{row['Nombre del Alumno']}", f"Mat: {row['Mat']}\nFis: {row['Fis']}\nQui: {row['Qui']}", 'Arial', 12)
                    for _, row in df.iterrows()
                ]
                pdf_file = generate_pdf("Resumen_Notas.pdf", document_title, author, chapters)
                st.success(f"PDF generado: {pdf_file}")
                with open(pdf_file, "rb") as pdf:
                    st.download_button(label="Descargar PDF", data=pdf, file_name="Resumen_Notas.pdf", mime='application/pdf')

if __name__ == "__main__":
    main()
