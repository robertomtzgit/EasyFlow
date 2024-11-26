import streamlit as st
from fpdf import FPDF
from docxtpl import DocxTemplate
import pandas as pd
from datetime import datetime
from dotenv import load_dotenv
import os
import smtplib
from email.message import EmailMessage

# Cargar variables de entorno
load_dotenv()
SMTP_USERNAME = os.getenv("smtp_username")
SMTP_PASSWORD = os.getenv("smtp_password")
SMTP_SERVER = os.getenv("smtp_server")
SMTP_PORT = int(os.getenv("smtp_port"))  # Convertir el puerto a entero
SENDER_EMAIL = os.getenv("sender_email")

# Función para enviar correo electrónico
def send_email(receiver_email, subject, body, attachments):
    try:
        msg = EmailMessage()
        msg["From"] = SENDER_EMAIL
        msg["To"] = receiver_email
        msg["Subject"] = subject
        msg.set_content(body)

        for attachment in attachments:
            with open(attachment, "rb") as file:
                msg.add_attachment(
                    file.read(),
                    maintype="application",
                    subtype="octet-stream",
                    filename=os.path.basename(attachment),
                )

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SMTP_USERNAME, SMTP_PASSWORD)
            server.send_message(msg)

        return True
    except Exception as e:
        st.error(f"❌ Error al enviar el correo: {e}")
        return False

# Función para generar documentos de Word
def generate_word(template_path, excel_data, output_folder):
    doc = DocxTemplate(template_path)
    nombre = "Roberto Martinez"
    telefono = "(385) 118 07 43"
    correo = "roberto.martinez8198@alumnos.udg.mx"
    fecha = datetime.today().strftime("%d/%m/%y")
    generated_files = []

    for _, row in excel_data.iterrows():
        context = {
            'nombre': nombre, 
            'telefono': telefono, 
            'correo': correo,
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
    st.title("📄 EasyFlow: Generador de Documentos y PDFs con Envío por Correo")
    
    # Explicación inicial
    st.write(
        """
        **EasyFlow** es una herramienta diseñada para automatizar la generación de documentos y PDFs personalizados a partir de plantillas y datos estructurados.  
        Además, puedes enviar los documentos generados por correo electrónico.  
        
        ### ¿Cómo funciona?  
        1. Sube un archivo Excel que contenga los datos necesarios.  
        2. Sube una plantilla de Word con las variables a personalizar.  
        3. Configura el título y autor del PDF si lo deseas.  
        4. Genera documentos Word personalizados y/o un PDF resumen.  
        5. Ingresa el correo del destinatario y envía los documentos.
        
        ### Requisitos para los archivos:  
        **Excel**:
        - Debe contener las columnas:  
            - `Nombre del Alumno`  
            - `Mat` (Nota de Matemáticas)  
            - `Fis` (Nota de Física)  
            - `Qui` (Nota de Química)  
        
        **Plantilla de Word**:  
        - Debe incluir las siguientes variables en formato Jinja:  
            - `{{nombre_alumno}}`, `{{nota_mat}}`, `{{nota_fis}}`, `{{nota_qui}}`, y `{{fecha}}`."
        """
    )

    # Cargar datos
    uploaded_excel = st.file_uploader("🔼 Sube tu archivo Excel", type=["xlsx"])
    uploaded_template = st.file_uploader("🔼 Sube tu plantilla de Word", type=["docx"])

    if uploaded_excel and uploaded_template:
        st.success("✅ ¡Archivos cargados correctamente!")
        try:
            df = pd.read_excel(uploaded_excel)
            required_columns = {'Nombre del Alumno', 'Mat', 'Fis', 'Qui'}
            if not required_columns.issubset(df.columns):
                st.error(f"El archivo Excel debe contener las columnas: {', '.join(required_columns)}.")
                return

            st.write("👀 **Vista previa de datos cargados:**")
            st.dataframe(df.head())

            document_title = st.text_input("Título del PDF", "Reporte de Notas")
            author = st.text_input("Autor del PDF", "EasyFlow")
            
            generate_word_docs = st.checkbox("📝 Generar documentos Word", value=True)
            generate_pdf_doc = st.checkbox("📕 Generar PDF resumen", value=True)

            generated_files = []
            if st.button("🚀 Generar Documentos"):
                if generate_word_docs:
                    word_files = generate_word(uploaded_template, df, ".")
                    generated_files.extend(word_files)
                    st.success("✅ Documentos Word generados.")
                
                if generate_pdf_doc:
                    chapters = [
                        (f"{row['Nombre del Alumno']}", f"Mat: {row['Mat']}\nFis: {row['Fis']}\nQui: {row['Qui']}", 'Arial', 12)
                        for _, row in df.iterrows()
                    ]
                    pdf_file = generate_pdf("Resumen_Notas.pdf", document_title, author, chapters)
                    generated_files.append(pdf_file)
                    st.success("✅ PDF generado.")

            # Sección de envío por correo
            if generated_files:
                st.write("### Enviar documentos por correo electrónico")
                receiver_email = st.text_input("Correo del destinatario")
                subject = st.text_input("Asunto del correo", "Documentos generados por EasyFlow")
                body = st.text_area("Mensaje", "Por favor, encuentra adjuntos los documentos generados.")
                
                if st.button("📧 Enviar Correo"):
                    if receiver_email:
                        success = send_email(receiver_email, subject, body, generated_files)
                        if success:
                            st.success("✅ ¡Correo enviado correctamente!")
                    else:
                        st.error("❌ Por favor, ingresa un correo válido.")
        except Exception as e:
            st.error(f"❌ Error al procesar los archivos: {e}")

if __name__ == "__main__":
    main()
