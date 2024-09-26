import streamlit as st
import base64
import pandas as pd
from io import BytesIO
from google.oauth2 import service_account
from google.cloud import documentai_v1 as documentai

# Configuration Google Document AI
key_json_base64 = st.secrets["GOOGLE_KEY_JSON"]  # Utiliser Streamlit secrets pour le déploiement sécurisé
key_json_content = base64.b64decode(key_json_base64).decode("utf-8")
credentials = service_account.Credentials.from_service_account_info(json.loads(key_json_content))

project_id = "74081051811"
location = "us"
ocr_processor_id = "f0108ad9f637ec0c"
form_parser_processor_id = "213655943885e363"

client = documentai.DocumentProcessorServiceClient(credentials=credentials)

# Fonction pour traiter un document avec Google Document AI (OCR)
def process_document_ocr(file_bytes, mime_type):
    name = f"projects/{project_id}/locations/{location}/processors/{ocr_processor_id}"
    document = {"content": file_bytes, "mime_type": mime_type}
    request = {"name": name, "raw_document": document}
    result = client.process_document(request=request)
    return result.document.text

# Fonction pour analyser un formulaire avec Google Document AI (Form Parsing)
def process_document_form(file_bytes, mime_type):
    name = f"projects/{project_id}/locations/{location}/processors/{form_parser_processor_id}"
    document = {"content": file_bytes, "mime_type": mime_type}
    request = {"name": name, "raw_document": document}
    result = client.process_document(request=request)
    
    form_data = {}
    for entity in result.document.entities:
        field_name = entity.field_name.text if entity.field_name else "Inconnu"
        field_value = entity.field_value.text if entity.field_value else "Inconnu"
        form_data[field_name] = field_value
    
    # Extraction des tables si présentes
    tables = []
    for page in result.document.pages:
        for table in page.tables:
            table_data = []
            for row in table.header_rows:
                table_data.append([cell.layout.text for cell in row.cells])
            for row in table.body_rows:
                table_data.append([cell.layout.text for cell in row.cells])
            tables.append(pd.DataFrame(table_data))
    
    return form_data, tables

# Fonction pour gérer le fichier uploadé et décider entre OCR et Form Parsing
def handle_uploaded_file(uploaded_file):
    file_bytes = uploaded_file.read()
    mime_type = uploaded_file.type

    if 'form' in uploaded_file.name.lower():
        extracted_data, tables = process_document_form(file_bytes, mime_type)
    else:
        extracted_data = process_document_ocr(file_bytes, mime_type)
        tables = None

    return extracted_data, tables

# Fonction pour exporter les résultats vers un fichier Excel
def to_excel(df, tables=None):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Extracted Text')
    
    if tables:
        for i, table in enumerate(tables):
            table.to_excel(writer, sheet_name=f'Table_{i+1}', index=False)
    
    writer.close()
    output.seek(0)
    return output

# Configuration de l'interface Streamlit
st.set_page_config(page_title="Document AI - OCR et Form Parsing", page_icon="📄", layout="wide")
st.title("📄 OCR et Extraction de Formulaires avec Document AI")

# Chargement de fichier
file_type = st.sidebar.selectbox("Format de téléchargement", ["TXT", "Excel"])
uploaded_file = st.file_uploader("Choisissez un fichier PDF ou image", type=["pdf", "png", "jpg", "jpeg"])

# Affichage du fichier uploadé
if uploaded_file:
    file_type_uploaded = uploaded_file.type.split('/')[0]
    if file_type_uploaded == 'image':
        st.image(uploaded_file, caption='📸 Fichier Image Téléchargé', use_column_width=True)
    elif uploaded_file.type == 'application/pdf':
        st.write("📄 Fichier PDF Téléchargé")
        
    # Bouton pour lancer l'extraction
    if st.button('💡 Extraire et Analyser'):
        try:
            with st.spinner("⚙️ Traitement en cours..."):
                extracted_data, tables = handle_uploaded_file(uploaded_file)

                if isinstance(extracted_data, dict):
                    st.subheader("📋 Données de Formulaire Extraites")
                    st.json(extracted_data)

                    if tables:
                        st.subheader("📊 Tables Détectées")
                        for idx, table in enumerate(tables, start=1):
                            st.write(f"Tableau {idx}")
                            st.dataframe(table)

                    # Option de téléchargement en Excel pour les formulaires
                    df = pd.DataFrame(extracted_data.items(), columns=["Champ", "Valeur"])
                    excel_data = to_excel(df, tables=tables)
                    st.download_button(label="💾 Télécharger le fichier Excel", data=excel_data, file_name="form_data.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                else:
                    st.subheader("📝 Texte Extrait")
                    st.text_area("Texte Extrait", extracted_data, height=300)

                    # Option de téléchargement pour le texte extrait
                    if file_type == "TXT":
                        st.download_button(label="💾 Télécharger le fichier TXT", data=extracted_data, file_name="extracted_text.txt", mime="text/plain")
                    elif file_type == "Excel":
                        df = pd.DataFrame([{'Texte Extrait': extracted_data}])
                        excel_data = to_excel(df)
                        st.download_button(label="💾 Télécharger le fichier Excel", data=excel_data, file_name="extracted_text.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"Erreur lors du traitement : {e}")

# Ajout du copyright dans la sidebar
st.sidebar.markdown("---")
st.sidebar.markdown(
    """
    <div style='text-align: center; color: grey;'>
        &copy; 2024 OCR PDF & Image by Gregory Mulamba
    </div>
    """, 
    unsafe_allow_html=True
)
