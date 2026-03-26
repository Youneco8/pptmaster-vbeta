import streamlit as st
import pandas as pd
from pptx import Presentation
import google.generativeai as genai
from PIL import Image
import io

# 1. Configuration de base
st.set_page_config(page_title="IA PPT Maker", layout="wide")

# 2. Fonction IA ultra-stable
def generate_slide_content(api_key, image_file, title, layout_type, lang):
    try:
        genai.configure(api_key=api_key)
        # Nom de modèle le plus compatible
        model = genai.GenerativeModel('gemini-1.5-flash')
        
        num_cols = 1 if "1" in layout_type else (2 if "2" in layout_type else 3)
        prompt = f"Analyse cette image pour un slide '{title}' en {lang}. Divise en {num_cols} blocs de texte pro séparés par '---BLOCK---'."
        
        img = Image.open(image_file)
        response = model.generate_content([prompt, img])
        
        blocks = [b.strip() for b in response.text.split("---BLOCK---") if b.strip()]
        while len(blocks) < num_cols:
            blocks.append("Analyse complémentaire...")
        return blocks[:num_cols]
    except Exception as e:
        return [f"Erreur IA : {str(e)}"] * 3

# 3. Interface Utilisateur
st.title("🪄 Mon Générateur PowerPoint IA")

# Initialisation de la mémoire session
if 'current_step' not in st.session_state:
    st.session_state.current_step = 0
if 'slides_data' not in st.session_state:
    st.session_state.slides_data = []

# Zone d'Upload
st.subheader("📁 Étape 1 : Charger les documents")
col_a, col_b = st.columns(2)
with col_a:
    excel_file = st.file_uploader("Fichier Excel", type=['xlsx'])
with col_b:
    template_file = st.file_uploader("Template PPTX", type=['pptx'])

# Si les fichiers sont là, on lance la machine
if excel_file and template_file:
    df = pd.read_excel(excel_file)
    first_title = str(df['action_title'].iloc[0]).lower()
    lang = "Français" if any(w in first_title for w in ["le", "la", "les", "analyse"]) else "Anglais"
    
    st.sidebar.info(f"Langue détectée : {lang}")

    # Tunnel d'upload des images
    curr = st.session_state.current_step
    if curr < len(df):
        st.write("---")
        row = df.iloc[curr]
        st.markdown(f"### Slide {curr + 1} : **{row['action_title']}**")
        
        img_ui = st.file_uploader(f"Image pour le slide {curr+1}", type=['png', 'jpg', 'jpeg'], key=f"img_{curr}")
        lay_ui = st.selectbox("Format du slide", ["1 bloc de texte", "2 colonnes", "3 colonnes"], key=f"lay_{curr}")
        
        btn_col1, btn_col2 = st.columns(2)
        if btn_col1.button("Suivant ➡️") and img_ui:
            st.session_state.slides_data.append({"title": row['action_title'], "img": img_ui, "lay": lay_ui})
            st.session_state.current_step += 1
            st.rerun()
            
        if btn_col2.button("✅ Terminer et Générer"):
            if img_ui:
                st.session_state.slides_data.append({"title": row['action_title'], "img": img_ui, "lay": lay_ui})
            st.session_state.current_step = 999
            st.rerun()

    # Génération finale
    if st.session_state.current_step >= len(df) or st.session_state.current_step == 999:
        st.write("---")
        st.header("🚀 Création du fichier final")
        api_key_input = st.text_input("Votre clé API Gemini", type="password")
        
        if st.button("Lancer la génération") and api_key_input:
            prs = Presentation(template_file)
            progress_bar = st.progress(0)
            
            for i, data in enumerate(st.session_state.slides_data):
                # Choix du layout (0, 1 ou 2 selon votre masque)
                l_idx = 0 if "1" in data['lay'] else (1 if "2" in data['lay'] else 2)
                slide = prs.slides.add_slide(prs.slide_layouts[l_idx])
                
                # Titre
                slide.shapes.title.text = data['title']
                
                # Image dans placeholder (Type 18)
                for shape in slide.placeholders:
                    if shape.placeholder_format.type == 18:
                        shape.insert_picture(data['img'])
                        break
                
                # Texte IA
                texts = generate_slide_content(api_key_input, data['img'], data['title'], data['lay'], lang)
                body_placeholders = [s for s in slide.placeholders if s.placeholder_format.type in [2, 7]]
                for j, t in enumerate(texts):
                    if j < len(body_placeholders):
                        body_placeholders[j].text = t
                
                progress_bar.progress((i + 1) / len(st.session_state.slides_data))
            
            # Export
            ppt_buf = io.BytesIO()
            prs.save(ppt_buf)
            st.download_button("📥 Télécharger mon PPT", data=ppt_buf.getvalue(), file_name="output_ia.pptx")

else:
    st.warning("Veuillez charger vos fichiers Excel et Template ci-dessus pour commencer.")
