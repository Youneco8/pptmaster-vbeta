import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches
import google.generativeai as genai
from PIL import Image
import io

st.set_page_config(page_title="IA Slide Generator", layout="wide")

# --- FONCTION DE GÉNÉRATION IA ---
def generate_slide_content(api_key, image_file, title, layout_type, lang):
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel('gemini-1.5-flash')
    
    # Adapter le prompt selon le nombre de colonnes
    num_cols = 1 if "1" in layout_type else (2 if "2" in layout_type else 3)
    
    prompt = f"""
    Analyse cette image pour un slide PowerPoint dont le titre est "{title}".
    La langue de rédaction doit être : {lang}.
    Génère une analyse professionnelle diviseé en exactement {num_cols} blocs de texte.
    Sépare chaque bloc par le marqueur "---BLOCK---".
    Chaque bloc doit contenir des bullet points et une petite conclusion.
    """
    
    img = Image.open(image_file)
    response = model.generate_content([prompt, img])
    return response.text.split("---BLOCK---")

# --- INTERFACE ---
st.title("🪄 Créateur de Slides Automatique")

col1, col2 = st.columns(2)
with col1:
    excel_file = st.file_uploader("1. Excel (slide_num, action_title)", type=['xlsx'])
with col2:
    template_file = st.file_uploader("2. Template PPT (.pptx)", type=['pptx'])

if 'current_step' not in st.session_state: st.session_state.current_step = 0
if 'slides_data' not in st.session_state: st.session_state.slides_data = []

if excel_file and template_file:
    df = pd.read_excel(excel_file)
    first_title = str(df['action_title'].iloc[0]).lower()
    detected_lang = "Français" if any(w in first_title for w in ["le", "la", "analyse", "projet"]) else "Anglais"
    
    st.sidebar.success(f"Langue : {detected_lang}")
    
    current_idx = st.session_state.current_step
    if current_idx < len(df):
        row = df.iloc[current_idx]
        st.subheader(f"Slide {current_idx + 1} : {row['action_title']}")
        
        img = st.file_uploader("Image du slide", type=['png', 'jpg'], key=f"img_{current_idx}")
        layout = st.selectbox("Disposition", ["1 bloc", "2 colonnes", "3 colonnes"], key=f"lay_{current_idx}")
        
        c1, c2 = st.columns(2)
        if c1.button("Suivant ➡️") and img:
            st.session_state.slides_data.append({"title": row['action_title'], "img": img, "lay": layout})
            st.session_state.current_step += 1
            st.rerun()
        if c2.button("✅ Terminer ici") and img:
            st.session_state.slides_data.append({"title": row['action_title'], "img": img, "lay": layout})
            st.session_state.current_step = 999
            st.rerun()

    # --- GÉNÉRATION FINALE ---
    if st.session_state.current_step >= len(df) or st.session_state.current_step == 999:
        api_key = st.text_input("Clé API Google Gemini", type="password")
        if st.button("Lancer la génération du PPTX") and api_key:
            prs = Presentation(template_file)
            progress = st.progress(0)
            
            for i, data in enumerate(st.session_state.slides_data):
                # Choisir le layout (0, 1 ou 2 selon votre template)
                lay_idx = 0 if "1" in data['lay'] else (1 if "2" in data['lay'] else 2)
                slide = prs.slides.add_slide(prs.slide_layouts[lay_idx])
                
                # Remplir le titre
                slide.shapes.title.text = data['title']
                
                # Appel IA pour le contenu
                texts = generate_slide_content(api_key, data['img'], data['title'], data['lay'], detected_lang)
                
                # Remplir les zones de texte (Placeholders)
                # On suppose que l'image est le 1er placeholder et les textes suivent
                for j, txt in enumerate(texts):
                    try:
                        # On cherche les zones de texte vides dans le slide
                        slide.placeholders[j+1].text = txt.strip()
                    except: pass
                
                progress.progress((i + 1) / len(st.session_state.slides_data))
            
            # Sauvegarde et téléchargement
            ppt_out = io.BytesIO()
            prs.save(ppt_out)
            st.download_button("📥 Télécharger votre présentation", data=ppt_out.getvalue(), file_name="presentation_ia.pptx")
