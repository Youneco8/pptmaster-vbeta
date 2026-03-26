import streamlit as st
import pandas as pd
from pptx import Presentation
import google.generativeai as genai
from PIL import Image
import io

st.set_page_config(page_title="IA Strategy Slide Generator", layout="wide")

# --- LE CERVEAU CONSULTING ---
def generate_strategic_analysis(api_key, image_file, original_title, layout_type, lang):
    try:
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel('gemini-1.5-flash')
        
        num_cols = 1 if "1" in layout_type else (2 if "2" in layout_type else 3)
        
        # Prompt ultra-précis type Cabinet de Conseil
        prompt = f"""
        Rôle : Tu es un consultant senior en stratégie (ex-McKinsey/Bain).
        Contexte : Tu dois analyser un visuel pour une présentation de direction.
        Titre actuel du slide : "{original_title}"
        Langue : {lang}

        Tâches :
        1. REVISITE LE TITRE : Crée un "Action Title" percutant qui résume le message clé (pas juste un nom, mais une conclusion).
        2. ANALYSE STRATÉGIQUE : Analyse le graphique/tableau fourni. Identifie les tendances, les anomalies et les implications business.
        3. FORMATAGE : Divise ton analyse en exactement {num_cols} blocs distincts.
        
        Contraintes :
        - Utilise des bullet points professionnels.
        - Style : Direct, analytique, haut niveau.
        - Sépare le NOUVEAU TITRE et les BLOCS par le marqueur "---SPLIT---".
        - Sépare chaque bloc de texte par le marqueur "---BLOCK---".
        
        Structure de réponse attendue :
        [Nouveau Titre Action]
        ---SPLIT---
        [Contenu Bloc 1]
        ---BLOCK---
        [Contenu Bloc 2 (si demandé)]
        """
        
        img = Image.open(image_file)
        response = model.generate_content([prompt, img])
        
        if response.text:
            # Séparation Titre / Corps
            parts = response.text.split("---SPLIT---")
            new_title = parts[0].strip()
            content_part = parts[1] if len(parts) > 1 else response.text
            
            # Séparation des blocs
            blocks = [b.strip() for b in content_part.split("---BLOCK---") if b.strip()]
            while len(blocks) < num_cols:
                blocks.append("Analyse en cours de finalisation...")
            
            return new_title, blocks[:num_cols]
        return original_title, ["Analyse indisponible"] * num_cols
    except Exception as e:
        return original_title, [f"Erreur IA : {str(e)}"] * 3

# --- INTERFACE ---
st.title("📊 IA Strategy Slide Master")

if 'current_step' not in st.session_state: st.session_state.current_step = 0
if 'slides_data' not in st.session_state: st.session_state.slides_data = []

col_a, col_b = st.columns(2)
with col_a:
    excel_file = st.file_uploader("Excel", type=['xlsx'])
with col_b:
    template_file = st.file_uploader("Template PPTX", type=['pptx'])

if excel_file and template_file:
    df = pd.read_excel(excel_file)
    first_t = str(df['action_title'].iloc[0]).lower()
    lang = "Français" if any(w in first_t for w in ["le", "la", "analyse"]) else "Anglais"

    curr = st.session_state.current_step
    if curr < len(df):
        row = df.iloc[curr]
        st.subheader(f"Slide {curr + 1} : {row['action_title']}")
        img_ui = st.file_uploader("Image (Graph/Tableau)", type=['png', 'jpg'], key=f"i_{curr}")
        lay_ui = st.selectbox("Layout", ["1 bloc", "2 colonnes", "3 colonnes"], key=f"l_{curr}")
        
        c1, c2 = st.columns(2)
        if c1.button("Suivant") and img_ui:
            st.session_state.slides_data.append({"title": row['action_title'], "img": img_ui, "lay": lay_ui})
            st.session_state.current_step += 1
            st.rerun()
        if c2.button("Terminer"):
            if img_ui: st.session_state.slides_data.append({"title": row['action_title'], "img": img_ui, "lay": lay_ui})
            st.session_state.current_step = 999
            st.rerun()

    if st.session_state.current_step >= len(df) or st.session_state.current_step == 999:
        api_key = st.text_input("Clé API Gemini", type="password")
        if st.button("🚀 Générer la présentation stratégique") and api_key:
            prs = Presentation(template_file)
            
            for i, data in enumerate(st.session_state.slides_data):
                l_idx = 0 if "1" in data['lay'] else (1 if "2" in data['lay'] else 2)
                slide = prs.slides.add_slide(prs.slide_layouts[l_idx])
                
                # Appel IA
                with st.spinner(f"Analyse stratégique du slide {i+1}..."):
                    new_title, texts = generate_strategic_analysis(api_key, data['img'], data['title'], data['lay'], lang)
                
                # Remplissage
                slide.shapes.title.text = new_title
                
                for shape in slide.placeholders:
                    if shape.placeholder_format.type == 18: # Image
                        shape.insert_picture(data['img'])
                    if shape.placeholder_format.type in [2, 7]: # Texte
                        # On essaye de remplir les blocs un par un
                        idx_txt = shape.placeholder_format.idx - 1 # Souvent l'index suit l'image
                        # Sécurité simple pour le texte :
                        if hasattr(shape, "text"):
                            try:
                                # On prend le texte correspondant au placeholder
                                # Cette partie dépend de l'ordre dans votre masque
                                shape.text = texts.pop(0)
                            except: pass

            out = io.BytesIO()
            prs.save(out)
            st.download_button("📥 Télécharger l'analyse stratégique", out.getvalue(), "Strategie_IA.pptx")
