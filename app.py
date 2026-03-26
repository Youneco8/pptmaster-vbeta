import streamlit as st
import pandas as pd
from pptx import Presentation
import google.generativeai as genai

# Configuration de la page
st.set_page_config(page_title="Générateur PPT IA", layout="wide")

st.title("🚀 Générateur de Présentation Intelligent")
st.info("Étape 1 : Chargez vos fichiers de base (Excel + Template PPT)")

# --- ZONE D'UPLOAD INITIALE ---
col1, col2 = st.columns(2)
with col1:
    excel_file = st.file_uploader("Fichier Excel (slide_num, action_title)", type=['xlsx'])
with col2:
    template_file = st.file_uploader("Template PowerPoint (.pptx)", type=['pptx'])

# Initialisation des variables de session (pour garder en mémoire vos choix)
if 'current_step' not in st.session_state:
    st.session_state.current_step = 0
if 'slides_data' not in st.session_state:
    st.session_state.slides_data = []

# --- TRAITEMENT SI EXCEL CHARGÉ ---
if excel_file and template_file:
    df = pd.read_excel(excel_file)
    
    # Détection simplifiée de la langue (peut être améliorée via IA plus tard)
    first_title = str(df['action_title'].iloc[0]).lower()
    # Logique simple de détection
    detected_lang = "Français" if any(word in first_title for word in ["le", "la", "les", "et", "analyse"]) else "Anglais"
    
    st.success(f"Langue détectée : **{detected_lang}** | Nombre de slides prévus : **{len(df)}**")
    
    # --- LE TUNNEL (Étape 2) ---
    current_idx = st.session_state.current_step
    
    if current_idx < len(df):
        row = df.iloc[current_idx]
        st.write("---")
        st.subheader(f"Configuration du Slide {current_idx + 1} / {len(df)}")
        st.write(f"**Titre prévu :** {row['action_title']}")
        
        # Champs de saisie pour ce slide
        img = st.file_uploader(f"Upload visuel pour Slide {current_idx + 1}", type=['png', 'jpg', 'jpeg'], key=f"img_{current_idx}")
        layout_choice = st.selectbox("Choisir la disposition", ["1 bloc de texte", "2 colonnes", "3 colonnes"], key=f"lay_{current_idx}")
        
        col_btn1, col_btn2 = st.columns(2)
        
        with col_btn1:
            if st.button("Slide Suivant ➡️"):
                if img:
                    # On enregistre les données
                    st.session_state.slides_data.append({
                        "title": row['action_title'],
                        "image": img,
                        "layout": layout_choice
                    })
                    st.session_state.current_step += 1
                    st.rerun()
                else:
                    st.warning("Veuillez uploader une image avant de continuer.")
        
        with col_btn2:
            if st.button("✅ Valider et Terminer maintenant"):
                if img:
                    st.session_state.slides_data.append({
                        "title": row['action_title'],
                        "image": img,
                        "layout": layout_choice
                    })
                st.session_state.current_step = 999 # On force la fin
                st.rerun()

    # --- ÉTAPE FINALE : GÉNÉRATION ---
    if st.session_state.current_step >= len(df) or st.session_state.current_step == 999:
        st.write("---")
        st.balloons()
        st.header("Prêt pour la génération !")
        st.write(f"Vous avez configuré {len(st.session_state.slides_data)} slides.")
        
        api_key = st.text_input("Entrez votre clé API Google Gemini pour lancer l'IA :", type="password")
        
        if st.button("🪄 Générer le PowerPoint Final"):
            st.write("Traitement en cours... (L'IA analyse vos images)")
            # Ici on ajoutera la logique python-pptx et l'appel à l'IA Gemini
            st.info("Logique de génération en cours de construction pour la prochaine étape.")
