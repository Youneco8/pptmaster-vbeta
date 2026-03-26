def generate_slide_content(api_key, image_file, title, layout_type, lang):
    try:
        genai.configure(api_key=api_key)
        # On utilise le nom de modèle complet et universel
        model = genai.GenerativeModel('gemini-1.5-flash')
        
        num_cols = 1 if "1" in layout_type else (2 if "2" in layout_type else 3)
        
        prompt = f"""
        Agis comme un analyste expert. Analyse cette image pour un slide PowerPoint.
        Titre du slide : "{title}".
        Langue de rédaction : {lang}.
        
        Instructions : 
        Génère une analyse professionnelle divisée en exactement {num_cols} blocs distincts.
        Sépare chaque bloc uniquement par le marqueur précis "---BLOCK---".
        Chaque bloc doit être concis et prêt à être inséré dans un slide (bullet points autorisés).
        """
        
        # On s'assure que l'image est bien lue
        img = Image.open(image_file)
        
        response = model.generate_content([prompt, img])
        
        if response.text:
            # On nettoie la réponse pour éviter les blocs vides
            blocks = [b.strip() for b in response.text.split("---BLOCK---") if b.strip()]
            # Si l'IA donne moins de blocs que prévu, on complète avec du vide
            while len(blocks) < num_cols:
                blocks.append("")
            return blocks[:num_cols] # On ne garde que le nombre demandé
        else:
            return ["Analyse indisponible"] * num_cols
            
    except Exception as e:
        st.error(f"Erreur technique avec l'IA : {str(e)}")
        return ["Erreur d'analyse"] * num_cols
