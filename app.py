def generate_slide_content(api_key, image_file, title, layout_type, lang):
    try:
        genai.configure(api_key=api_key)
        # Utilisation du nom standard sans préfixe
        model = genai.GenerativeModel('gemini-1.5-flash') 
        
        num_cols = 1 if "1" in layout_type else (2 if "2" in layout_type else 3)
        prompt = f"Agis comme un expert. Analyse l'image pour un slide intitulé '{title}' en {lang}. Divise ton analyse en exactement {num_cols} blocs distincts séparés par le marqueur '---BLOCK---'."
        
        img = Image.open(image_file)
        response = model.generate_content([prompt, img])
        
        if response.text:
            blocks = [b.strip() for b in response.text.split("---BLOCK---") if b.strip()]
            while len(blocks) < num_cols: blocks.append("Analyse en cours...")
            return blocks[:num_cols]
        return ["Analyse indisponible"] * num_cols
    except Exception as e:
        st.error(f"Détail de l'erreur IA : {str(e)}")
        return [f"Erreur: {str(e)}"] * 3
