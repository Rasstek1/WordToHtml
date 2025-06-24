import streamlit as st
import tempfile
import zipfile
from io import BytesIO
import xml.etree.ElementTree as ET
from bs4 import BeautifulSoup
import re
import os

# Configuration de la page
st.set_page_config(
    page_title="Word to HTML Converter",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS personnalisé
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        color: #2c3e50;
        text-align: center;
        margin-bottom: 2rem;
        padding: 1rem;
        background: linear-gradient(90deg, #3498db, #2ecc71);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-weight: bold;
    }
    .upload-section {
        background-color: #f8f9fa;
        padding: 2rem;
        border-radius: 10px;
        border: 2px dashed #3498db;
        margin: 1rem 0;
    }
    .result-section {
        background-color: #e8f5e8;
        padding: 1.5rem;
        border-radius: 10px;
        border-left: 5px solid #2ecc71;
        margin: 1rem 0;
    }
    .info-box {
        background-color: #e3f2fd;
        padding: 1rem;
        border-radius: 8px;
        border-left: 4px solid #2196f3;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

# VOTRE CODE - Adapté pour Streamlit
def ameliorations_typographiques(soup):
    """Applique des améliorations typographiques au contenu HTML"""
    apostrophes_changees = 0
    mots_tirets_wrapes = 0
    
    # Parcourir tous les éléments textuels
    for element in soup.find_all(string=True):
        if element.parent.name not in ['script', 'style']:
            texte_original = str(element)
            texte_modifie = texte_original
            
            # 1. Remplacer les apostrophes droites par des apostrophes courbes
            apostrophes_avant = texte_modifie.count("'")
            texte_modifie = texte_modifie.replace("'", "'")
            apostrophes_changees += apostrophes_avant
            
            # 2. Identifier et protéger les mots avec tirets
            pattern_tirets = r'\b([a-zA-ZÀ-ÿ]+(?:-[a-zA-ZÀ-ÿ]+)+)\b'
            
            def wrap_mot_tiret(match):
                mot = match.group(1)
                nonlocal mots_tirets_wrapes
                mots_tirets_wrapes += 1
                return f'<span class="nowrap">{mot}</span>'
            
            texte_modifie = re.sub(pattern_tirets, wrap_mot_tiret, texte_modifie)
            
            # Remplacer le texte si il y a eu des modifications
            if texte_modifie != texte_original:
                if '<span' in texte_modifie:
                    fragment = BeautifulSoup(texte_modifie, 'html.parser')
                    element.replace_with(*fragment.contents)
                else:
                    element.replace_with(texte_modifie)
    
    return apostrophes_changees, mots_tirets_wrapes

def detecter_et_convertir_titres(soup):
    """Convertit les <p><strong> en <h2> quand c'est approprié"""
    titres_convertis = 0
    
    for p in soup.find_all('p'):
        strong_tags = p.find_all('strong')
        
        if strong_tags:
            texte_total = p.get_text().strip()
            texte_strong = ''.join([s.get_text().strip() for s in strong_tags])
            
            if (len(texte_strong) > 0 and 
                len(texte_total) > 0 and
                len(texte_strong) / len(texte_total) >= 0.8 and
                len(texte_total) <= 100 and
                not texte_total.endswith('.') and
                not texte_total.endswith(',') and
                len(texte_total.split()) <= 15):
                
                if (not any(mot in texte_total.lower() for mot in ['cliquez', 'voir', 'télécharger', 'lire', 'plus d\'info']) and
                    not re.search(r'\d+\s*%', texte_total) and
                    not re.search(r'\$\d+', texte_total)):
                    
                    h2 = soup.new_tag('h2')
                    h2.string = texte_total
                    p.replace_with(h2)
                    titres_convertis += 1
    
    return titres_convertis

def analyser_structure_document_bytes(fichier_word_bytes):
    """Analyse la structure complète du document Word pour préserver l'ordre exact"""
    elements_document = []
    
    try:
        with zipfile.ZipFile(BytesIO(fichier_word_bytes), 'r') as docx_zip:
            document_xml = docx_zip.read('word/document.xml')
            root = ET.fromstring(document_xml)
            
            namespaces = {
                'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
                'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
                'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
                'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
            }
            
            paragraphes = root.findall('.//w:p', namespaces)
            
            for i, para in enumerate(paragraphes):
                texte_runs = para.findall('.//w:t', namespaces)
                texte_paragraphe = ''.join([t.text or '' for t in texte_runs])
                
                drawings = para.findall('.//w:drawing', namespaces)
                objects = para.findall('.//w:object', namespaces)
                
                has_image = len(drawings) > 0 or len(objects) > 0
                
                elements_document.append({
                    'type': 'paragraphe',
                    'index': i,
                    'texte': texte_paragraphe.strip(),
                    'has_image': has_image,
                    'nb_images': len(drawings) + len(objects)
                })
            
            return elements_document
            
    except Exception as e:
        st.error(f"Erreur lors de l'analyse XML: {e}")
        return []

def nettoyer_images_dans_html(html_content):
    """Remplace toutes les images avec des données longues par l'image sample"""
    placeholder_path = "img_sample.jpg"  # Utilise toujours img_sample.jpg
    pattern_img_longue = r'<img[^>]*src="data:image/[^"]{100,}"[^>]*/?>'
    images_longues = re.findall(pattern_img_longue, html_content)
    
    counter = [1]
    
    def remplacer_image(match):
        img_tag = f'<img src="{placeholder_path}" alt="Image {counter[0]}" class="sample-img" style="max-width: 300px; height: auto; border: 1px solid #ddd; margin: 10px 0;" />'
        counter[0] += 1
        return img_tag
    
    html_nettoye = re.sub(pattern_img_longue, remplacer_image, html_content)
    return html_nettoye, len(images_longues)

def convertir_word_vers_html_complet(fichier_word_bytes, nom_fichier):
    """Convertit un fichier Word en HTML avec votre code complet"""
    try:
        import mammoth
        
        structure_originale = analyser_structure_document_bytes(fichier_word_bytes)
        
        with BytesIO(fichier_word_bytes) as docx_file:
            result = mammoth.convert_to_html(docx_file)
            html = result.value
        
        html_nettoye, nb_images_remplacees = nettoyer_images_dans_html(html)
        soup = BeautifulSoup(html_nettoye, 'html.parser')
        
        # ÉTAPE 1 : Détecter et convertir les titres
        titres_convertis = detecter_et_convertir_titres(soup)
        
        # ÉTAPE 2 : Améliorations typographiques
        apostrophes_changees, mots_tirets_wrapes = ameliorations_typographiques(soup)
        
        # Gestion des images manquantes
        images_html = soup.find_all('img')
        images_attendues = sum(elem['nb_images'] for elem in structure_originale if elem['has_image'])
        
        if len(images_html) < images_attendues:
            paragraphes_html = soup.find_all(['p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6'])
            placeholder_path = "img_sample.jpg"
            
            images_ajoutees = len(images_html)
            for elem_xml in structure_originale:
                if elem_xml['has_image'] and elem_xml['texte'] and len(elem_xml['texte']) > 10:
                    for para_html in paragraphes_html:
                        texte_html = para_html.get_text().strip()
                        
                        if texte_html and len(texte_html) > 10:
                            if any(mot in texte_html.lower() for mot in elem_xml['texte'][:50].lower().split() if len(mot) > 3):
                                if not para_html.find_next_sibling('img'):
                                    for _ in range(elem_xml['nb_images']):
                                        if images_ajoutees < images_attendues:
                                            img_tag = soup.new_tag('img', 
                                                                 src=placeholder_path, 
                                                                 alt=f'Image {images_ajoutees + 1}',
                                                                 class_='sample-img',
                                                                 style='max-width: 300px; height: auto; border: 1px solid #ddd; margin: 10px 0;')
                                            para_html.insert_after(img_tag)
                                            images_ajoutees += 1
                                break
        
        # Nettoyage final (votre code de nettoyage)
        for p in soup.find_all('p'):
            if not p.get_text(strip=True) and not p.find('img') and not p.find_next_sibling('img'):
                p.decompose()
        
        for div in soup.find_all('div'):
            div.replace_with_children()
        
        for tag in soup.find_all(True):
            if tag.name not in ['img', 'h2', 'span']:
                if tag.has_attr('style'):
                    del tag['style']
                for attr in ['class', 'id', 'name']:
                    if tag.has_attr(attr):
                        del tag[attr]
            elif tag.name == 'span':
                if tag.get('class') != ['nowrap']:
                    tag.attrs = {}
                    if 'nowrap' in str(tag):
                        tag['class'] = 'nowrap'
        
        for ins in soup.find_all('ins'):
            ins.replace_with(ins.text)
        
        for span in soup.find_all('span'):
            if not span.get('class') or 'nowrap' not in span.get('class', []):
                span.replace_with(span.text)
        
        # Traitement des tableaux
        for table in soup.find_all('table'):
            table['class'] = 'table table-bordered'
            
            caption = table.find('caption')
            if not caption:
                caption = soup.new_tag('caption')
                caption.string = '(Tableau)'
                table.insert(0, caption)
            
            thead = table.find('thead')
            if not thead and table.tr:
                thead = soup.new_tag('thead')
                thead['class'] = 'well'
                first_row = table.tr
                thead.append(first_row.extract())
                table.insert(1, thead)
                
                for td in thead.find_all('td'):
                    th = soup.new_tag('th')
                    th['scope'] = 'col'
                    th.string = td.get_text()
                    td.replace_with(th)
            
            tbody = table.find('tbody')
            if not tbody:
                tbody = soup.new_tag('tbody')
                for tr in table.find_all('tr'):
                    tbody.append(tr.extract())
                table.append(tbody)
            
            for tr in tbody.find_all('tr'):
                cells = tr.find_all('td')
                if cells:
                    cells[0]['scope'] = 'row'
            
            responsive_div = soup.new_tag('div')
            responsive_div['class'] = 'table-responsive'
            table.wrap(responsive_div)
        
        # CSS final avec votre style
        html_final = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>{nom_fichier} - Converti</title>
    <style>
        body {{ 
            font-family: Georgia, "Times New Roman", serif; 
            line-height: 1.6; 
            max-width: 800px; 
            margin: 0 auto; 
            padding: 20px; 
            color: #333;
        }}
        h2 {{ 
            color: #2c3e50; 
            font-size: 1.5em; 
            margin: 25px 0 15px 0; 
            padding-bottom: 8px; 
            border-bottom: 2px solid #3498db; 
            font-family: Arial, sans-serif;
        }}
        .sample-img {{ 
            max-width: 300px; 
            height: auto; 
            border: 2px solid #3498db; 
            border-radius: 5px;
            display: block; 
            margin: 15px 0;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }}
        .nowrap {{ 
            white-space: nowrap; 
            color: #2c3e50;
            font-weight: 500;
        }}
        strong {{ 
            font-weight: bold; 
            color: #2c3e50; 
        }}
        p {{ 
            margin: 12px 0; 
            text-align: justify;
        }}
        table {{ width: 100%; border-collapse: collapse; margin: 20px 0; }}
        th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
        th {{ background-color: #f2f2f2; font-weight: bold; }}
        body {{
            font-feature-settings: "liga", "kern";
            text-rendering: optimizeLegibility;
        }}
    </style>
</head>
<body>
{soup.body.decode_contents() if soup.body else str(soup)}
</body>
</html>"""
        
        # Extraire seulement le contenu body pour avoir le HTML final propre
        soup_final = BeautifulSoup(html_final, 'html.parser')
        clean_content = ""
        if soup_final.body:
            for element in soup_final.body.contents:
                if hasattr(element, 'name') and element.name:
                    clean_content += str(element)
                elif hasattr(element, 'strip') and element.strip():
                    clean_content += str(element)
        
        stats = {
            'nb_images': html_final.count('<img'),
            'titres_convertis': titres_convertis,
            'apostrophes_changees': apostrophes_changees,
            'mots_wrapes': mots_tirets_wrapes,
            'nb_paragraphes': len(structure_originale)
        }
        
        return clean_content, stats
        
    except Exception as e:
        st.error(f"Erreur lors de la conversion: {e}")
        return None, None

# Interface Streamlit
def main():
    # Header
    st.markdown('<h1 class="main-header">📄 Word to HTML Converter</h1>', unsafe_allow_html=True)
    
    # Sidebar
    with st.sidebar:
        st.header("⚙️ Options")
        
        afficher_stats = st.checkbox(
            "Afficher les statistiques", 
            value=True,
            help="Montre les détails de la conversion"
        )
        
        st.markdown("---")
        st.markdown("### 📋 Instructions")
        st.markdown("""
        1. **Uploadez** votre fichier Word (.docx)
        2. **Attendez** la conversion
        3. **Téléchargez** le HTML résultant
        4. Les images seront remplacées par `img_sample.jpg`
        5. Les titres seront automatiquement détectés
        6. Les apostrophes seront corrigées (') → (')
        7. Les mots à tirets seront protégés
        """)
    
    # Zone d'upload principale
    st.markdown('<div class="upload-section">', unsafe_allow_html=True)
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.markdown("### 📤 Upload your Word file")
        uploaded_file = st.file_uploader(
            "Choose a Word document",
            type=['docx'],
            help="Glissez-déposez votre fichier .docx ici ou cliquez pour parcourir"
        )
    
    with col2:
        st.markdown("### 📋 Supported formats")
        st.info("""
        **Formats acceptés:**
        - .docx (Word 2007+)
        - Taille max: 200MB
        - Images → img_sample.jpg
        """)
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    # Traitement du fichier
    if uploaded_file is not None:
        # Informations sur le fichier
        st.markdown('<div class="info-box">', unsafe_allow_html=True)
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("📄 Nom du fichier", uploaded_file.name)
        with col2:
            st.metric("📊 Taille", f"{uploaded_file.size / 1024:.1f} KB")
        with col3:
            st.metric("🗂️ Type", uploaded_file.type)
        st.markdown('</div>', unsafe_allow_html=True)
        
        # Bouton de conversion
        if st.button("🚀 Convertir en HTML", type="primary", use_container_width=True):
            with st.spinner("Conversion en cours..."):
                # Lire le fichier
                fichier_bytes = uploaded_file.getvalue()
                
                # Convertir avec votre code complet
                html_resultat, stats = convertir_word_vers_html_complet(fichier_bytes, uploaded_file.name)
                
                if html_resultat and stats:
                    # Section des résultats
                    st.markdown('<div class="result-section">', unsafe_allow_html=True)
                    st.success("✅ Conversion réussie!")
                    
                    # Statistiques
                    if afficher_stats:
                        st.markdown("### 📊 Statistiques de conversion")
                        col1, col2, col3, col4, col5 = st.columns(5)
                        with col1:
                            st.metric("🖼️ Images", stats['nb_images'])
                        with col2:
                            st.metric("📑 Titres H2", stats['titres_convertis'])
                        with col3:
                            st.metric("🔗 Mots protégés", stats['mots_wrapes'])
                        with col4:
                            st.metric("✏️ Apostrophes", stats['apostrophes_changees'])
                        with col5:
                            st.metric("📝 Paragraphes", stats['nb_paragraphes'])
                    
                    # Prévisualisation
                    st.markdown("### 👁️ Prévisualisation")
                    with st.expander("Voir le code HTML généré", expanded=False):
                        st.code(html_resultat, language='html')
                    
                    # Téléchargement
                    nom_sortie = uploaded_file.name.replace('.docx', '_converted.html')
                    
                    st.download_button(
                        label="⬇️ Télécharger le fichier HTML",
                        data=html_resultat,
                        file_name=nom_sortie,
                        mime="text/html",
                        type="primary",
                        use_container_width=True
                    )
                    
                    st.markdown('</div>', unsafe_allow_html=True)
                    
                    # Aperçu (note: les images img_sample.jpg ne s'afficheront que si le fichier existe)
                    st.markdown("### 🌐 Aperçu du rendu")
                    st.components.v1.html(
                        f"<div style='font-family: Georgia; line-height: 1.6; padding: 20px;'>{html_resultat}</div>", 
                        height=600, 
                        scrolling=True
                    )
                
                else:
                    st.error("❌ Échec de la conversion. Vérifiez que votre fichier est un document Word valide.")
    
    else:
        # Instructions quand aucun fichier n'est uploadé
        st.markdown("### 🎯 Fonctionnalités de ce convertisseur")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.markdown("""
            **🔧 Conversions automatiques:**
            - Word (.docx) → HTML propre
            - `<p><strong>` → `<h2>`
            - Images → `img_sample.jpg`
            - Nettoyage complet du code
            """)
        
        with col2:
            st.markdown("""
            **✨ Améliorations typographiques:**
            - `'` → `'` (apostrophes courbes)
            - `peut-il` → `<span class="nowrap">`
            - Suppression des div inutiles
            - CSS optimisé
            """)
        
        with col3:
            st.markdown("""
            **📊 Détection intelligente:**
            - Analyse de la structure XML
            - Position exacte des images
            - Titres vs texte en gras
            - Préservation de l'ordre
            """)

if __name__ == "__main__":
    main()