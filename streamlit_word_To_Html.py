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

def detecter_et_convertir_table_matieres(soup):
    """
    Détecte et convertit les tables de matières en listes avec liens
    """
    toc_convertie = False
    
    # Chercher les titres "Table des matières" (français) et "Table of contents" (anglais)
    for h2 in soup.find_all('h2'):
        texte_titre = h2.get_text().lower()
        
        if ('table des matières' in texte_titre or 
            'table of contents' in texte_titre or
            'table des matieres' in texte_titre):  # Sans accent aussi
            
            st.info(f"📋 Table des matières détectée: {h2.get_text()}")
            
            # Chercher la liste qui suit ce titre
            next_element = h2.find_next_sibling()
            while next_element and next_element.name not in ['ol', 'ul']:
                next_element = next_element.find_next_sibling()
            
            if next_element and next_element.name in ['ol', 'ul']:
                # Convertir cette liste en table des matières avec liens
                convertir_liste_en_toc(next_element, soup)
                toc_convertie = True
                break
    
    return toc_convertie

def convertir_liste_en_toc(liste_element, soup):
    """
    Convertit une liste ordinaire en table des matières avec liens et numérotation
    """
    def extraire_texte_propre(li):
        """Extrait le texte d'un élément li en excluant les sous-listes"""
        texte = ""
        for content in li.contents:
            if hasattr(content, 'name'):
                if content.name not in ['ol', 'ul']:
                    if hasattr(content, 'get_text'):
                        texte += content.get_text()
                    else:
                        texte += str(content)
            else:
                texte += str(content)
        
        # Nettoyer le texte
        texte = texte.strip()
        texte = re.sub(r'^\d+\.?\s*', '', texte)  # Supprimer numérotation existante
        texte = re.sub(r'^\d+\.\d+\.?\s*', '', texte)
        texte = re.sub(r'^\d+\.\d+\.\d+\.?\s*', '', texte)
        texte = re.sub(r'\s+', ' ', texte)  # Normaliser les espaces
        
        return texte.strip()
    
    
    
    def traiter_niveau(items, niveau=1, numero_parent=""):
        """Traite récursivement chaque niveau de la liste"""
        for index, li in enumerate(items):
            numero_actuel = f"{numero_parent}{index + 1}" if numero_parent else str(index + 1)
            
            # Extraire le texte propre
            texte = extraire_texte_propre(li)
            
            if texte:
                # Trouver les sous-listes dans l'élément original
                sous_listes = li.find_all(['ol', 'ul'], recursive=False)
                
                # Vider l'élément li
                li.clear()
                
                # Créer le nouveau lien
                lien = soup.new_tag('a', href=f"#{numero_actuel}")
                lien.string = f"{numero_actuel}.&nbsp;{texte}"
                li.append(lien)
                
                # Traiter les sous-listes
                if sous_listes:
                    nouvelle_sous_liste = soup.new_tag('ul')
                    sous_items = sous_listes[0].find_all('li', recursive=False)
                    
                    for sous_index, sous_li in enumerate(sous_items):
                        sous_numero = f"{numero_actuel}.{sous_index + 1}"
                        sous_texte = extraire_texte_propre(sous_li)
                        
                        if sous_texte:
                            nouveau_sous_li = soup.new_tag('li')
                            sous_lien = soup.new_tag('a', href=f"#{sous_numero}")
                            sous_lien.string = f"{sous_numero}&nbsp;{sous_texte}"
                            nouveau_sous_li.append(sous_lien)
                            
                            # Gérer le troisième niveau
                            sous_sous_listes = sous_li.find_all(['ol', 'ul'], recursive=False)
                            if sous_sous_listes:
                                sous_sous_liste = soup.new_tag('ul')
                                sous_sous_items = sous_sous_listes[0].find_all('li', recursive=False)
                                
                                for sss_index, sss_li in enumerate(sous_sous_items):
                                    sss_numero = f"{sous_numero}.{sss_index + 1}"
                                    sss_texte = extraire_texte_propre(sss_li)
                                    
                                    if sss_texte:
                                        nouveau_sss_li = soup.new_tag('li')
                                        sss_lien = soup.new_tag('a', href=f"#{sss_numero}")
                                        sss_lien.string = f"{sss_numero}&nbsp;{sss_texte}"
                                        nouveau_sss_li.append(sss_lien)
                                        sous_sous_liste.append(nouveau_sss_li)
                                
                                if sous_sous_liste.contents:
                                    nouveau_sous_li.append(sous_sous_liste)
                            
                            nouvelle_sous_liste.append(nouveau_sous_li)
                    
                    if nouvelle_sous_liste.contents:
                        li.append(nouvelle_sous_liste)
    
    # Convertir ol en ul
    if liste_element.name == 'ol':
        liste_element.name = 'ul'
    
    # Traiter tous les éléments de premier niveau
    items_premier_niveau = liste_element.find_all('li', recursive=False)
    traiter_niveau(items_premier_niveau)
    

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
    placeholder_path = "img_sample.jpg"
    pattern_img_longue = r'<img[^>]*src="data:image/[^"]{100,}"[^>]*/?>'
    images_longues = re.findall(pattern_img_longue, html_content)
    
    counter = [1]
    
    def remplacer_image(match):
        img_tag = f'<img src="{placeholder_path}" alt="Image {counter[0]}" class="" />'
        counter[0] += 1
        return img_tag
    
    html_nettoye = re.sub(pattern_img_longue, remplacer_image, html_content)
    return html_nettoye, len(images_longues)

def convertir_word_vers_html_complet(fichier_word_bytes, nom_fichier):
    """Convertit un fichier Word en HTML avec toutes les fonctionnalités"""
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
        
        # NOUVELLE ÉTAPE 3 : Conversion des tables de matières
        toc_convertie = detecter_et_convertir_table_matieres(soup)
        
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
        
        # Nettoyage final
        for p in soup.find_all('p'):
            if not p.get_text(strip=True) and not p.find('img') and not p.find_next_sibling('img'):
                p.decompose()
        
        for div in soup.find_all('div'):
            div.replace_with_children()
        
        for tag in soup.find_all(True):
            if tag.name not in ['img', 'h2', 'span', 'a', 'ul', 'li']:  # Préserver les liens et listes
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
        
        # CSS final avec style pour les tables de matières
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
        /* Styles pour les tables de matières */
        ul {{ 
            margin: 10px 0; 
            padding-left: 20px;
        }}
        ul ul {{ 
            margin: 5px 0; 
            padding-left: 25px;
        }}
        ul ul ul {{ 
            margin: 3px 0; 
            padding-left: 25px;
        }}
        li {{ 
            margin: 3px 0; 
            line-height: 1.4;
        }}
        a {{ 
            color: #2c3e50; 
            text-decoration: none;
            border-bottom: 1px dotted #3498db;
        }}
        a:hover {{ 
            color: #3498db; 
            border-bottom: 1px solid #3498db;
        }}
        /* Tables */
        table {{ width: 100%; border-collapse: collapse; margin: 20px 0; }}
        th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
        th {{ background-color: #f2f2f2; font-weight: bold; }}
        /* Typographie */
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
            'toc_convertie': toc_convertie,
            'nb_paragraphes': len(structure_originale)
        }
        
        return clean_content, stats
        
    except Exception as e:
        st.error(f"Erreur lors de la conversion: {e}")
        return None, None

# Interface Streamlit
def main():
    # Header
    st.markdown("""
<div class="title-container" style="
    text-align: center;
    margin: 30px auto;
    padding: 20px;
    background: linear-gradient(135deg, #1a2a6c, #b21f1f, #fdbb2d);
    border-radius: 15px;
    box-shadow: 0 10px 30px rgba(0, 0, 0, 0.3);
    animation: glow 2s infinite alternate;
">
    <h1 style="
        font-family: 'Trebuchet MS', sans-serif;
        font-size: 3.2rem;
        font-weight: 800;
        letter-spacing: 3px;
        color: #fff;
        text-transform: uppercase;
        margin: 0;
        padding: 0;
        text-shadow: 3px 3px 10px rgba(0, 0, 0, 0.4);
    ">
        <span style="
            background: linear-gradient(90deg, #ff8a00, #e52e71);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            text-shadow: none;
        "></span>
        ☠️WORD TO HTML CONVERTER☠️
        <span style="
            background: linear-gradient(90deg, #e52e71, #ff8a00);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            text-shadow: none;
        "></span>
    </h1>
    <p style="
        font-family: 'Arial', sans-serif;
        font-size: 1.2rem;
        color: rgba(255, 255, 255, 0.8);
        margin-top: 10px;
        font-style: italic;
    ">Transform your documents with just one damn click!</p>
</div>

<style>
@keyframes glow {
    from {
        box-shadow: 0 0 20px -10px rgba(66, 133, 244, 0.8);
    }
    to {
        box-shadow: 0 0 25px 5px rgba(66, 133, 244, 0.8);
    }
}
</style>
""", unsafe_allow_html=True)
    st.markdown("""
<style>
.rasstek-signature {
    text-align: center;
    margin-top: 10px;
    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    font-size: 12px;
    font-weight: 500;
    color: #6c757d;
    letter-spacing: 1.5px;
    position: relative;
    overflow: hidden;
}
.rasstek-signature:after {
    content: '';
    position: absolute;
    bottom: 0;
    left: 0;
    width: 0;
    height: 2px;
    background: linear-gradient(90deg, #3498db, transparent);
    animation: line-anim 2s infinite;
}
@keyframes line-anim {
    0% { width: 0; left: 0; }
    50% { width: 100%; left: 0; }
    100% { width: 0; left: 100%; }
}
</style>
<div class="rasstek-signature">
    BUILT BY RASSTEK<sup>©</sup>
</div>
""", unsafe_allow_html=True)
    
    # Sidebar
    with st.sidebar:
        st.header("⚙️ Options")
        
        afficher_stats = st.checkbox(
            "Afficher les statistiques", 
            value=True,
            help="Montre les détails de la conversion"
        )
        
        # Ajout de la section de personnalisation des classes CSS
        st.markdown("---")
        st.header("🎨 Personnalisation CSS")
        
        # Créer un dictionnaire pour stocker les classes personnalisées
        custom_classes = {}
        
        # Ajouter un expander pour ne pas surcharger l'interface
        with st.expander("Personnaliser les balises HTML", expanded=False):
            st.markdown("Ajoutez des classes CSS à des balises HTML spécifiques (sans inclure l'attribut `class=`):")
            st.info("Exemple: Pour obtenir `<h1 class=\"h2\">`, entrez simplement `h2`")
            
            # Liste des balises que l'utilisateur peut personnaliser
            tags_to_customize = ['h1', 'h2', 'h3', 'p', 'ul', 'ol', 'li', 'table', 'img']
            
            # Créer une interface pour chaque tag
            for tag in tags_to_customize:
                col1, col2 = st.columns([1, 2])
                with col1:
                    st.markdown(f"**&lt;{tag}&gt;**")
                with col2:
                    class_value = st.text_input("", key=f"class_{tag}", placeholder=f"Nom(s) de classe pour {tag}")
                    if class_value:
                        custom_classes[tag] = class_value
        
        st.markdown("---")
        st.markdown("### 📋 Instructions")
        st.markdown("""
        1. **Uploadez** votre fichier Word (.docx)
        2. **Attendez** la conversion
        3. **Téléchargez** le HTML résultant
        
        **✨ Nouvelles fonctionnalités:**
        - Images → `img_sample.jpg`
        - Titres automatiquement détectés
        - Apostrophes corrigées (') → (')
        - Mots à tirets protégés
        - **📋 Tables de matières avec liens !**
        - **🎨 Personnalisation des classes CSS**
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
        - Tables de matières auto-détectées
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
                    # Appliquer les classes personnalisées si nécessaire
                    if custom_classes:
                        html_resultat = appliquer_classes_personnalisees(html_resultat, custom_classes)
                    
                    # Section des résultats
                    st.markdown('<div class="result-section">', unsafe_allow_html=True)
                    st.success("✅ Conversion réussie!")
                    
                    # Statistiques
                    if afficher_stats:
                        st.markdown("### 📊 Statistiques de conversion")
                        col1, col2, col3, col4, col5, col6 = st.columns(6)
                        with col1:
                            st.metric("🖼️ Images", stats['nb_images'])
                        with col2:
                            st.metric("📑 Titres H2", stats['titres_convertis'])
                        with col3:
                            st.metric("🔗 Mots protégés", stats['mots_wrapes'])
                        with col4:
                            st.metric("✏️ Apostrophes", stats['apostrophes_changees'])
                        with col5:
                            st.metric("📋 Table matières", "✅" if stats['toc_convertie'] else "❌")
                        with col6:
                            st.metric("📝 Paragraphes", stats['nb_paragraphes'])
                        
                        # Ajouter des statistiques sur les classes personnalisées
                        if custom_classes:
                            st.markdown("### 🎨 Classes CSS personnalisées appliquées")
                            classes_cols = st.columns(len(custom_classes))
                            for i, (tag, classe) in enumerate(custom_classes.items()):
                                with classes_cols[i]:
                                    st.metric(f"Tag <{tag}>", classe)
                    
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
                    
                    # Aperçu
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
        st.markdown("### 🎯 Fonctionnalités du convertisseur")
        
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
            **📋 Tables de matières:**
            - Détection automatique
            - Conversion `<ol>` → `<ul>`
            - Liens avec ancres `#1.2.3`
            - Numérotation hiérarchique
            """)
            
            st.markdown("""
            **🎨 Personnalisation CSS:**
            - Ajout de classes aux balises HTML
            - Exemple: `h2` → `h2 class="my-custom-class"`
            - Styles personnalisés pour tous les éléments
            """)


# Fonction pour appliquer les classes personnalisées au HTML
def appliquer_classes_personnalisees(html_content, custom_classes):
    """
    Applique les classes CSS personnalisées aux balises HTML spécifiées.
    
    Args:
        html_content (str): Le contenu HTML à modifier
        custom_classes (dict): Dictionnaire des balises et leurs classes à appliquer
    
    Returns:
        str: Le HTML modifié avec les classes appliquées
    """
    if not custom_classes:
        return html_content
    
    try:
        from bs4 import BeautifulSoup
        
        # Parser le HTML
        soup = BeautifulSoup(html_content, 'html.parser')
        
        # Appliquer les classes à chaque type de balise
        for tag, class_value in custom_classes.items():
            # Trouver toutes les balises du type spécifié
            elements = soup.find_all(tag)
            
            for element in elements:
                # Nettoyer la valeur de classe entrée par l'utilisateur
                # Supprimer les attributs 'class=' ou class=" s'ils sont inclus
                cleaned_value = class_value.replace('class=', '').strip()
                if cleaned_value.startswith('"') and cleaned_value.endswith('"'):
                    cleaned_value = cleaned_value[1:-1]
                elif cleaned_value.startswith("'") and cleaned_value.endswith("'"):
                    cleaned_value = cleaned_value[1:-1]
                
                # Diviser en classes individuelles
                new_classes = cleaned_value.split()
                
                # Récupérer les classes existantes
                existing_classes = element.get('class', [])
                
                # Convertir en liste si c'est une chaîne ou None
                if existing_classes is None:
                    existing_classes = []
                elif isinstance(existing_classes, str):
                    existing_classes = [existing_classes]
                
                # Fusionner les classes sans duplicats
                final_classes = list(existing_classes)
                for cls in new_classes:
                    if cls not in final_classes:
                        final_classes.append(cls)
                
                # Appliquer les classes fusionnées
                element['class'] = final_classes
        
        # Convertir le soup modifié en string
        return str(soup)
    
    except Exception as e:
        # En cas d'erreur, retourner le HTML original
        print(f"Erreur lors de l'application des classes personnalisées: {e}")
        return html_content

if __name__ == "__main__":
    main()