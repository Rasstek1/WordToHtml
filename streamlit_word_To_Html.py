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
    
    for element in soup.find_all(string=True):
        if element.parent.name not in ['script', 'style']:
            texte_original = str(element)
            texte_modifie = texte_original
            
            apostrophes_avant = texte_modifie.count("'")
            texte_modifie = texte_modifie.replace("'", "'")
            apostrophes_changees += apostrophes_avant
            
            pattern_tirets = r'\b([a-zA-ZÀ-ÿ]+(?:-[a-zA-ZÀ-ÿ]+)+)\b'
            
            def wrap_mot_tiret(match):
                mot = match.group(1)
                nonlocal mots_tirets_wrapes
                mots_tirets_wrapes += 1
                return f'<span class="nowrap">{mot}</span>'
            
            texte_modifie = re.sub(pattern_tirets, wrap_mot_tiret, texte_modifie)
            
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
    """Détecte et convertit les tables de matières en listes avec liens"""
    toc_convertie = False
    
    for h2 in soup.find_all('h2'):
        texte_titre = h2.get_text().lower()
        
        if ('table des matières' in texte_titre or 
            'table of contents' in texte_titre or
            'table des matieres' in texte_titre):
            
            st.info(f"📋 Table des matières détectée: {h2.get_text()}")
            
            next_element = h2.find_next_sibling()
            while next_element and next_element.name not in ['ol', 'ul']:
                next_element = next_element.find_next_sibling()
            
            if next_element and next_element.name in ['ol', 'ul']:
                convertir_liste_en_toc_simple(next_element, soup)
                toc_convertie = True
                break
    
    return toc_convertie

def convertir_liste_en_toc_simple(liste_element, soup):
    """Convertit une liste en table des matières avec liens"""
    
    def nettoyer_texte(texte):
        texte = texte.strip()
        texte = re.sub(r'^\d+\.?\s*', '', texte)
        texte = re.sub(r'^\d+\.\d+\.?\s*', '', texte)
        texte = re.sub(r'^\d+\.\d+\.\d+\.?\s*', '', texte)
        return texte.strip()
    
    if liste_element.name == 'ol':
        liste_element.name = 'ul'
    
    items_niveau_1 = liste_element.find_all('li', recursive=False)
    
    for i, li in enumerate(items_niveau_1):
        numero_1 = str(i + 1)
        
        # Extraire le texte du li principal
        texte_li = ""
        for content in li.contents:
            if hasattr(content, 'name'):
                if content.name not in ['ol', 'ul']:
                    texte_li += content.get_text() if hasattr(content, 'get_text') else str(content)
            else:
                texte_li += str(content)
        
        texte_clean = nettoyer_texte(texte_li)
        
        # Sauvegarder les sous-listes avant de vider
        sous_listes_niveau_2 = li.find_all(['ol', 'ul'], recursive=False)
        
        # Vider et reconstruire
        li.clear()
        
        # Créer le lien principal
        if texte_clean:
            lien = soup.new_tag('a', href=f"#{numero_1}")
            lien.string = f"{numero_1}.&nbsp;{texte_clean}"
            li.append(lien)
        
        # Traiter niveau 2
        if sous_listes_niveau_2:
            nouvelle_ul_2 = soup.new_tag('ul')
            items_niveau_2 = sous_listes_niveau_2[0].find_all('li', recursive=False)
            
            for j, li_2 in enumerate(items_niveau_2):
                numero_2 = f"{numero_1}.{j + 1}"
                
                # Extraire texte niveau 2
                texte_li_2 = ""
                for content in li_2.contents:
                    if hasattr(content, 'name'):
                        if content.name not in ['ol', 'ul']:
                            texte_li_2 += content.get_text() if hasattr(content, 'get_text') else str(content)
                    else:
                        texte_li_2 += str(content)
                
                texte_clean_2 = nettoyer_texte(texte_li_2)
                
                # Sauvegarder sous-listes niveau 3
                sous_listes_niveau_3 = li_2.find_all(['ol', 'ul'], recursive=False)
                
                # Créer nouveau li niveau 2
                nouveau_li_2 = soup.new_tag('li')
                
                if texte_clean_2:
                    lien_2 = soup.new_tag('a', href=f"#{numero_2}")
                    lien_2.string = f"{numero_2}&nbsp;{texte_clean_2}"
                    nouveau_li_2.append(lien_2)
                
                # Traiter niveau 3
                if sous_listes_niveau_3:
                    nouvelle_ul_3 = soup.new_tag('ul')
                    items_niveau_3 = sous_listes_niveau_3[0].find_all('li', recursive=False)
                    
                    for k, li_3 in enumerate(items_niveau_3):
                        numero_3 = f"{numero_2}.{k + 1}"
                        
                        texte_li_3 = ""
                        for content in li_3.contents:
                            if hasattr(content, 'name'):
                                if content.name not in ['ol', 'ul']:
                                    texte_li_3 += content.get_text() if hasattr(content, 'get_text') else str(content)
                            else:
                                texte_li_3 += str(content)
                        
                        texte_clean_3 = nettoyer_texte(texte_li_3)
                        
                        if texte_clean_3:
                            nouveau_li_3 = soup.new_tag('li')
                            lien_3 = soup.new_tag('a', href=f"#{numero_3}")
                            lien_3.string = f"{numero_3}&nbsp;{texte_clean_3}"
                            nouveau_li_3.append(lien_3)
                            nouvelle_ul_3.append(nouveau_li_3)
                    
                    if nouvelle_ul_3.contents:
                        nouveau_li_2.append(nouvelle_ul_3)
                
                nouvelle_ul_2.append(nouveau_li_2)
            
            if nouvelle_ul_2.contents:
                li.append(nouvelle_ul_2)

def ameliorer_traitement_tableaux(soup):
    """Améliore le traitement des tableaux"""
    tableaux_traites = 0
    
    # Gérer les listes qui précèdent les tableaux
    for ol in soup.find_all('ol'):
        items = ol.find_all('li', recursive=False)
        
        if len(items) == 1:
            texte_item = items[0].get_text().strip()
            next_sibling = ol.find_next_sibling()
            
            while next_sibling and not (next_sibling.name == 'div' and 'table-responsive' in str(next_sibling.get('class', []))):
                next_sibling = next_sibling.find_next_sibling()
            
            if (next_sibling and 
                'table-responsive' in str(next_sibling.get('class', [])) and
                len(texte_item) < 100):
                
                table = next_sibling.find('table')
                if table:
                    caption = table.find('caption')
                    if caption:
                        caption.string = texte_item
                    ol.decompose()
    
    # Améliorer la structure des tableaux
    for table_div in soup.find_all('div', class_='table-responsive'):
        table = table_div.find('table')
        if not table:
            continue
            
        tableaux_traites += 1
        
        thead = table.find('thead')
        tbody = table.find('tbody')
        
        if tbody and not thead:
            rows = tbody.find_all('tr')
            if rows:
                thead = soup.new_tag('thead')
                thead['class'] = 'well'
                
                header_row = soup.new_tag('tr')
                
                th1 = soup.new_tag('th', scope='col')
                th1.string = "Élément"
                
                th2 = soup.new_tag('th', scope='col')
                th2.string = "Description"
                
                header_row.append(th1)
                header_row.append(th2)
                thead.append(header_row)
                
                if table.caption:
                    table.insert(table.contents.index(table.caption) + 1, thead)
                else:
                    table.insert(0, thead)
        
        # Nettoyer le contenu des cellules
        for cell in table.find_all(['td', 'th']):
            for p in cell.find_all('p'):
                if p.get_text().strip():
                    p.replace_with(p.get_text())
                else:
                    p.decompose()
        
        # S'assurer que les cellules de la première colonne ont scope="row"
        if tbody:
            for row in tbody.find_all('tr'):
                cells = row.find_all(['td', 'th'])
                if cells:
                    cells[0]['scope'] = 'row'
    
    return tableaux_traites

def ajouter_ids_aux_titres(soup):
    """Ajoute des IDs aux titres H2"""
    titres_avec_id = 0
    
    for i, h2 in enumerate(soup.find_all('h2')):
        if not h2.get('id'):
            texte_titre = h2.get_text().strip()
            
            match = re.match(r'^(\d+)\.?\s*', texte_titre)
            if match:
                numero = match.group(1)
                h2['id'] = numero
            else:
                h2['id'] = str(i + 1)
            
            titres_avec_id += 1
    
    return titres_avec_id

def analyser_structure_document_bytes(fichier_word_bytes):
    """Analyse la structure complète du document Word"""
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
    """Remplace les images avec des données longues"""
    placeholder_path = "img_sample.jpg"
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
    """Convertit un fichier Word en HTML avec toutes les fonctionnalités"""
    try:
        import mammoth
        
        structure_originale = analyser_structure_document_bytes(fichier_word_bytes)
        
        with BytesIO(fichier_word_bytes) as docx_file:
            result = mammoth.convert_to_html(docx_file)
            html = result.value
        
        html_nettoye, nb_images_remplacees = nettoyer_images_dans_html(html)
        soup = BeautifulSoup(html_nettoye, 'html.parser')
        
        # Toutes les étapes de traitement
        titres_convertis = detecter_et_convertir_titres(soup)
        apostrophes_changees, mots_tirets_wrapes = ameliorations_typographiques(soup)
        toc_convertie = detecter_et_convertir_table_matieres(soup)
        tableaux_traites = ameliorer_traitement_tableaux(soup)
        titres_avec_id = ajouter_ids_aux_titres(soup)
        
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
            if tag.name not in ['img', 'h2', 'span', 'a', 'ul', 'li']:
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
        
        # CSS final
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
            'tableaux_traites': tableaux_traites,
            'titres_avec_id': titres_avec_id,
            'nb_paragraphes': len(structure_originale)
        }
        
        return clean_content, stats
        
    except Exception as e:
        st.error(f"Erreur lors de la conversion: {e}")
        return None, None

# Interface Streamlit
def main():
    st.markdown('<h1 class="main-header">📄 Word to HTML Converter</h1>', unsafe_allow_html=True)
    
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
        
        **✨ Fonctionnalités:**
        - Images → `img_sample.jpg`
        - Titres automatiquement détectés
        - Apostrophes corrigées (') → (')
        - Mots à tirets protégés
        - **📋 Tables de matières avec liens**
        - **📊 Tableaux améliorés**
        """)
    
    st.markdown('<div class="upload-section">', unsafe_allow_html=True)
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.markdown("### 📤 Upload your Word file")
        uploaded_file = st.file_uploader(
            "Choose a Word document",
            type=['docx'],
            help="Glissez-déposez votre fichier .docx ici"
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
    
    if uploaded_file is not None:
        st.markdown('<div class="info-box">', unsafe_allow_html=True)
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("📄 Nom du fichier", uploaded_file.name)
        with col2:
            st.metric("📊 Taille", f"{uploaded_file.size / 1024:.1f} KB")
        with col3:
            st.metric("🗂️ Type", uploaded_file.type)
        st.markdown('</div>', unsafe_allow_html=True)
        
        if st.button("🚀 Convertir en HTML", type="primary", use_container_width=True):
            with st.spinner("Conversion en cours..."):
                fichier_bytes = uploaded_file.getvalue()
                html_resultat, stats = convertir_word_vers_html_complet(fichier_bytes, uploaded_file.name)
                
                if html_resultat and stats:
                    st.markdown('<div class="result-section">', unsafe_allow_html=True)
                    st.success("✅ Conversion réussie!")
                    
                    if afficher_stats:
                        st.markdown("### 📊 Statistiques de conversion")
                        col1, col2, col3, col4 = st.columns(4)
                        with col1:
                            st.metric("🖼️ Images", stats['nb_images'])
                            st.metric("📑 Titres H2", stats['titres_convertis'])
                        with col2:
                            st.metric("🔗 Mots protégés", stats['mots_wrapes'])
                            st.metric("✏️ Apostrophes", stats['apostrophes_changees'])
                        with col3:
                            st.metric("📋 Table matières", "✅" if stats['toc_convertie'] else "❌")
                            st.metric("📊 Tableaux", stats['tableaux_traites'])
                        with col4:
                            st.metric("🔗 IDs titres", stats['titres_avec_id'])
                            st.metric("📝 Paragraphes", stats['nb_paragraphes'])
                    
                    st.markdown("### 👁️ Prévisualisation")
                    with st.expander("Voir le code HTML généré", expanded=False):
                        st.code(html_resultat, language='html')
                    
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
                    
                    st.markdown("### 🌐 Aperçu du rendu")
                    st.components.v1.html(
                        f"<div style='font-family: Georgia; line-height: 1.6; padding: 20px;'>{html_resultat}</div>", 
                        height=600, 
                        scrolling=True
                    )
                
                else:
                    st.error("❌ Échec de la conversion.")
    
    else:
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
            - Conversion avec liens
            - Numérotation hiérarchique
            - **📊 Tableaux améliorés**
            """)

if __name__ == "__main__":
    main()