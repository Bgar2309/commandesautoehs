#!/usr/bin/env python3
"""
Application Flask pour traitement des commandes Prozon
Avec interface d'édition des références et poids
"""

from flask import Flask, render_template, request, jsonify, send_file, redirect, url_for
import pandas as pd
import PyPDF2
import os
import re
from pathlib import Path
from werkzeug.utils import secure_filename
from typing import Dict, List
import json

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'outputs'
app.config['EXCEL_FILE'] = 'uploads/Produits_référencés_EHS.xlsx'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max

# Créer les dossiers s'ils n'existent pas
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)


class ProzonOrderProcessor:
    def __init__(self, excel_path: str):
        """Initialise le processeur avec le fichier Excel de correspondances"""
        try:
            self.excel_path = excel_path
            self.df = pd.read_excel(excel_path)
            print(f"✅ Chargé {len(self.df)} références depuis {Path(excel_path).name}")
        except FileNotFoundError:
            # Créer un fichier Excel vide si inexistant
            self.df = pd.DataFrame(columns=[
                'Références Prozon',
                'Noms des produits',
                'Références EHS',
                'Prix',
                'poids'
            ])
            self.excel_path = excel_path
            self.save_excel()
            print(f"✅ Fichier Excel créé : {excel_path}")
    
    def save_excel(self):
        """Sauvegarde le DataFrame dans le fichier Excel"""
        self.df.to_excel(self.excel_path, index=False)
        print(f"💾 Fichier Excel sauvegardé : {self.excel_path}")
    
    def add_or_update_reference(self, ref_prozon: str, ref_ehs: str, 
                                nom_produit: str, poids: float, prix: float = None):
        """Ajoute ou met à jour une référence"""
        # Vérifier si la référence existe déjà
        existing = self.df[self.df['Références Prozon'] == ref_prozon]
        
        if not existing.empty:
            # Mettre à jour
            idx = existing.index[0]
            self.df.at[idx, 'Références EHS'] = ref_ehs
            self.df.at[idx, 'Noms des produits'] = nom_produit
            self.df.at[idx, 'poids'] = poids
            if prix:
                self.df.at[idx, 'Prix'] = prix
            action = "mise à jour"
        else:
            # Ajouter
            new_row = pd.DataFrame({
                'Références Prozon': [ref_prozon],
                'Noms des produits': [nom_produit],
                'Références EHS': [ref_ehs],
                'Prix': [prix if prix else ''],
                'poids': [poids]
            })
            self.df = pd.concat([self.df, new_row], ignore_index=True)
            action = "ajout"
        
        self.save_excel()
        return action
    
    def extract_text_from_pdf(self, pdf_path: str) -> str:
        """Extrait le texte brut d'un PDF"""
        text = ""
        with open(pdf_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            for page in reader.pages:
                text += page.extract_text()
        return text
    
    def parse_order(self, pdf_text: str) -> Dict:
        """Parse le texte du PDF pour extraire les informations structurées"""
        order = {
            'numero_commande': None,
            'ref_commande': None,
            'date': None,
            'adresse': {},
            'produits': []
        }
        
        # Extraire le numéro de commande
        num_match = re.search(r'#(LI\d+)', pdf_text)
        if num_match:
            order['numero_commande'] = num_match.group(1)
        
        # Extraire la référence de commande
        ref_match = re.search(r'Réf\. de commande.*?\n\s*(\w+)', pdf_text)
        if ref_match:
            order['ref_commande'] = ref_match.group(1)
        
        # Extraire la date
        date_match = re.search(r'LIVRAISON\s+(\d{2}/\d{2}/\d{4})', pdf_text)
        if date_match:
            order['date'] = date_match.group(1)
        
        # Extraire l'adresse de livraison
        adresse_section = re.search(
            r'Adresse de livraison\s+(.*?)(?=Réf\. de commande|$)', 
            pdf_text, 
            re.DOTALL
        )
        if adresse_section:
            adresse_text = adresse_section.group(1).strip()
            lines = [l.strip() for l in adresse_text.split('\n') if l.strip()]
            
            # Extraire tous les numéros de téléphone (format français)
            telephone_pattern = r'0\d{9}'
            telephones = []
            for line in lines:
                phone_matches = re.findall(telephone_pattern, line)
                telephones.extend(phone_matches)
            
            # Filtrer les lignes sans téléphone pour construire l'adresse
            adresse_lines = []
            for line in lines:
                # Garder la ligne si elle ne contient pas QUE des chiffres ou "France"
                if not re.match(r'^0\d{9}$', line) and line.lower() != 'france':
                    adresse_lines.append(line)
            
            # Construction intelligente de l'adresse
            nom_complet = ''
            rue = ''
            ville = ''
            pays = 'France'
            
            if len(adresse_lines) >= 1:
                # Première(s) ligne(s) = nom/société
                if len(adresse_lines) >= 3:
                    nom_complet = ' '.join(adresse_lines[:2])
                    rue = adresse_lines[2] if len(adresse_lines) > 2 else ''
                    ville = adresse_lines[3] if len(adresse_lines) > 3 else ''
                else:
                    nom_complet = adresse_lines[0]
                    rue = adresse_lines[1] if len(adresse_lines) > 1 else ''
                    ville = adresse_lines[2] if len(adresse_lines) > 2 else ''
            
            order['adresse'] = {
                'nom_complet': nom_complet,
                'rue': rue,
                'ville': ville,
                'pays': pays,
                'telephone': ', '.join(telephones) if telephones else '',
                'adresse_complete': '\n'.join(lines)
            }
        
        # Extraire les produits
        produits_section = re.search(
            r'Référence\s+Produit\s+Qté\s+(.*?)(?=Le destinataire|$)', 
            pdf_text, 
            re.DOTALL
        )
        
        if produits_section:
            produits_text = produits_section.group(1)
            ref_pattern = r'(\d{5}-\d+)\s+(.*?)(?=\d{5}-\d+|Le destinataire|$)'
            matches = re.finditer(ref_pattern, produits_text, re.DOTALL)
            
            for match in matches:
                ref_prozon = match.group(1)
                description_full = match.group(2).strip()
                
                qte_match = re.search(r'(\d+)\s*$', description_full)
                
                if qte_match:
                    qte = int(qte_match.group(1))
                    description = description_full[:qte_match.start()].strip()
                else:
                    qte = 1
                    description = description_full
                
                description = ' '.join(description.split())
                
                order['produits'].append({
                    'reference_prozon': ref_prozon,
                    'description': description,
                    'quantite': qte
                })
        
        return order
    
    def convert_reference(self, ref_prozon: str) -> List[Dict]:
        """Convertit une référence Prozon en référence(s) EHS avec poids"""
        matches = self.df[self.df['Références Prozon'] == ref_prozon]
        
        if matches.empty:
            return []
        
        results = []
        for _, row in matches.iterrows():
            results.append({
                'reference_ehs': row['Références EHS'],
                'nom_produit': row['Noms des produits'],
                'poids_unitaire': row['poids'] if pd.notna(row['poids']) else None,
                'prix': row['Prix'] if pd.notna(row['Prix']) else None
            })
        
        return results
    
    def process_pdf(self, pdf_path: str) -> Dict:
        """Traite un PDF complet et enrichit avec les correspondances"""
        text = self.extract_text_from_pdf(pdf_path)
        order = self.parse_order(text)
        
        # Enrichissement avec les correspondances EHS
        produits_expanded = []
        for produit in order['produits']:
            ref_prozon = produit['reference_prozon']
            correspondances = self.convert_reference(ref_prozon)
            
            if correspondances:
                for correspondance in correspondances:
                    produit_ehs = produit.copy()
                    produit_ehs['reference_ehs'] = correspondance['reference_ehs']
                    produit_ehs['nom_produit_ehs'] = correspondance['nom_produit']
                    produit_ehs['poids_unitaire'] = correspondance['poids_unitaire']
                    produit_ehs['prix_unitaire'] = correspondance['prix']
                    
                    if correspondance['poids_unitaire']:
                        produit_ehs['poids_total'] = correspondance['poids_unitaire'] * produit['quantite']
                    else:
                        produit_ehs['poids_total'] = None
                    
                    produit_ehs['statut'] = 'OK' if correspondance['poids_unitaire'] else 'POIDS_MANQUANT'
                    produits_expanded.append(produit_ehs)
            else:
                produit['reference_ehs'] = None
                produit['statut'] = 'NON_TROUVEE'
                produits_expanded.append(produit)
        
        order['produits'] = produits_expanded
        return order
    
    def export_to_csv(self, orders: List[Dict], output_path: str):
        """Exporte les commandes vers un CSV"""
        rows = []
        
        for order in orders:
            for prod in order['produits']:
                row = {
                    'Numero_Commande': order['numero_commande'],
                    'Ref_Commande': order['ref_commande'],
                    'Date': order['date'],
                    'Client': order['adresse']['nom_complet'],
                    'Adresse_Livraison': order['adresse']['rue'],
                    'Ville': order['adresse']['ville'],
                    'Telephone': order['adresse']['telephone'],
                    'Ref_Prozon': prod['reference_prozon'],
                    'Ref_EHS': prod.get('reference_ehs', 'NON_TROUVEE'),
                    'Quantite': prod['quantite'],
                    'Poids_Unitaire': prod.get('poids_unitaire', ''),
                    'Poids_Total': prod.get('poids_total', ''),
                    'Statut': prod['statut']
                }
                rows.append(row)
        
        df_export = pd.DataFrame(rows)
        df_export.to_csv(output_path, index=False, encoding='utf-8-sig')
        return df_export


# Instance globale du processeur
processor = None

def get_processor():
    global processor
    if processor is None:
        processor = ProzonOrderProcessor(app.config['EXCEL_FILE'])
    return processor


@app.route('/')
def index():
    """Page d'accueil"""
    return render_template('index.html')


@app.route('/references')
def references():
    """Page de gestion des références"""
    proc = get_processor()
    references_list = proc.df.to_dict('records')
    return render_template('references.html', references=references_list)


@app.route('/api/references', methods=['GET'])
def get_references():
    """API : Récupérer toutes les références"""
    proc = get_processor()
    return jsonify(proc.df.to_dict('records'))


@app.route('/api/references/add', methods=['POST'])
def add_reference():
    """API : Ajouter ou modifier une référence"""
    data = request.json
    proc = get_processor()
    
    try:
        action = proc.add_or_update_reference(
            ref_prozon=data['ref_prozon'],
            ref_ehs=data['ref_ehs'],
            nom_produit=data['nom_produit'],
            poids=float(data['poids']),
            prix=float(data.get('prix', 0)) if data.get('prix') else None
        )
        
        return jsonify({
            'success': True,
            'action': action,
            'message': f'Référence {data["ref_prozon"]} {action} avec succès'
        })
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        }), 400


@app.route('/api/upload', methods=['POST'])
def upload_files():
    """API : Upload de PDFs"""
    if 'files[]' not in request.files:
        return jsonify({'error': 'Aucun fichier'}), 400
    
    files = request.files.getlist('files[]')
    uploaded_files = []
    
    for file in files:
        if file and file.filename.endswith('.pdf'):
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)
            uploaded_files.append(filename)
    
    return jsonify({
        'success': True,
        'files': uploaded_files,
        'count': len(uploaded_files)
    })


@app.route('/api/process', methods=['POST'])
def process_orders():
    """API : Traiter les commandes uploadées"""
    proc = get_processor()
    
    # Trouver tous les PDFs dans uploads
    pdf_files = list(Path(app.config['UPLOAD_FOLDER']).glob('*.pdf'))
    
    if not pdf_files:
        return jsonify({'error': 'Aucun PDF à traiter'}), 400
    
    orders = []
    for pdf_file in pdf_files:
        try:
            order = proc.process_pdf(str(pdf_file))
            orders.append(order)
        except Exception as e:
            print(f"Erreur traitement {pdf_file}: {e}")
    
    return jsonify({
        'success': True,
        'orders': orders,
        'count': len(orders)
    })


@app.route('/api/export', methods=['POST'])
def export_csv():
    """API : Exporter les commandes en CSV"""
    data = request.json
    orders = data.get('orders', [])
    
    if not orders:
        return jsonify({'error': 'Aucune commande à exporter'}), 400
    
    proc = get_processor()
    output_path = os.path.join(app.config['OUTPUT_FOLDER'], 'export_commandes.csv')
    
    try:
        proc.export_to_csv(orders, output_path)
        return jsonify({
            'success': True,
            'file': 'export_commandes.csv',
            'path': output_path
        })
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        }), 400


@app.route('/download/<filename>')
def download_file(filename):
    """Télécharger un fichier exporté"""
    filepath = os.path.join(app.config['OUTPUT_FOLDER'], filename)
    return send_file(filepath, as_attachment=True)


if __name__ == '__main__':
    # Vérifier si le fichier Excel existe, sinon le créer
    if not os.path.exists(app.config['EXCEL_FILE']):
        print("Création du fichier Excel...")
        proc = ProzonOrderProcessor(app.config['EXCEL_FILE'])
    
    print("\n" + "="*80)
    print("🚀 APPLICATION PROZON - DÉMARRÉE")
    print("="*80)
    print(f"\n📂 Dossier uploads: {app.config['UPLOAD_FOLDER']}")
    print(f"📂 Dossier outputs: {app.config['OUTPUT_FOLDER']}")
    print(f"📊 Fichier Excel: {app.config['EXCEL_FILE']}")
    print(f"\n🌐 Ouvrir dans votre navigateur:")
    print(f"   → http://localhost:5000")
    print("\n" + "="*80 + "\n")
    
    app.run(debug=True, host='0.0.0.0', port=5000)