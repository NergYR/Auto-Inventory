from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.label import Label
from plyer import camera # Décommenté
from pyzbar.pyzbar import decode # Décommenté
from PIL import Image # Décommenté
import os # Décommenté
# Ajout pour la gestion des états et des données
from kivy.uix.textinput import TextInput
from plyer import storagepath # Ajout pour le chemin de sauvegarde
from datetime import datetime # Pour nommer le fichier Excel
import openpyxl # Pour la création du fichier Excel
from kivy.utils import get_color_from_hex # Pour les couleurs
from kivy.uix.popup import Popup # Pour les confirmations (si besoin plus tard)
from kivy.properties import ListProperty # Pour la couleur du label

# Définition des couleurs
COLOR_SUCCESS = get_color_from_hex("#32CD32")  # Vert lime
COLOR_ERROR = get_color_from_hex("#FF0000")    # Rouge
COLOR_INFO = get_color_from_hex("#1E90FF")     # Bleu Dodger
COLOR_DEFAULT = get_color_from_hex("#FFFFFF")  # Blanc (ou la couleur par défaut du thème)

class StatusLabel(Label):
    background_color = ListProperty(get_color_from_hex('#00000000')) # Fond transparent par défaut
    # Pour KivyMD, on utiliserait des propriétés différentes, mais pour Kivy standard, on peut jouer avec le canvas
    # ou mettre un fond au BoxLayout parent.
    # Pour la simplicité, nous allons nous concentrer sur la couleur du texte.
    pass

class InventoryApp(App):
    def build(self):
        self.title = "Inventaire Automatique V2"
        self.scanning_state = "SERIAL"  # États possibles: "SERIAL", "MAC"
        self.current_serial = None
        self.inventory_data = [] # Pour stocker les paires (N/S, MAC)

        # Layout principal
        main_layout = BoxLayout(orientation='vertical', padding=10, spacing=10)

        # Label de statut amélioré
        self.status_label = StatusLabel(
            text="Prêt à scanner le NUMÉRO DE SÉRIE...", 
            size_hint_y=None, 
            height=50, # Hauteur augmentée
            font_size='18sp', # Taille de police augmentée
            color=COLOR_INFO # Couleur initiale
        )
        # Pour centrer le texte, si ce n'est pas par défaut avec le widget:
        self.status_label.halign = 'center'
        self.status_label.valign = 'middle'
        # S'assurer que le label peut s'étendre pour centrer le texte
        self.status_label.bind(size=self.status_label.setter('text_size')) 

        # Bouton de scan
        scan_button = Button(text="Scanner", on_press=self.scan_barcode, size_hint_y=None, height=50)
        
        # Layout pour les informations et actions sur les données
        data_actions_layout = BoxLayout(orientation='horizontal', spacing=10, size_hint_y=None, height=50)
        self.items_count_label = Label(text="Éléments: 0", size_hint_x=0.4)
        delete_last_button = Button(text="Effacer Dernier", on_press=self.delete_last_entry, size_hint_x=0.6)
        data_actions_layout.add_widget(self.items_count_label)
        data_actions_layout.add_widget(delete_last_button)

        # Layout pour les boutons de sauvegarde/export
        export_layout = BoxLayout(orientation='horizontal', spacing=10, size_hint_y=None, height=50)
        excel_button = Button(text="Sauvegarder en Excel", on_press=self.save_to_excel, disabled=False)
        glpi_button = Button(text="Envoyer vers GLPI (Bientôt)", disabled=True) # Gardé pour plus tard
        export_layout.add_widget(excel_button)
        export_layout.add_widget(glpi_button)

        main_layout.add_widget(self.status_label)
        main_layout.add_widget(scan_button)
        main_layout.add_widget(data_actions_layout)
        main_layout.add_widget(export_layout)
        
        self._update_items_count_label() # Mise à jour initiale du compteur
        return main_layout

    def _update_status_label(self, message, msg_type="INFO"):
        self.status_label.text = message
        if msg_type == "SUCCESS":
            self.status_label.color = COLOR_SUCCESS
        elif msg_type == "ERROR":
            self.status_label.color = COLOR_ERROR
        elif msg_type == "INFO":
            self.status_label.color = COLOR_INFO
        else:
            self.status_label.color = COLOR_DEFAULT

    def _update_items_count_label(self):
        self.items_count_label.text = f"Éléments: {len(self.inventory_data)}"

    def scan_barcode(self, instance):
        try:
            self.temp_image_path = os.path.join(os.getcwd(), "temp_scan_inv.png")
            camera.take_picture(filename=self.temp_image_path,
                                on_complete=self.process_image)
            self._update_status_label("Ouverture de la caméra...", "INFO")
        except NotImplementedError:
            self._update_status_label("L\'accès à la caméra n\'est pas supporté.", "ERROR")
        except Exception as e:
            self._update_status_label(f"Erreur caméra: {e}", "ERROR")

    def process_image(self, file_path):
        if not file_path or not os.path.exists(file_path):
            self._update_status_label("Chemin de l\'image invalide ou fichier non trouvé.", "ERROR")
            return

        self._update_status_label(f"Traitement de: {os.path.basename(file_path)}", "INFO")
        try:
            img = Image.open(file_path)
            barcodes = decode(img)
            
            if barcodes:
                barcode_data = barcodes[0].data.decode('utf-8')
                # barcode_type = barcodes[0].type # Moins utile pour l'utilisateur final ici
                
                if self.scanning_state == "SERIAL":
                    self.current_serial = barcode_data
                    self.scanning_state = "MAC"
                    self._update_status_label(f"N/S: {self.current_serial}. Scannez l\'ADRESSE MAC.", "INFO")
                elif self.scanning_state == "MAC":
                    mac_address = barcode_data
                    self.inventory_data.append({"serial": self.current_serial, "mac": mac_address})
                    self._update_status_label(f"N/S: {self.current_serial}, MAC: {mac_address} - Enregistré!", "SUCCESS")
                    print(f"Données enregistrées: {self.inventory_data[-1]}") 
                    self.current_serial = None
                    self.scanning_state = "SERIAL"
                    self._update_items_count_label()
                    # Prêt pour le prochain N/S, le message sera mis à jour au prochain scan ou action
            else:
                current_action = "NUMÉRO DE SÉRIE" if self.scanning_state == "SERIAL" else "ADRESSE MAC"
                ns_info = f" pour N/S: {self.current_serial}" if self.scanning_state == "MAC" and self.current_serial else ""
                self._update_status_label(f"Aucun {current_action} détecté{ns_info}. Réessayez.", "ERROR")
        except Exception as e:
            self._update_status_label(f"Erreur de décodage: {e}", "ERROR")
        finally:
            if os.path.exists(file_path):
                try:
                    os.remove(file_path)
                except Exception as e:
                    print(f"Erreur suppression fichier temp: {e}")

    def delete_last_entry(self, instance):
        if self.inventory_data:
            last_item = self.inventory_data.pop()
            self._update_status_label(f"Dernière entrée effacée: N/S {last_item['serial']}, MAC {last_item['mac']}", "INFO")
            self._update_items_count_label()
            # Si on a supprimé une MAC, on devrait peut-être revenir à l'état de scan MAC pour le N/S précédent?
            # Pour l'instant, on réinitialise à SERIAL pour la simplicité.
            self.scanning_state = "SERIAL"
            self.current_serial = None 
            # Le prochain scan demandera un N/S. Si l'utilisateur voulait juste corriger la MAC,
            # il devra rescanner le N/S puis la nouvelle MAC.
            # Une logique plus complexe pourrait être ajoutée ici si nécessaire.
        else:
            self._update_status_label("Aucune entrée à effacer.", "ERROR")

    def save_to_excel(self, instance):
        if not self.inventory_data:
            self._update_status_label("Aucune donnée à sauvegarder.", "ERROR")
            return
        # ... (le reste de la fonction save_to_excel reste majoritairement la même,
        # mais utilisera _update_status_label)
        try:
            # ... (début de la logique de sauvegarde existante)
            try:
                documents_dir = storagepath.get_documents_dir()
                if documents_dir is None: 
                    raise NotImplementedError("Documents directory not available via plyer")
            except (NotImplementedError, Exception) as e:
                print(f"Plyer storagepath non disponible ({e}), utilisation du répertoire courant.")
                documents_dir = os.getcwd()
            
            if not documents_dir:
                self._update_status_label("Impossible de déterminer le dossier de sauvegarde.", "ERROR")
                return

            inventory_save_dir = os.path.join(documents_dir, "InventairesAuto")
            if not os.path.exists(inventory_save_dir):
                os.makedirs(inventory_save_dir)

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"inventaire_{timestamp}.xlsx"
            filepath = os.path.join(inventory_save_dir, filename)

            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.title = "Inventaire"
            sheet["A1"] = "Numéro de Série"
            sheet["B1"] = "Adresse MAC"

            for index, item in enumerate(self.inventory_data, start=2):
                sheet[f"A{index}"] = item["serial"]
                sheet[f"B{index}"] = item["mac"]
            
            workbook.save(filepath)
            self._update_status_label(f"Inventaire sauvegardé: {filename} ({len(self.inventory_data)} équip.)", "SUCCESS")
            self.inventory_data = [] 
            self._update_items_count_label()
            self.scanning_state = "SERIAL"
            self.current_serial = None
            # Après sauvegarde, on pourrait explicitement mettre "Prêt pour nouvel inventaire"
            # self._update_status_label("Prêt pour un nouvel inventaire. Scannez un N/S.", "INFO") 
            # Cependant, le prochain scan de N/S mettra à jour le message de manière appropriée.

        except Exception as e:
            self._update_status_label(f"Erreur sauvegarde Excel: {e}", "ERROR")
            print(f"Erreur détaillée sauvegarde Excel: {e}")

if __name__ == '__main__':
    InventoryApp().run()
