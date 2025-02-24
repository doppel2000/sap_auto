import os
import configparser
import tkinter as tk
from tkinter import ttk, messagebox
import pythoncom
import win32com.client
import subprocess
import time
import sys

CONFIG_FILE = "config.ini"

def load_config():
    """Charge la config depuis config.ini (ou crée par défaut si absent)."""
    config = configparser.ConfigParser()
    if not os.path.exists(CONFIG_FILE):
        # Créer un config.ini par défaut
        config["SAP"] = {
            "saplogon_path": r"C:\Program Files\SAP\FrontEnd\SAPGUI\saplogon.exe",
            "sap_environment": "1075 PW1 - EWM Production with SSO"
        }
        config["Transactions"] = {
            "favorites": "/n/scwm/mon, /n/scwm/packspec, /n/scwm/mat1"
        }
        config["App"] = {
            "version": "1.0.0"
        }
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            config.write(f)

    config.read(CONFIG_FILE, encoding="utf-8")

    saplogon_path = config["SAP"].get("saplogon_path", r"C:\Program Files\SAP\FrontEnd\SAPGUI\saplogon.exe")
    sap_env = config["SAP"].get("sap_environment", "1075 PW1 - EWM Production with SSO")

    tx_list = config["Transactions"].get("favorites", "/n/scwm/mon,/n/scwm/packspec").split(",")
    tx_list = [tx.strip() for tx in tx_list if tx.strip()]

    version_app = config["App"].get("version", "1.0.0")

    return saplogon_path, sap_env, tx_list, version_app

def save_config(saplogon_path, sap_env, tx_list, version_app="1.0.0"):
    """Enregistre la config dans config.ini."""
    config = configparser.ConfigParser()
    config["SAP"] = {
        "saplogon_path": saplogon_path,
        "sap_environment": sap_env
    }
    config["Transactions"] = {
        "favorites": ", ".join(tx_list)
    }
    config["App"] = {
        "version": version_app
    }
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        config.write(f)

def get_sap_gui_object(timeout=30):
    """
    Tente de récupérer l'objet SAPGUI (ScriptingEngine) pendant 'timeout' secondes.
    Retourne l'objet si disponible, sinon None après expiration.
    """
    SapGuiAuto = None
    start_time = time.time()
    while True:
        try:
            SapGuiAuto = win32com.client.GetObject("SAPGUI")
            break
        except pythoncom.com_error:
            time.sleep(1)
            if (time.time() - start_time) > timeout:
                return None
    return SapGuiAuto

def find_or_open_connection(saplogon_path, sap_env):
    """
    Lance SAP Logon si nécessaire, récupère la connexion SSO vers 'sap_env'.
    Renvoie (connection, application) ou lève une exception en cas d'erreur.
    """
    try:
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
    except pythoncom.com_error:
        SapGuiAuto = None

    if not SapGuiAuto:
        try:
            subprocess.Popen([saplogon_path])
        except FileNotFoundError:
            raise FileNotFoundError(f"Impossible de trouver saplogon.exe : {saplogon_path}")

        SapGuiAuto = get_sap_gui_object(timeout=30)
        if not SapGuiAuto:
            raise TimeoutError("Timeout : SAP GUI ne s'est pas lancé dans les 30 secondes.")

    application = SapGuiAuto.GetScriptingEngine
    if not application:
        raise RuntimeError("Impossible d'obtenir la ScriptingEngine de SAPGUI.")

    connection = None
    for i in range(application.Children.Count):
        conn = application.Children(i)
        conn_name = getattr(conn, "Name", "")
        conn_desc = getattr(conn, "Description", "")
        if sap_env in conn_name or sap_env in conn_desc:
            connection = conn
            break

    if not connection:
        try:
            connection = application.OpenConnection(sap_env, True)
        except Exception as e:
            raise RuntimeError(f"Impossible d'ouvrir la connexion '{sap_env}' : {e}")

    return connection, application

def launch_sap_transaction(saplogon_path, sap_env, transaction_code, session_choice, sessions_map):
    """
    Lance la transaction dans la session choisie.
    - session_choice = "Nouvelle session" => on utilise /o <transaction> dans la PREMIÈRE session existante.
    - Sinon, on lance dans la session spécifiée par sessions_map.
    """
    connection, _ = find_or_open_connection(saplogon_path, sap_env)
    nb_sessions = connection.Children.Count

    if session_choice == "Nouvelle session":
        if nb_sessions == 0:
            raise RuntimeError("Aucune session existante pour effectuer l'ouverture /o.")
        session = connection.Children(0)
        if transaction_code.startswith("/n"):
            transaction_code = "/o" + transaction_code[2:]
        elif not transaction_code.startswith("/o"):
            transaction_code = "/o" + transaction_code.lstrip("/")
        session.findById("wnd[0]/tbar[0]/okcd").text = transaction_code
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(2)
        nb_after = connection.Children.Count
        if nb_after <= nb_sessions:
            raise RuntimeError("La nouvelle session n'a pas été créée malgré /o.")
        new_sess = connection.Children(nb_after - 1)
        new_sess.findById("wnd[0]").maximize()
    else:
        if session_choice not in sessions_map:
            raise ValueError(f"La session '{session_choice}' n'existe pas.")
        idx = sessions_map[session_choice]
        if idx < 0 or idx >= nb_sessions:
            raise IndexError(f"La session d'index {idx} n'existe pas (il y a {nb_sessions} session(s)).")
        session = connection.Children(idx)
        if not transaction_code.startswith("/"):
            transaction_code = "/n" + transaction_code
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = transaction_code
        session.findById("wnd[0]").sendVKey(0)

#
# ---- INTERFACE GRAPHIQUE (tkinter) ----
#
class MainApp(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("SAP Auto")
        self.geometry("800x250")

        # Charger config
        self.saplogon_path, self.sap_env, self.tx_list, self.version_app = load_config()

        # sessions_map : { "Session 0 - TitreFenetre" : 0, ... }
        self.sessions_map = {}

        self.create_menu_bar()
        self.create_main_widgets()

    def create_menu_bar(self):
        menubar = tk.Menu(self)

        # Menu Fichier
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Créer des HU", command=self.open_hu_creation_window)
        file_menu.add_command(label="Quitter", command=self.quit_application)
        menubar.add_cascade(label="Fichier", menu=file_menu)

        # Menu Configuration
        config_menu = tk.Menu(menubar, tearoff=0)
        config_menu.add_command(label="Modifier", command=self.show_config_window)
        menubar.add_cascade(label="Configuration", menu=config_menu)

        # Menu Help
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="Aide", command=self.show_help)
        help_menu.add_command(label="À propos", command=self.show_about)
        menubar.add_cascade(label="Help", menu=help_menu)

        self.config(menu=menubar)

    def open_hu_creation_window(self):
        HUCreationWindow(self)

    def quit_application(self):
        self.destroy()

    def show_config_window(self):
        config_win = tk.Toplevel(self)
        config_win.title("Configuration SAP")
        config_win.geometry("500x200")
        config_win.transient(self)  # La fenêtre devient transitoire par rapport à la fenêtre parente
        config_win.grab_set()       # Capture tous les événements clavier/souris pour cette fenêtre

        tk.Label(config_win, text="Chemin saplogon.exe :").pack(pady=5)
        entry_path = tk.Entry(config_win, width=60)
        entry_path.pack()
        entry_path.insert(0, self.saplogon_path)

        tk.Label(config_win, text="Environnement SAP :").pack(pady=5)
        entry_env = tk.Entry(config_win, width=60)
        entry_env.pack()
        entry_env.insert(0, self.sap_env)

        def save_conf():
            self.saplogon_path = entry_path.get()
            self.sap_env = entry_env.get()
            save_config(self.saplogon_path, self.sap_env, self.tx_list, self.version_app)
            messagebox.showinfo("Info", "Configuration mise à jour.")
            config_win.destroy()

        tk.Button(config_win, text="Enregistrer", command=save_conf).pack(pady=10)

        self.wait_window(config_win)  # Attend que config_win soit fermé

    def show_help(self):
        msg = (
            "Aide de l'application :\n\n"
            "1) Menu Configuration pour modifier le chemin saplogon.exe et l'environnement SAP.\n"
            "2) Sélectionnez une transaction à lancer dans la liste déroulante ou choisissez 'Créer des HU' dans le menu Fichier.\n"
            "3) Pour lancer une transaction, choisissez la session ou optez pour 'Nouvelle session'.\n"
            "4) Dans la fenêtre 'Créer des HU', renseignez les champs et lancez la création.\n"
        )
        messagebox.showinfo("Aide", msg)

    def show_about(self):
        messagebox.showinfo("À propos", f"Version de l'application : {self.version_app}")

    def create_main_widgets(self):
        tk.Label(self, text="Lancement d'une transaction SAP", font=("Arial", 14, "bold")).pack(pady=10)

        frame_tx = tk.Frame(self)
        frame_tx.pack(pady=5)

        tk.Label(frame_tx, text="Transaction :").pack(side=tk.LEFT, padx=5)
        self.combo_tx = ttk.Combobox(frame_tx, values=self.tx_list, width=90)
        if self.tx_list:
            self.combo_tx.set(self.tx_list[0])
        self.combo_tx.pack(side=tk.LEFT)

        update_tx_btn = tk.Button(frame_tx, text="Enregistrer transaction", command=self.add_transaction_to_list)
        update_tx_btn.pack(side=tk.LEFT, padx=5)

        frame_sessions = tk.Frame(self)
        frame_sessions.pack(pady=5)

        tk.Label(frame_sessions, text="Session :").pack(side=tk.LEFT, padx=5)
        self.combo_session = ttk.Combobox(frame_sessions, width=90)
        self.combo_session.pack(side=tk.LEFT)

        refresh_btn = tk.Button(frame_sessions, text="Rafraîchir sessions", command=self.refresh_sessions)
        refresh_btn.pack(side=tk.LEFT, padx=5)

        # Au lancement, on charge la liste des sessions
        self.refresh_sessions()

        launch_btn = tk.Button(self, text="Lancer la transaction", bg="green", fg="white", command=self.on_launch_click)
        launch_btn.pack(pady=10)

    def refresh_sessions(self):
        """Met à jour la liste des sessions en affichant "Session i - <titre>" puis ajoute "Nouvelle session"."""
        try:
            connection, _ = find_or_open_connection(self.saplogon_path, self.sap_env)
            nb_sessions = connection.Children.Count

            sessions_display = []
            self.sessions_map = {}

            for i in range(nb_sessions):
                sess = connection.Children(i)
                try:
                    wnd_title = sess.findById("wnd[0]").Text
                except:
                    wnd_title = "(Sans titre)"
                display_name = f"Session {i} - {wnd_title}"
                sessions_display.append(display_name)
                self.sessions_map[display_name] = i

            sessions_display.append("Nouvelle session")
            self.combo_session['values'] = sessions_display

            if sessions_display:
                self.combo_session.set(sessions_display[0])
            else:
                self.combo_session.set("Nouvelle session")

        except Exception as e:
            self.combo_session['values'] = ["Nouvelle session"]
            self.combo_session.set("Nouvelle session")
            messagebox.showwarning("Attention", f"Impossible de lire la liste des sessions : {e}")

    def add_transaction_to_list(self):
        """Ajoute la transaction saisie dans la combobox à la liste, puis sauvegarde."""
        new_tx = self.combo_tx.get().strip()
        if new_tx and new_tx not in self.tx_list:
            self.tx_list.append(new_tx)
            self.combo_tx['values'] = self.tx_list
            save_config(self.saplogon_path, self.sap_env, self.tx_list, self.version_app)
            messagebox.showinfo("Info", f"Transaction '{new_tx}' ajoutée à la liste.")
        else:
            messagebox.showinfo("Info", f"La transaction '{new_tx}' est déjà dans la liste ou est vide.")

    def on_launch_click(self):
        transaction_code = self.combo_tx.get().strip()
        session_choice = self.combo_session.get().strip()

        if not transaction_code:
            messagebox.showerror("Erreur", "Veuillez saisir ou sélectionner une transaction.")
            return

        try:
            launch_sap_transaction(
                self.saplogon_path,
                self.sap_env,
                transaction_code,
                session_choice,
                self.sessions_map
            )
            messagebox.showinfo("Succès", f"La transaction {transaction_code} a été lancée.")
        except Exception as e:
            messagebox.showerror("Erreur", f"{type(e).__name__}: {e}")

        self.refresh_sessions()

#
# ---- Fenêtre de création des HU ----
#
class HUCreationWindow(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Créer des HU")
        self.geometry("600x500")
        self.parent = parent
        
        # La fenêtre devient transitoire par rapport à la fenêtre parente
        self.transient(self.parent)
        # Capture tous les événements clavier/souris pour cette fenêtre
        self.grab_set()                         

        # Charger ou initialiser la config HU dans le fichier .ini
        self.config_parser = configparser.ConfigParser()
        self.config_parser.read(CONFIG_FILE, encoding="utf-8")
        if not self.config_parser.has_section("HU"):
            self.config_parser.add_section("HU")
            self.config_parser.set("HU", "work_center", "GPAK")
            self.config_parser.set("HU", "storage_bin", "COOL-PACK, GR-ZONE")
            self.config_parser.set("HU", "hu_type", "PAC0002, PAC0005, PAC0008, PAC0011, PAC0012")            
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                self.config_parser.write(f)

        # Charcher et trier alphabétiquement les valeurs
        wc_values = sorted([v.strip() for v in self.config_parser.get("HU", "work_center").split(",") if v.strip()])
        sb_values = sorted([v.strip() for v in self.config_parser.get("HU", "storage_bin").split(",") if v.strip()])
        ht_values = sorted([v.strip() for v in self.config_parser.get("HU", "hu_type").split(",") if v.strip()])

        # Choix de la session
        session_frame = tk.Frame(self)
        session_frame.grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky="w")
        self.combo_session_hu = ttk.Combobox(session_frame, width=40)
        self.combo_session_hu.pack(side=tk.LEFT, padx=5)
        self.btn_refresh_sessions = tk.Button(session_frame, text="Rafraîchir sessions", command=self.refresh_sessions)
        self.btn_refresh_sessions.pack(side=tk.LEFT, padx=5)
        self.refresh_sessions()  # initialisation

        # Création des contrôles pour Work Center
        tk.Label(self, text="Work Center:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.combo_wc = ttk.Combobox(self, values=wc_values, width=20)
        self.combo_wc.grid(row=1, column=1, padx=5, pady=5)
        self.combo_wc.bind("<Delete>", self.delete_wc_value)

        # Création des contrôles pour Storage BIN
        tk.Label(self, text="Storage BIN:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.combo_sb = ttk.Combobox(self, values=sb_values, width=20)
        self.combo_sb.grid(row=2, column=1, padx=5, pady=5)
        self.combo_sb.bind("<Delete>", self.delete_sb_value)

        # Création des contrôles pour HU Type
        tk.Label(self, text="HU Type:").grid(row=3, column=0, padx=5, pady=5, sticky="w")
        self.combo_hu_type = ttk.Combobox(self, values=ht_values, width=20)
        self.combo_hu_type.grid(row=3, column=1, padx=5, pady=5)
        self.combo_hu_type.bind("<<ComboboxSelected>>", self.on_hu_type_change)
        # ToDo : Ajouter la partie delete + essayer de faire une fonction générique pour les suppressions de valeurs

        # Zone pour les options supplémentaires selon le type de HU
        self.frame_hu_details = tk.Frame(self)
        self.frame_hu_details.grid(row=4, column=0, columnspan=2, padx=5, pady=5, sticky="w")
        self.create_hu_detail_fields("PAC0011")

        # Bouton pour lancer la création des HU
        tk.Button(self, text="Lancer création HU", command=self.lancer_creation_hu).grid(row=5, column=0, columnspan=2, pady=10)

        # Zone de log pour afficher les status messages avec barre de défilement verticale
        self.text_log = tk.Text(self, height=10, width=70)
        self.text_log.grid(row=6, column=0, columnspan=2, padx=5, pady=5)
        self.scrollbar = tk.Scrollbar(self, orient="vertical", command=self.text_log.yview)
        self.scrollbar.grid(row=6, column=2, sticky="ns", padx=(0,5))
        self.text_log.configure(yscrollcommand=self.scrollbar.set)
        
        # La fenêtre parente doit attendre que la fenêtre soit fermée
        self.parent.wait_window(self)  

    def delete_wc_value(self, event):
        current_value = self.combo_wc.get().strip()
        if not current_value:
            return "break"
        if messagebox.askyesno("Confirmation", f"Voulez-vous vraiment supprimer '{current_value}' du Work Center ?"):
            values = list(self.combo_wc['values'])
            if current_value in values:
                values.remove(current_value)
                # Mise à jour de la combobox et de la config
                self.combo_wc['values'] = sorted(values)
                self.combo_wc.set("")
                self.config_parser.set("HU", "work_center", ", ".join(sorted(values)))
                with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                    self.config_parser.write(f)
            messagebox.showinfo("Info", f"'{current_value}' supprimé de Work Center.")
        return "break"

    def delete_sb_value(self, event):
        current_value = self.combo_sb.get().strip()
        if not current_value:
            return "break"
        if messagebox.askyesno("Confirmation", f"Voulez-vous vraiment supprimer '{current_value}' du Storage BIN ?"):
            values = list(self.combo_sb['values'])
            if current_value in values:
                values.remove(current_value)
                # Mise à jour de la combobox et de la config
                self.combo_sb['values'] = sorted(values)
                self.combo_sb.set("")
                self.config_parser.set("HU", "storage_bin", ", ".join(sorted(values)))
                with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                    self.config_parser.write(f)
            messagebox.showinfo("Info", f"'{current_value}' supprimé de Storage BIN.")
        return "break"

    def create_hu_detail_fields(self, hu_type):
        for widget in self.frame_hu_details.winfo_children():
            widget.destroy()
        if hu_type == "PAC0012":
            tk.Label(self.frame_hu_details, text="Nombre de HU à créer:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
            self.spin_hu_number = tk.Spinbox(self.frame_hu_details, from_=1, to=100, width=5)
            self.spin_hu_number.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        else:
            tk.Label(self.frame_hu_details, text="Liste des numéros de HU:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
            self.text_hu_list = tk.Text(self.frame_hu_details, height=4, width=40)
            self.text_hu_list.grid(row=1, column=0, columnspan=2, padx=5, pady=5)

    def on_hu_type_change(self, event):
        hu_type = self.combo_hu_type.get()
        self.create_hu_detail_fields(hu_type)

    def refresh_sessions(self):
        try:
            connection, _ = find_or_open_connection(self.parent.saplogon_path, self.parent.sap_env)
            nb_sessions = connection.Children.Count

            sessions_display = []
            self.sessions_map_hu = {}

            for i in range(nb_sessions):
                sess = connection.Children(i)
                try:
                    wnd_title = sess.findById("wnd[0]").Text
                except:
                    wnd_title = "(Sans titre)"
                display_name = f"Session {i} - {wnd_title}"
                sessions_display.append(display_name)
                self.sessions_map_hu[display_name] = i

            sessions_display.append("Nouvelle session")
            self.combo_session_hu['values'] = sessions_display

            if sessions_display:
                self.combo_session_hu.set(sessions_display[0])
            else:
                self.combo_session_hu.set("Nouvelle session")
        except Exception as e:
            self.combo_session_hu['values'] = ["Nouvelle session"]
            self.combo_session_hu.set("Nouvelle session")
            self.log(f"Erreur rafraîchissement sessions: {e}")

    def lancer_creation_hu(self):
        self.text_log.delete('1.0', tk.END)
        self.update_idletasks()
        wc = self.combo_wc.get().strip()
        sb = self.combo_sb.get().strip()
        hu_type = self.combo_hu_type.get().strip()

        self.update_hu_config(wc, sb)

        try:
            # Récupération de la session sélectionnée dans la fenêtre HU
            connection, _ = find_or_open_connection(self.parent.saplogon_path, self.parent.sap_env)
            session_choice = self.combo_session_hu.get().strip()
            nb_sessions = connection.Children.Count
            if session_choice == "Nouvelle session":
                if nb_sessions == 0:
                    self.log("Aucune session existante pour effectuer l'ouverture /o.")
                    return
                session = connection.Children(0)
            else:
                if session_choice not in self.sessions_map_hu:
                    self.log(f"La session '{session_choice}' n'existe pas.")
                    return
                idx = self.sessions_map_hu[session_choice]
                if idx < 0 or idx >= nb_sessions:
                    self.log(f"La session d'index {idx} n'existe pas.")
                    return
                session = connection.Children(idx)

            transaction_code = "/n/scwm/pack"
            session.findById("wnd[0]/tbar[0]/okcd").text = transaction_code
            session.findById("wnd[0]").sendVKey(0)
            time.sleep(2)

            try:
                session.findById("wnd[0]/usr/ctxtPA_WRKST").text = wc
            except Exception as e:
                self.log(f"Erreur lors du remplissage des champs wc: {e}")
                return
            try:
                session.findById("wnd[0]/usr/ctxtSOLGPLA-LOW").text = sb
            except Exception as e:
                self.log(f"Erreur lors du remplissage des champs sb: {e}")
                return
            try:
                session.findById("wnd[0]").sendVKey(8)
            except Exception as e:
                self.log(f"Erreur lors du click: {e}")
                return

            if hu_type == "PAC0012":
                nb = int(self.spin_hu_number.get())
                hu_list = ["" for i in range(1, nb+1)]
            else:
                raw_text = self.text_hu_list.get("1.0", tk.END).strip()
                if not raw_text:
                    self.log("Veuillez entrer au moins un numéro de HU.")
                    return
                hu_list = [s.strip() for s in raw_text.replace(",", "\n").split("\n") if s.strip()]

            for hu in hu_list:
                session.FindById("wnd[0]/usr/subSUB_SCANNER:/SCWM/SAPLUI_PACKING:0200/tabsTS_SCANNER/tabpHU_CREATE/ssubSS_SCANNER:/SCWM/SAPLUI_PACKING:0202/txt/SCWM/S_PACK_VIEW_SCANNER-DEST_HU_UI").Text = hu
                session.FindById("wnd[0]/usr/subSUB_SCANNER:/SCWM/SAPLUI_PACKING:0200/tabsTS_SCANNER/tabpHU_CREATE/ssubSS_SCANNER:/SCWM/SAPLUI_PACKING:0202/ctxt/SCWM/S_PACK_VIEW_SCANNER-DEST_PMAT_NO").Text = hu_type
                session.FindById("wnd[0]/usr/subSUB_SCANNER:/SCWM/SAPLUI_PACKING:0200/tabsTS_SCANNER/tabpHU_CREATE/ssubSS_SCANNER:/SCWM/SAPLUI_PACKING:0202/txt/SCWM/S_PACK_VIEW_SCANNER-DEST_ID").Text = sb
                session.findById("wnd[0]").sendVKey(8)
                time.sleep(1)
                status = session.findById("wnd[0]/sbar").text
                if status[-15:] != "was constructed":
                    session.findById("wnd[0]").sendVKey(2)
                self.log(f"HU {hu} : {status}")

        except Exception as e:
            self.log(f"Erreur lors de la création des HU: {e}")

    def update_hu_config(self, wc, sb):
        self.config_parser.read(CONFIG_FILE, encoding="utf-8")
        if not self.config_parser.has_section("HU"):
            self.config_parser.add_section("HU")
        existing_wc = self.config_parser.get("HU", "work_center", fallback="GPAK")
        existing_sb = self.config_parser.get("HU", "storage_bin", fallback="GR-ZONE, COOL")
        wc_list = {v.strip() for v in existing_wc.split(",") if v.strip()}
        sb_list = {v.strip() for v in existing_sb.split(",") if v.strip()}
        wc_list.add(wc)
        sb_list.add(sb)
        self.config_parser.set("HU", "work_center", ", ".join(wc_list))
        self.config_parser.set("HU", "storage_bin", ", ".join(sb_list))
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            self.config_parser.write(f)

    def log(self, message):
        self.text_log.insert(tk.END, message + "\n")
        self.text_log.see(tk.END)
        self.update_idletasks()

def main():
    app = MainApp()
    app.mainloop()

if __name__ == "__main__":
    main()
