import customtkinter
import xlwings as xw


class AppView(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        self.buttons = {}

        self.alert_window = None
        self.alert_window_close = None

        # Evénement déclenché lorsque l'utilisateur tente de fermer la fenêtre
        self.protocol("WM_DELETE_WINDOW", self.on_quit)

        customtkinter.set_appearance_mode("dark")

        self.title("Ezfacture excel")
        self.geometry(self.window_to_right_bottom("320", "550"))

        self.checkbox_frame_connexion = customtkinter.CTkFrame(self, fg_color="#242424")
        self.checkbox_frame_connexion.grid(
            row=0, column=0, padx=15, pady=(10, 0), sticky="w"
        )
        self.buttons["connexion"] = customtkinter.CTkButton(
            self.checkbox_frame_connexion,
            text="Démarrer",
            command=None,
            width=290,
            fg_color="#3498db",
            hover_color="#2980b9",
        )
        self.buttons["connexion"].grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.checkbox_frame_menu = customtkinter.CTkFrame(self, fg_color="#242424")
        self.checkbox_frame_menu.grid(
            row=1, column=0, padx=15, pady=(10, 0), sticky="nsw"
        )

        self.menu_nouveau = customtkinter.CTkOptionMenu(
            self.checkbox_frame_menu,
            values=[],
            command=None,
            corner_radius=10,
            width=70,
            fg_color="#34495e",
            button_color="#34495e",
            button_hover_color="#2c3e50",
            dropdown_fg_color=("#ffffff", "#ffffff"),
            dropdown_text_color="#333333",
            dropdown_hover_color="#bdc3c7",
        )
        self.menu_nouveau.grid(row=1, column=0, padx=5, pady=10, sticky="w")
        self.menu_nouveau.set("Nouveau")

        self.menu_ouvrir = customtkinter.CTkOptionMenu(
            self.checkbox_frame_menu,
            values=[],
            command=None,
            corner_radius=10,
            width=70,
            fg_color="#34495e",
            button_color="#34495e",
            button_hover_color="#2c3e50",
            dropdown_fg_color=("#ffffff", "#ffffff"),
            dropdown_text_color="#333333",
            dropdown_hover_color="#bdc3c7",
        )
        self.menu_ouvrir.grid(row=1, column=1, padx=5, pady=10, sticky="w")
        self.menu_ouvrir.set("Ouvrir")

        self.buttons["reglages"] = customtkinter.CTkButton(
            self.checkbox_frame_menu,
            text="Réglages",
            command=None,
            width=15,
            fg_color="#34495e",
        )
        self.buttons["reglages"].grid(row=1, column=2, padx=5, pady=10, sticky="w")

        self.buttons["aide"] = customtkinter.CTkButton(
            self.checkbox_frame_menu,
            text="?",
            command=None,
            width=15,
            fg_color="#34495e",
        )
        self.buttons["aide"].grid(row=1, column=3, padx=5, pady=10, sticky="w")

        self.checkbox_frame_actions = customtkinter.CTkFrame(self, fg_color="#242424")
        self.checkbox_frame_actions.grid(
            row=3, column=0, padx=15, pady=(10, 10), sticky="nsw"
        )
        self.buttons["save"] = customtkinter.CTkButton(
            self.checkbox_frame_actions,
            text="Prévisualiser (PDF)",
            command=None,
            width=290,
            fg_color="#3498db",
            hover_color="#2980b9",
            state="disabled",
            text_color_disabled="#2980b9",
        )
        self.buttons["save"].grid(row=0, column=0, padx=5, pady=10, sticky="ew")

        self.buttons["valider"] = customtkinter.CTkButton(
            self.checkbox_frame_actions,
            text="Valider le document",
            command=None,
            width=290,
            fg_color="#27ae60",
            hover_color="#2ecc71",
            state="disabled",
            text_color_disabled="#1a964e",
        )
        self.buttons["valider"].grid(row=1, column=0, padx=5, pady=10, sticky="ew")

        self.textbox_infos = customtkinter.CTkTextbox(
            master=self,
            width=280,
            height=115,
            corner_radius=10,
            border_width=1,
            border_color="#525252",
            wrap="word",
            fg_color="#242424",
            text_color="#bdc3c7",
            state="disabled",
        )
        self.textbox_infos.grid(row=2, column=0, padx=20, pady=10, sticky="nsew")

        self.textbox_feedback = customtkinter.CTkTextbox(
            master=self,
            width=280,
            height=150,
            corner_radius=10,
            wrap="word",
            fg_color="#242424",
        )
        self.textbox_feedback.grid(row=4, column=0, padx=10, pady=5, sticky="nsew")

    def show_restart(self):
        self.buttons["restart"] = customtkinter.CTkButton(
            self.checkbox_frame_actions,
            text="Redémarrer",
            command=None,
            fg_color="#e67e22",
            hover_color="#d35400",
        )
        self.buttons["restart"].grid(
            row=2, column=0, padx=5, pady=10, sticky="ew", columnspan=2
        )

    def set_actions(self, actions):
        for name, action in actions.items():
            if name in self.buttons:
                self.buttons[name].configure(command=action)

    def set_menu_nouveau(self, action, values):
        self.menu_nouveau.configure(command=action, values=values)

    def reset_menu_nouveau(self, choice):
        self.menu_nouveau.set("Nouveau")

    def set_menu_ouvrir(self, action, values):
        self.menu_ouvrir.configure(command=action, values=values)

    def reset_menu_ouvrir(self, choice):
        self.menu_ouvrir.set("Ouvrir")

    def block_ui(self):
        self.menu_nouveau.configure(state="disabled")
        self.menu_ouvrir.configure(state="disabled")
        self.block_boutons(["save", "valider", "reglages"])

    def enable_ui(self):
        self.menu_nouveau.configure(state="normal")
        self.menu_ouvrir.configure(state="normal")
        self.enable_boutons(["save", "valider", "reglages"])

    def block_boutons(self, boutons):
        """
        :param boutons: list
        :return: None
        """
        for bouton in boutons:
            self.buttons[bouton].configure(state="disabled")

    def enable_boutons(self, boutons):
        """
        :param boutons: list
        :return: None
        """
        for bouton in boutons:
            self.buttons[bouton].configure(state="normal")

    def window_to_right_bottom(self, width, height):
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = screen_width - int(width) - 20
        y = screen_height - int(height) - 40
        return f"{width}x{height}+{x}+{y}"

    def show_feedback(self, txt="", message_type="feedback", stack=False):
        """
        Affiche un message dans la zone textbox de feedback
        :param txt: String. texte à afficher.
        :param message_type: String. "feedback", "error", "success"
        :param stack: Bool. Empile les messages ou remplace le message précédent.
        :return: None
        """
        self.textbox_feedback.configure(state="normal")

        if message_type == "feedback":
            self.textbox_feedback.configure(text_color="#3498db")
        if message_type == "error":
            self.textbox_feedback.configure(text_color="#e67e22")
        if message_type == "success":
            self.textbox_feedback.configure(text_color="#27ae60")

        if not stack:
            self.textbox_feedback.delete("0.0", "end")
        self.textbox_feedback.insert("end", txt + "\n")

        self.textbox_feedback.configure(state="disabled")

    def show_infos(self, doc='', etat='', fichier='', date='', numero=''):
        """
        Affiche un message dans la zone textbox infos
        :param txt: String. texte à afficher.
        :return: None
        """
        self.textbox_infos.configure(state="normal")
        self.textbox_infos.delete("0.0", "end")

        txt=(f"Document : {doc.upper()}\n"
            f"Etat : {etat}\n"
            f"Fichier : {fichier}\n"
            f"Date de création : {date}\n"
            f"Numéro : {numero} \n")

        self.textbox_infos.insert("end", txt + "\n")
        self.textbox_infos.configure(state="disabled")

    def delete_messages(self, zone):
        if zone == "feedback":
            self.textbox_feedback.configure(state="normal")
            self.textbox_feedback.delete("0.0", "end")
            self.textbox_feedback.configure(state="disabled")
        if zone == "infos":
            self.textbox_infos.configure(state="normal")
            self.textbox_infos.delete("0.0", "end")
            self.textbox_infos.configure(state="disabled")

    def close(self):
        self.quit()
        self.destroy()

    def open_alert_quit(self, txt):
        if self.alert_window is None or not self.alert_window.winfo_exists():
            self.alert_window = AlertWindowQuit(self, txt=txt)
        else:
            self.alert_window.focus()

    def on_quit(self):
        self.open_alert_quit(
            txt="Fermer l'application ? \n"
            "Cette action fermera également le document \n"
            "sans l'enregistrer."
        )

    def open_alert_close(self, txt):
        if (
            self.alert_window_close is None
            or not self.alert_window_close.winfo_exists()
        ):
            self.alert_window_close = AlertWindowClose(self, txt=txt)
            # Bloquer l’exécution jusqu’à ce que la fenêtre soit détruite
            self.alert_window_close.wait_window()
        else:
            self.alert_window_close.focus()

        return self.alert_window_close.response


class AlertWindow(customtkinter.CTkToplevel):
    """
    Classe pour créer une fenêtre d'alerte générique
    """
    def __init__(self, parent, txt="", *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.parent = parent
        self.title("Fermeture de l'application")

        window_width = 340
        window_height = 160
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        position_x = (screen_width // 2) - (window_width // 2)
        position_y = (screen_height // 2) - (window_height // 2)

        self.geometry(f"{window_width}x{window_height}+{position_x}+{position_y}")

        self.label = customtkinter.CTkLabel(self, text=txt)
        self.label.pack(padx=20, pady=20)

        self.button_yes = customtkinter.CTkButton(self, text="OUI", command=self.on_yes)
        self.button_yes.pack(side="left", padx=20, pady=20)

        self.button_no = customtkinter.CTkButton(self, text="NON", command=self.on_no)
        self.button_no.pack(side="right", padx=20, pady=20)


class AlertWindowQuit(AlertWindow):
    """
    A instancier avant de quitter l'application
    """
    def __init__(self, parent, txt="Quitter sans enregistrer ?", *args, **kwargs):
        super().__init__(parent, txt, *args, **kwargs)
        self.response = None

    def on_yes(self):
        try:
            xw.apps.active.quit()
        except AttributeError:  # Excel est déjà fermé
            pass
        self.destroy()
        self.parent.close()

    def on_no(self):
        self.destroy()


class AlertWindowClose(AlertWindow):
    """
    A instancier pour valider une action
    """
    def __init__(self, parent, txt="Quitter sans enregistrer ?", *args, **kwargs):
        super().__init__(parent, txt, *args, **kwargs)
        self.response = None

    def on_yes(self):
        self.response = True
        self.destroy()

    def on_no(self):
        self.response = False
        self.destroy()
