# src/ui/components/new_evidence_window.py
# Baseado no seu código, adaptado para o novo fluxo com botões separados.

import customtkinter as ctk
from tkinter import messagebox
from core.folder_creator import FolderCreator
from ui.components.document_generator_window import DocGeneratorWindow


class NewDocumentWindow(ctk.CTkToplevel):
    def __init__(
        self,
        parent,
        project_type: str,
        dataframe: object,
        source_file_path: str,
        measurement_month: int,
        measurement_year: int,
    ):
        super().__init__(parent)

        self.parent_app = parent
        self.project_type = project_type
        self.dataframe = dataframe
        self.source_file_path = source_file_path
        self.measurement_month = measurement_month
        self.measurement_year = measurement_year

        self.title("Novo Registro de Evidência")
        self.geometry("450x300")  # Aumentei a altura para os checkboxes

        parent_x, parent_y = parent.winfo_x(), parent.winfo_y()
        x = parent_x + (parent.winfo_width() // 2) - (450 // 2)
        y = parent_y + (parent.winfo_height() // 2) - (300 // 2)
        self.geometry(f"+{x}+{y}")

        self.transient(parent)
        self.grab_set()

        self.main_label = ctk.CTkLabel(
            self,
            text="Opções de Criação de Pastas:",
            font=ctk.CTkFont(size=14, weight="bold"),
        )
        self.main_label.pack(pady=(15, 10))

        # --- NOVOS CHECKBOXES ---
        checkbox_frame = ctk.CTkFrame(self, fg_color="transparent")
        checkbox_frame.pack(pady=10, padx=20, fill="x")

        self.sondagens_var = ctk.IntVar(value=1)  # Começa marcado por padrão
        self.sondagens_check = ctk.CTkCheckBox(
            checkbox_frame,
            text="Criar pasta 'Sondagens'",
            variable=self.sondagens_var,
        )
        self.sondagens_check.pack(anchor="w", pady=5)

        self.ensaios_var = ctk.IntVar(value=1)  # Começa marcado por padrão
        self.ensaios_check = ctk.CTkCheckBox(
            checkbox_frame,
            text="Criar pasta 'Ensaios Especiais'",
            variable=self.ensaios_var,
        )
        self.ensaios_check.pack(anchor="w", pady=5)

        # --- Botões de Ação ---
        button_frame = ctk.CTkFrame(self, fg_color="transparent")
        button_frame.pack(pady=20, padx=20, fill="x", side="bottom")
        button_frame.grid_columnconfigure(0, weight=1)

        self.create_folders_button = ctk.CTkButton(
            button_frame,
            text="Executar Criação de Pastas",
            command=self.create_folders_action,
        )
        self.create_folders_button.grid(row=0, column=0, pady=(0, 10), sticky="ew")

        self.generate_doc_button = ctk.CTkButton(
            button_frame, text="Gerar Documento", command=self.generate_doc_action
        )
        self.generate_doc_button.grid(row=1, column=0, pady=10, sticky="ew")

    def create_folders_action(self):
        """Executa a criação de pastas com base na seleção dos checkboxes."""
        creator = FolderCreator(
            project_type=self.project_type,
            df=self.dataframe,
            file_path=self.source_file_path,
        )

        # Obtém o estado dos checkboxes (1 para sim, 0 para não) e converte para booleano
        create_sondagens = bool(self.sondagens_var.get())
        create_ensaios = bool(self.ensaios_var.get())

        success, message, _ = creator.create_folders_for_active_data(
            create_sondagens=create_sondagens, create_ensaios=create_ensaios
        )

        if success:
            messagebox.showinfo("Processo Concluído", message, parent=self)
        else:
            messagebox.showerror("Erro ao Criar Pastas", message, parent=self)

    def generate_doc_action(self):
        """Fecha esta janela e abre a de geração de documento, passando o contexto."""
        self.destroy()
        DocGeneratorWindow(
            parent=self.parent_app,
            project_type=self.project_type,
            dataframe=self.dataframe,
            source_file_path=self.source_file_path,
            measurement_month=self.measurement_month,
            measurement_year=self.measurement_year,
        )
