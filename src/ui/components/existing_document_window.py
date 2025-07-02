# src/ui/components/existing_evidence_window.py
# VERSÃO FINAL: Adicionados os checkboxes para controlar a criação de pastas,
# alinhando esta janela com a de "Nova Evidência".

import os
import customtkinter as ctk
from tkinter import filedialog, messagebox
from core.folder_creator import FolderCreator
from .add_document_window import AddMeasurementWindow


class ExistingDocumentWindow(ctk.CTkToplevel):
    """Janela de ações para quando um documento JÁ EXISTE."""

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

        # Guarda todos os dados recebidos da janela principal
        self.parent_app = parent
        self.project_type = project_type
        self.dataframe = dataframe
        self.source_file_path = source_file_path
        self.measurement_month = measurement_month
        self.measurement_year = measurement_year

        self.title("Registro de Evidência Existente")
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

        # --- CHECKBOXES PARA CONTROLO DE PASTAS ---
        checkbox_frame = ctk.CTkFrame(self, fg_color="transparent")
        checkbox_frame.pack(pady=10, padx=20, fill="x")

        self.sondagens_var = ctk.IntVar(value=1)
        self.sondagens_check = ctk.CTkCheckBox(
            checkbox_frame,
            text="Criar pasta 'Sondagens' (com subpastas)",
            variable=self.sondagens_var,
        )
        self.sondagens_check.pack(anchor="w", pady=5)

        self.ensaios_var = ctk.IntVar(value=1)
        self.ensaios_check = ctk.CTkCheckBox(
            checkbox_frame,
            text="Criar pasta 'Ensaios Especiais'",
            variable=self.ensaios_var,
        )
        self.ensaios_check.pack(anchor="w", pady=5)

        # --- Botões de Ação ---
        button_frame = ctk.CTkFrame(self, fg_color="transparent")
        button_frame.pack(pady=10, padx=20, fill="x", side="bottom")
        button_frame.grid_columnconfigure(0, weight=1)

        self.create_folders_button = ctk.CTkButton(
            button_frame,
            text="Executar Criação de Pastas",
            command=self.create_folders_action,
        )
        self.create_folders_button.grid(
            row=0, column=0, padx=(0, 5), pady=10, sticky="ew"
        )

        self.add_to_doc_button = ctk.CTkButton(
            button_frame, text="Adicionar ao Documento", command=self.add_to_doc_action
        )
        self.add_to_doc_button.grid(
            row=1, column=0, columnspan=2, padx=0, pady=(0, 10), sticky="ew"
        )

    def create_folders_action(self):
        """Executa a criação de pastas com base na seleção dos checkboxes."""
        creator = FolderCreator(
            project_type=self.project_type,
            df=self.dataframe,
            file_path=self.source_file_path,
        )

        create_sondagens = bool(self.sondagens_var.get())
        create_ensaios = bool(self.ensaios_var.get())

        success, message, _ = creator.create_folders_for_active_data(
            create_sondagens=create_sondagens, create_ensaios=create_ensaios
        )

        if success:
            messagebox.showinfo("Processo Concluído", message, parent=self)
        else:
            messagebox.showerror("Erro ao Criar Pastas", message, parent=self)

    def add_to_doc_action(self):
        docx_path = filedialog.askopenfilename(
            title="Selecione o documento .docx existente",
            filetypes=[("Documento Word", "*.docx")],
        )
        if not docx_path:
            return

        base_evidence_path = filedialog.askdirectory(
            title="Selecione a pasta raiz de evidências (ex: ..._Evidencias (1))"
        )
        if not base_evidence_path:
            return

        try:
            disciplines_found = [
                d
                for d in os.listdir(base_evidence_path)
                if os.path.isdir(os.path.join(base_evidence_path, d))
            ]
        except Exception as e:
            messagebox.showerror(
                "Erro",
                f"Não foi possível ler as pastas do diretório selecionado.\n\nErro: {e}",
                parent=self,
            )
            return

        if not disciplines_found:
            messagebox.showwarning(
                "Atenção",
                "Nenhuma subpasta de disciplina foi encontrada no diretório selecionado.",
                parent=self,
            )
            return

        self.destroy()
        AddMeasurementWindow(
            parent=self.parent_app,
            docx_path=docx_path,
            disciplines=sorted(disciplines_found),
            base_evidence_path=base_evidence_path,
            measurement_month=self.measurement_month,
            measurement_year=self.measurement_year,
        )
