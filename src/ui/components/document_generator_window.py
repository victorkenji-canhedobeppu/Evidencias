# src/ui/components/doc_generator_window.py
# VERSÃO FINAL: Simplificada para aceitar apenas input do utilizador.

from tkinter import filedialog, messagebox
import customtkinter as ctk
from datetime import datetime
from core.doc_creator import DocxCreator
import os

from ui.components.add_document_window import AddMeasurementWindow


class DocGeneratorWindow(ctk.CTkToplevel):
    """
    Janela para inserir dados para preencher o template .docx.
    Os campos começam sempre vazios.
    """

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
        self.title("Preencher Dados do Documento")
        self.geometry("700x750")

        # Guarda as informações recebidas para poder passá-las à próxima janela
        self.parent_app = parent
        self.project_type = project_type
        self.dataframe = dataframe
        self.source_file_path = source_file_path
        self.measurement_month = measurement_month
        self.measurement_year = measurement_year

        # Centraliza a janela
        parent_x, parent_y = parent.winfo_x(), parent.winfo_y()
        parent_width, parent_height = parent.winfo_width(), parent.winfo_height()
        x = parent_x + (parent_width // 2) - (700 // 2)
        y = parent_y + (parent_height // 2) - (650 // 2)
        self.geometry(f"+{x}+{y}")

        self.transient(parent)
        self.grab_set()

        # --- Frame Principal ---
        main_frame = ctk.CTkFrame(self)
        main_frame.pack(expand=True, fill="both", padx=15, pady=15)

        # --- Secção de Dados Principais ---
        main_data_frame = ctk.CTkFrame(main_frame)
        main_data_frame.pack(fill="x", padx=10, pady=10)
        main_data_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            main_data_frame,
            text="Dados Principais",
            font=ctk.CTkFont(size=14, weight="bold"),
        ).grid(row=0, column=0, columnspan=2, pady=(0, 10))

        self.main_fields = {}
        main_field_definitions = {
            "Código ANTT": "Código ANTT:",
            "Código Interno": "Código Interno:",
            "Emitente": "Emitente:",
            "Data Emissão Inicial": "Data Emissão:",
            "Rodovia": "Rodovia:",
            "Projetista": "Projetista:",
            "Trecho": "Trecho:",
            "Objeto": "Objeto:",
        }

        for i, (key, label_text) in enumerate(main_field_definitions.items()):
            label = ctk.CTkLabel(main_data_frame, text=label_text)
            label.grid(row=i + 1, column=0, padx=10, pady=5, sticky="w")

            entry = ctk.CTkEntry(main_data_frame, width=300)
            entry.grid(row=i + 1, column=1, padx=10, pady=5, sticky="ew")
            # A linha que preenchia os dados foi REMOVIDA
            self.main_fields[key] = entry

        # --- NOVA SECÇÃO: Conteúdo de Texto ---
        text_content_frame = ctk.CTkFrame(main_frame)
        text_content_frame.pack(fill="x", padx=10, pady=10)

        ctk.CTkLabel(text_content_frame, text="Texto para Título 1 (Introdução)").pack(
            anchor="w", padx=5
        )
        self.intro_textbox = ctk.CTkTextbox(text_content_frame, height=80)
        self.intro_textbox.pack(fill="x", expand=True, padx=5, pady=(0, 10))

        ctk.CTkLabel(text_content_frame, text="Texto para Título 2").pack(
            anchor="w", padx=5
        )
        self.title2_textbox = ctk.CTkTextbox(text_content_frame, height=80)
        self.title2_textbox.pack(fill="x", expand=True, padx=5, pady=(0, 5))

        # --- Botão de Ação ---
        generate_button = ctk.CTkButton(
            self, text="Salvar Documento", command=self.generate_action, height=40
        )
        generate_button.pack(pady=15, padx=15, fill="x")

    def generate_action(self):
        main_data = {key: entry.get() for key, entry in self.main_fields.items()}
        # --- Adiciona os novos textos ao dicionário de dados ---
        main_data["INTRO_TEXT"] = self.intro_textbox.get(
            "1.0", "end-1c"
        )  # Pega todo o texto
        main_data["TITLE2_TEXT"] = self.title2_textbox.get("1.0", "end-1c")

        table_data = []
        current_datetime = datetime.now()
        formatted_current_date = current_datetime.strftime("%d/%m/%Y")
        row = {
            "revisao": 0,
            "data": formatted_current_date,
            "descricao": "Emissão Inicial",
        }
        table_data.append(row)

        template_path = os.path.join(
            "src", "templates", "RSP-116RJ-000+000-GER-EXE-RT-Z9-001-R06.docx"
        )
        creator = DocxCreator(template_path=template_path)
        saved_docx_path = creator.generate_document(
            main_data=main_data, table_data=table_data
        )
        # 4. Pede para selecionar a pasta de evidências
        base_evidence_path = filedialog.askdirectory(
            title="Selecione a pasta raiz de evidências"
        )
        if not base_evidence_path:
            return

        # 5. Obtém a lista de disciplinas a partir dos nomes das pastas
        try:
            disciplines = [
                d
                for d in os.listdir(base_evidence_path)
                if os.path.isdir(os.path.join(base_evidence_path, d))
            ]
        except Exception as e:
            messagebox.showerror(
                "Erro", f"Não foi possível ler as pastas: {e}", parent=self
            )
            return

        # 6. Fecha esta janela e abre a de adicionar medição
        self.destroy()
        AddMeasurementWindow(
            parent=self.parent_app,
            docx_path=saved_docx_path,  # Usa o caminho do ficheiro recém-criado
            disciplines=sorted(disciplines),
            base_evidence_path=base_evidence_path,
            measurement_month=self.measurement_month,
            measurement_year=self.measurement_year,
        )
