# src/ui/components/add_measurement_window.py
import customtkinter as ctk
from core.append_to_doc import DocxAppender


class AddMeasurementWindow(ctk.CTkToplevel):
    """Janela para inserir textos para cada disciplina da medição."""

    def __init__(
        self,
        parent,
        docx_path: str,
        disciplines: list,
        base_evidence_path: str,
        measurement_month: int,
        measurement_year: int,
    ):
        super().__init__(parent)

        self.docx_path = docx_path
        self.disciplines = disciplines
        self.base_evidence_path = base_evidence_path
        self.measurement_month = measurement_month
        self.measurement_year = measurement_year

        self.title("Adicionar Textos da Medição")
        self.geometry("600x500")

        parent_x, parent_y = parent.winfo_x(), parent.winfo_y()
        x = parent_x + (parent.winfo_width() // 2) - (600 // 2)
        y = parent_y + (parent.winfo_height() // 2) - (500 // 2)
        self.geometry(f"+{x}+{y}")
        self.transient(parent)
        self.grab_set()

        # --- Criação dos campos de texto ---
        scrollable_frame = ctk.CTkScrollableFrame(
            self, label_text="Insira os textos para cada disciplina"
        )
        scrollable_frame.pack(expand=True, fill="both", padx=15, pady=15)

        self.text_entries = {}
        label = ctk.CTkLabel(
            scrollable_frame, text="Projeto", font=ctk.CTkFont(weight="bold")
        )
        label.pack(fill="x", padx=10, pady=(10, 2))

        textbox = ctk.CTkTextbox(scrollable_frame, height=80)
        textbox.pack(fill="x", expand=True, padx=10, pady=(0, 10))
        self.text_entries["Projeto"] = textbox

        for discipline in self.disciplines:
            label = ctk.CTkLabel(
                scrollable_frame, text=discipline, font=ctk.CTkFont(weight="bold")
            )
            label.pack(fill="x", padx=10, pady=(10, 2))

            textbox = ctk.CTkTextbox(scrollable_frame, height=80)
            textbox.pack(fill="x", expand=True, padx=10, pady=(0, 10))
            self.text_entries[discipline] = textbox

        # --- Botão de ação ---
        save_button = ctk.CTkButton(
            self, text="Salvar no Documento", height=40, command=self.save_to_document
        )
        save_button.pack(side="bottom", fill="x", padx=15, pady=15)

    def save_to_document(self):
        # Recolhe os textos dos campos
        user_texts = {
            discipline: textbox.get("1.0", "end-1c")
            for discipline, textbox in self.text_entries.items()
        }

        # Chama a lógica do backend
        appender = DocxAppender(
            self.docx_path, self.measurement_month, self.measurement_year
        )
        appender.append_measurement(
            disciplines=self.disciplines,
            user_texts=user_texts,
            base_evidence_path=self.base_evidence_path,
            measurement_month=self.measurement_month,
            measurement_year=self.measurement_year,
        )
        self.destroy()
