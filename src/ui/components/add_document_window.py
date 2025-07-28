import customtkinter as ctk
from config.settings import DISCIPLINE_DEFAULT_TEXT
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
        self.selected_project_type = ctk.StringVar(value="Geometria")  # Default value

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

        # --- Dropdown para Tipo de Projeto ---
        project_type_label = ctk.CTkLabel(
            scrollable_frame, text="Tipo de Projeto:", font=ctk.CTkFont(weight="bold")
        )
        project_type_label.pack(fill="x", padx=10, pady=(10, 2))

        project_type_options = ["Geometria", "Topografia"]
        project_type_dropdown = ctk.CTkOptionMenu(
            scrollable_frame,
            values=project_type_options,
            variable=self.selected_project_type,
            command=self.update_project_text,
        )
        project_type_dropdown.pack(fill="x", padx=10, pady=(0, 10))

        self.projeto_default_texts = {
            "Geometria": (
                "No período de medição de {month}/{year}, foram realizados os "
                "serviços de [inserir serviços de geometria aqui, ex: 'levantamento "
                "planialtimétrico de áreas, batimetrias, georreferenciamento'], "
                "conforme cronograma e especificações técnicas."
            ),
            "Topografia": (
                "No período de medição de {month}/{year}, foram realizados os "
                "serviços de [inserir serviços de topografia aqui, ex: 'implantação "
                "de marcos topográficos, monitoramento de recalques, cálculo de volumes'], "
                "conforme cronograma e especificações técnicas."
            ),
        }

        self.project_title_label = ctk.CTkLabel(
            scrollable_frame,
            text=f"Projeto de {self.selected_project_type.get()}",
            font=ctk.CTkFont(weight="bold"),
        )
        self.project_title_label.pack(fill="x", padx=10, pady=(10, 2))

        self.project_textbox = ctk.CTkTextbox(scrollable_frame, height=80)
        self.project_textbox.pack(fill="x", expand=True, padx=10, pady=(0, 10))
        self.text_entries["Projeto"] = self.project_textbox

        self.update_project_text()  # Define o texto inicial com base no tipo de projeto padrão

        line_number = 2.0
        for discipline in self.disciplines:
            label = ctk.CTkLabel(
                scrollable_frame, text=discipline, font=ctk.CTkFont(weight="bold")
            )
            label.pack(fill="x", padx=10, pady=(10, 2))

            textbox = ctk.CTkTextbox(scrollable_frame, height=80)
            textbox.pack(fill="x", expand=True, padx=10, pady=(0, 10))
            textbox.insert(f"{line_number}", DISCIPLINE_DEFAULT_TEXT[discipline])
            self.text_entries[discipline] = textbox
            line_number += 1.0

        # --- Botão de ação ---
        save_button = ctk.CTkButton(
            self, text="Salvar no Documento", height=40, command=self.save_to_document
        )
        save_button.pack(side="bottom", fill="x", padx=15, pady=15)

    def update_project_text(self, *args):
        """Atualiza o texto padrão do campo 'Projeto' com base no tipo de projeto selecionado."""
        selected_type = self.selected_project_type.get()
        default_text = self.projeto_default_texts[selected_type].format(
            month=self.measurement_month, year=self.measurement_year
        )
        self.project_textbox.delete("1.0", "end")
        self.project_textbox.insert("1.0", default_text)

        self.update_project_title()

    def update_project_title(self):
        self.project_title_label.configure(
            text=f"Projeto de {self.selected_project_type.get()}"
        )

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
