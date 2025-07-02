# ui/app.py
# VERSÃO FINAL: O filtro de data foi simplificado para usar o novo
# componente MonthYearPicker.

import customtkinter as ctk
from tkinter import filedialog, messagebox, ttk
from tkinter import font as tkfont
from datetime import datetime
import pandas as pd
from math import ceil
import calendar

import os
import base64
import io
from PIL import Image, ImageTk

# --- IMPORTAÇÕES DO PROJETO ---
from core.folder_creator import FolderCreator
from core.excel_reader import Excel_Reader
from ui.components.custom_month_year import MonthYearPicker
from ui.components.existing_document_window import ExistingDocumentWindow
from ui.components.new_document_window import NewDocumentWindow

# Importa o novo seletor e remove a referência ao calendário antigo
from .image import BASE64_IMAGE


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Relatório Personalizado")
        self.geometry("1100x900")
        self.minsize(950, 600)

        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        # --- Variáveis de Estado ---
        self.current_file_path = None
        self.date_column = "VERSÃO ATUAL - DATA"
        self.original_df = pd.DataFrame()
        self.active_df = pd.DataFrame()
        self.current_page = 1
        self.ROWS_PER_PAGE = 50
        self.footer_image = None
        self.project_type = ctk.StringVar(value="ARTESP")

        # Variáveis para guardar o mês e ano selecionados
        self.selected_month = ctk.StringVar()
        self.selected_year = ctk.StringVar()

        self.month_map = {
            name: num for num, name in enumerate(calendar.month_name) if num > 0
        }

        # (O resto do __init__ e Estilos permanecem os mesmos)
        style = ttk.Style(self)
        style.theme_use("default")
        self.theme_bg_color = self._apply_appearance_mode(
            ctk.ThemeManager.theme["CTkFrame"]["fg_color"]
        )
        self.theme_text_color = self._apply_appearance_mode(
            ctk.ThemeManager.theme["CTkLabel"]["text_color"]
        )
        self.theme_selection_color = self._apply_appearance_mode(
            ctk.ThemeManager.theme["CTkButton"]["fg_color"]
        )
        self.even_row_color = self._apply_appearance_mode(("#F7F9FA", "#2D2D2D"))
        style.configure(
            "Treeview",
            background=self.theme_bg_color,
            foreground=self.theme_text_color,
            fieldbackground=self.theme_bg_color,
            borderwidth=0,
            rowheight=28,
        )
        header_bg_color = self._apply_appearance_mode(("#F0F2F5", "#2D2D2D"))
        style.configure(
            "Treeview.Heading",
            background=header_bg_color,
            foreground=self.theme_text_color,
            relief="flat",
        )
        style.map(
            "Treeview.Heading",
            background=[("active", header_bg_color)],
            relief=[("active", "flat"), ("pressed", "flat")],
        )
        style.configure("evenrow.Treeview", background=self.even_row_color)
        style.map("Treeview", background=[("selected", self.theme_selection_color)])
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        controls_container = ctk.CTkFrame(self)
        controls_container.grid(row=0, column=0, padx=10, pady=(10, 0), sticky="ew")
        controls_container.grid_columnconfigure(0, weight=1)

        file_ops_frame = ctk.CTkFrame(controls_container, fg_color="transparent")
        file_ops_frame.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
        file_ops_frame.grid_columnconfigure(4, weight=1)
        self.load_button = ctk.CTkButton(
            file_ops_frame, text="Carregar Arquivo", command=self.load_file
        )
        self.load_button.grid(row=0, column=0, padx=(0, 10))
        self.project_selector_label = ctk.CTkLabel(file_ops_frame, text="Projeto:")
        self.project_selector_label.grid(row=0, column=1, padx=(10, 5))
        self.project_selector = ctk.CTkOptionMenu(
            file_ops_frame, values=["ARTESP", "ANTT"], variable=self.project_type
        )
        self.project_selector.grid(row=0, column=2, padx=(0, 10))
        self.document_button = ctk.CTkButton(
            file_ops_frame,
            text="Documento",
            command=self.on_document_button_click,
            state="disabled",
        )
        self.document_button.grid(row=0, column=3, padx=(10, 0))
        self.file_label = ctk.CTkLabel(
            file_ops_frame, text="Nenhum arquivo carregado", text_color="gray"
        )
        self.file_label.grid(row=0, column=5, padx=(10, 0), sticky="e")

        # --- Linha de Filtragem SIMPLIFICADA ---
        filter_ops_frame = ctk.CTkFrame(controls_container, fg_color="transparent")
        filter_ops_frame.grid(row=1, column=0, sticky="ew", padx=5, pady=5)

        # Botão para abrir o seletor
        self.picker_button = ctk.CTkButton(
            filter_ops_frame,
            text="Selecionar Mês/Ano",
            command=self.open_month_year_picker,
        )
        self.picker_button.pack(side="left")

        # Rótulos para mostrar a data selecionada
        self.selected_date_label = ctk.CTkLabel(
            filter_ops_frame, text="Nenhuma data selecionada."
        )
        self.selected_date_label.pack(side="left", padx=15)

        # Botões de Ação
        self.filter_button = ctk.CTkButton(
            filter_ops_frame, text="Filtrar", command=self.filter_data, state="disabled"
        )
        self.filter_button.pack(side="left", padx=5)
        self.clear_filter_button = ctk.CTkButton(
            filter_ops_frame,
            text="Limpar Filtro",
            command=self.clear_filter,
            state="disabled",
        )
        self.clear_filter_button.pack(side="left", padx=5)

        # (O resto da UI de Paginação, Tabela e Rodapé permanece o mesmo)
        main_frame = ctk.CTkFrame(self)
        main_frame.grid(row=1, column=0, padx=10, pady=5, sticky="nsew")
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_rowconfigure(1, weight=1)
        self.pagination_frame = ctk.CTkFrame(main_frame)
        self.pagination_frame.grid(row=0, column=0, padx=10, pady=5, sticky="ew")
        self.prev_button = ctk.CTkButton(
            self.pagination_frame,
            text="< Anterior",
            command=self.prev_page,
            state="disabled",
        )
        self.prev_button.pack(side="left", padx=5, pady=5)
        self.page_label = ctk.CTkLabel(self.pagination_frame, text="")
        self.page_label.pack(side="left", padx=10, pady=5)
        self.next_button = ctk.CTkButton(
            self.pagination_frame,
            text="Próxima >",
            command=self.next_page,
            state="disabled",
        )
        self.next_button.pack(side="left", padx=(0, 20), pady=5)
        self.goto_label = ctk.CTkLabel(self.pagination_frame, text="Ir para pág:")
        self.goto_label.pack(side="left", pady=5)
        self.page_entry = ctk.CTkEntry(self.pagination_frame, width=50)
        self.page_entry.pack(side="left", padx=5, pady=5)
        self.page_entry.bind("<Return>", self.goto_page)
        self.goto_button = ctk.CTkButton(
            self.pagination_frame, text="Ir", width=40, command=self.goto_page
        )
        self.goto_button.pack(side="left", pady=5)
        table_frame = ctk.CTkFrame(main_frame)
        table_frame.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)
        scrollbar_x = ctk.CTkScrollbar(table_frame, orientation="horizontal")
        scrollbar_y = ctk.CTkScrollbar(table_frame, orientation="vertical")
        self.tree = ttk.Treeview(
            table_frame,
            show="headings",
            yscrollcommand=scrollbar_y.set,
            xscrollcommand=scrollbar_x.set,
        )
        scrollbar_x.pack(side="bottom", fill="x")
        scrollbar_y.pack(side="right", fill="y")
        self.tree.pack(fill="both", expand=True)
        scrollbar_x.configure(command=self.tree.xview)
        scrollbar_y.configure(command=self.tree.yview)
        footer_frame = ctk.CTkFrame(self, fg_color="transparent")
        footer_frame.grid(row=2, column=0, padx=10, pady=(5, 10), sticky="ew")
        footer_frame.grid_columnconfigure(1, weight=1)
        try:
            image_data = base64.b64decode(BASE64_IMAGE)
            image_stream = io.BytesIO(image_data)
            img_pil = Image.open(image_stream)
            aspect_ratio = img_pil.width / img_pil.height
            new_height = 35
            new_width = int(new_height * aspect_ratio)
            resized_img = img_pil.resize(
                (new_width, new_height), Image.Resampling.LANCZOS
            )
            self.footer_image = ImageTk.PhotoImage(resized_img)
            image_label = ctk.CTkLabel(footer_frame, image=self.footer_image, text="")
            image_label.grid(row=0, column=0, sticky="w", padx=10)
        except Exception:
            image_label = ctk.CTkLabel(footer_frame, text="[Logo]", text_color="gray")
            image_label.grid(row=0, column=0, sticky="w", padx=10)
        copyright_text = f"© {datetime.now().year} Canhedo Beppu Engenheiros Associados LTDA - Todos os direitos reservados."
        copyright_label = ctk.CTkLabel(
            footer_frame,
            text=copyright_text,
            font=ctk.CTkFont(size=11),
            text_color="gray",
        )
        copyright_label.grid(row=0, column=1, sticky="e", padx=10)
        self.update_paginated_view()

    # --- NOVA FUNÇÃO PARA ABRIR O SELETOR ---
    def open_month_year_picker(self):
        def on_selection(month, year):
            self.selected_month.set(month)
            self.selected_year.set(year)
            self.selected_date_label.configure(text=f"Data selecionada: {month}/{year}")

        MonthYearPicker(self, on_selection)

    # (A função on_document_button_click e outras permanecem as mesmas)
    def on_document_button_click(self):
        if self.original_df.empty:
            messagebox.showwarning(
                "Atenção", "Por favor, carregue um arquivo primeiro."
            )
            return

        month_str = self.selected_month.get()
        year_str = self.selected_year.get()

        # 2. Valida se uma data foi selecionada para a medição
        if not month_str or not year_str:
            messagebox.showwarning(
                "Data da Medição Necessária",
                "Por favor, use o botão 'Selecionar Mês/Ano' para definir a data da medição antes de continuar.",
            )
            return

        selected_month_num = self.month_map[month_str]
        
        answer = messagebox.askquestion(
            "Verificação de Documento",
            "Já existe um documento de evidência para este registro?",
            icon="question",
            type="yesno",
            default="yes",
        )
        action_data = {
            "parent": self,
            "project_type": self.project_type.get(),
            "dataframe": self.active_df,
            "source_file_path": self.current_file_path,
            "measurement_month": selected_month_num,
            "measurement_year": int(year_str),
        }
        if answer == "no":
            NewDocumentWindow(**action_data)
        else:
            ExistingDocumentWindow(**action_data)

    def clear_filter(self):
        if self.original_df.empty:
            return
        self.selected_month.set("")
        self.selected_year.set("")
        self.selected_date_label.configure(text="Nenhuma data selecionada.")
        self.active_df = self.original_df
        self.current_page = 1
        self.update_paginated_view()

    def load_file(self):
        file_path = filedialog.askopenfilename(
            title="Selecione o arquivo Excel",
            filetypes=(("Arquivos Excel", "*.xlsx *.xls"),),
        )
        if not file_path:
            return
        self.current_file_path = file_path
        self.file_label.configure(text="Carregando...", text_color="orange")
        self.update_idletasks()
        try:
            reader = Excel_Reader(file_path=file_path)
            data_by_sheet = reader.get_data_as_dataframe(
                date_column_name=self.date_column
            )
            if not data_by_sheet:
                raise ValueError("O leitor de Excel não retornou dados válidos.")
            self.original_df = pd.concat(
                list(data_by_sheet.values()), ignore_index=True
            )
            self.active_df = self.original_df
            self.file_label.configure(text=file_path.split("/")[-1], text_color="green")
            self.filter_button.configure(state="normal")
            self.clear_filter_button.configure(state="normal")
            self.document_button.configure(state="normal")
            self.current_page = 1
            self.update_paginated_view()
        except Exception as e:
            messagebox.showerror(
                "Erro ao Carregar Arquivo",
                f"Não foi possível ler ou processar o arquivo.\n\nDetalhe: {e}",
            )
            self.original_df, self.active_df = pd.DataFrame(), pd.DataFrame()
            self.file_label.configure(text="Falha ao carregar", text_color="red")
            self.filter_button.configure(state="disabled")
            self.clear_filter_button.configure(state="disabled")
            self.document_button.configure(state="disabled")
            self.update_paginated_view()

    def filter_data(self):
        if self.original_df.empty:
            messagebox.showwarning("Atenção", "Carregue um arquivo Excel primeiro.")
            return

        # Pega os valores das novas variáveis
        month_str = self.selected_month.get()
        year_str = self.selected_year.get()

        if not month_str or not year_str:
            messagebox.showwarning(
                "Data Incompleta",
                "Por favor, selecione um mês e um ano antes de filtrar.",
            )
            return

        selected_month = self.month_map[month_str]
        selected_year = int(year_str)

        df_copy = self.original_df.copy()
        if self.date_column not in df_copy.columns:
            messagebox.showwarning(
                "Coluna Não Encontrada",
                f"A coluna de data '{self.date_column}' não foi encontrada.",
            )
            return

        df_copy["datetime_col"] = pd.to_datetime(
            df_copy[self.date_column], dayfirst=True, errors="coerce"
        )
        df_copy.dropna(subset=["datetime_col"], inplace=True)

        year_mask = df_copy["datetime_col"].dt.year == selected_year
        month_mask = df_copy["datetime_col"].dt.month == selected_month
        final_mask = year_mask & month_mask

        filtered_results = df_copy[final_mask].drop(columns=["datetime_col"])

        self.active_df = filtered_results
        if filtered_results.empty:
            messagebox.showinfo(
                "Busca Concluída",
                "Nenhum registro encontrado para o Mês/Ano selecionado.",
            )

        self.current_page = 1
        self.update_paginated_view()

    # (O resto das funções permanecem as mesmas)
    def goto_page(self, event=None):
        page_num_str = self.page_entry.get()
        if not page_num_str:
            return
        try:
            page_num = int(page_num_str)
            total_pages = (
                ceil(len(self.active_df) / self.ROWS_PER_PAGE)
                if len(self.active_df) > 0
                else 1
            )
            if 1 <= page_num <= total_pages:
                self.current_page = page_num
                self.update_paginated_view()
            else:
                messagebox.showwarning(
                    "Página Inválida",
                    f"Por favor, insira um número entre 1 e {total_pages}.",
                )
        except ValueError:
            messagebox.showerror(
                "Entrada Inválida", "Por favor, insira apenas números."
            )
        finally:
            self.page_entry.delete(0, "end")

    def display_dataframe(self, df: pd.DataFrame):
        self.tree.delete(*self.tree.get_children())
        if df.empty:
            self.tree["column"] = []
            return
        self.tree["column"] = list(df.columns)
        self.tree["show"] = "headings"
        self.tree.column("#0", width=0, stretch=False)
        for column in self.tree["column"]:
            self.tree.heading(column, text=column)
            self.tree.column(column, width=120, anchor="w", stretch=False)
        df_rows = df.astype(str).to_numpy().tolist()
        for i, row in enumerate(df_rows):
            tags = ("evenrow",) if i % 2 == 0 else ()
            self.tree.insert("", "end", iid=i, values=row, tags=tags)
        self.autosize_columns(df)
        self.update_idletasks()

    def prev_page(self):
        if self.current_page > 1:
            self.current_page -= 1
            self.update_paginated_view()

    def next_page(self):
        total_rows = len(self.active_df)
        total_pages = ceil(total_rows / self.ROWS_PER_PAGE) if total_rows > 0 else 1
        if self.current_page < total_pages:
            self.current_page += 1
            self.update_paginated_view()

    def update_paginated_view(self):
        total_rows = len(self.active_df)
        nav_state = "normal" if total_rows > 0 else "disabled"
        if total_rows == 0:
            self.display_dataframe(pd.DataFrame())
            self.page_label.configure(text="Nenhum registro para exibir")
            self.pagination_frame.grid_remove()
        else:
            self.pagination_frame.grid()
            total_pages = ceil(total_rows / self.ROWS_PER_PAGE)
            if self.current_page > total_pages:
                self.current_page = total_pages
            start_index = (self.current_page - 1) * self.ROWS_PER_PAGE
            end_index = start_index + self.ROWS_PER_PAGE
            page_df = self.active_df.iloc[start_index:end_index]
            self.display_dataframe(page_df)
            self.page_label.configure(
                text=f"Página {self.current_page} de {total_pages} ({total_rows} registros)"
            )

        is_last_page = self.current_page >= (
            ceil(total_rows / self.ROWS_PER_PAGE) if total_rows > 0 else 1
        )
        self.prev_button.configure(
            state=nav_state if self.current_page > 1 else "disabled"
        )
        self.next_button.configure(state=nav_state if not is_last_page else "disabled")
        self.goto_button.configure(state=nav_state)
        self.page_entry.configure(state=nav_state)

    def autosize_columns(self, df: pd.DataFrame):
        font = tkfont.Font(font="TkDefaultFont")
        for col in self.tree["columns"]:
            max_width = font.measure(col)
            if col in df.columns:
                for cell_value in df[col].astype(str).dropna().head(100):
                    cell_width = font.measure(cell_value)
                    if cell_width > max_width:
                        max_width = cell_width
            max_width = min(max_width + 20, 500)
            self.tree.column(
                col, width=max_width, minwidth=60, anchor="w", stretch=False
            )
