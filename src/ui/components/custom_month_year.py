# src/ui/components/month_year_picker.py
# Um novo seletor de Mês/Ano, mais limpo e moderno.

from tkinter import messagebox
import customtkinter as ctk
import calendar
from datetime import datetime


class MonthYearPicker(ctk.CTkToplevel):
    def __init__(self, parent, on_select_callback):
        super().__init__(parent)

        self.on_select_callback = on_select_callback

        self.title("Selecionar Mês/Ano")
        self.geometry("300x300")

        # Centraliza a janela
        parent_x, parent_y = parent.winfo_x(), parent.winfo_y()
        x = parent_x + (parent.winfo_width() // 2) - (300 // 2)
        y = parent_y + (parent.winfo_height() // 2) - (200 // 2)
        self.geometry(f"+{x}+{y}")

        self.transient(parent)
        self.grab_set()

        # --- Widgets ---
        main_frame = ctk.CTkFrame(self)
        main_frame.pack(expand=True, fill="both", padx=15, pady=15)

        # Seletor de Mês
        month_label = ctk.CTkLabel(main_frame, text="Mês:")
        month_label.pack(pady=(5, 2))

        self.month_var = ctk.StringVar(value=calendar.month_name[datetime.now().month])
        month_menu = ctk.CTkOptionMenu(
            main_frame,
            variable=self.month_var,
            values=[m for m in calendar.month_name if m],
        )
        month_menu.pack(pady=(0, 10), padx=20, fill="x")

        # Seletor de Ano
        year_label = ctk.CTkLabel(main_frame, text="Ano:")
        year_label.pack(pady=(5, 2))

        self.year_entry = ctk.CTkEntry(main_frame, placeholder_text="AAAA")
        self.year_entry.insert(0, str(datetime.now().year))
        self.year_entry.pack(pady=(0, 10), padx=20, fill="x")

        # Botão de Confirmação
        confirm_button = ctk.CTkButton(
            self, text="Confirmar", command=self._on_confirm, height=40
        )
        confirm_button.pack(side="bottom", pady=15, padx=15, fill="x")

    def _on_confirm(self):
        """Valida os dados e chama o callback."""
        selected_month = self.month_var.get()
        year_str = self.year_entry.get()

        if not year_str.isdigit() or len(year_str) != 4:
            messagebox.showwarning(
                "Ano Inválido",
                "Por favor, insira um ano válido com 4 dígitos.",
                parent=self,
            )
            return

        self.on_select_callback(selected_month, year_str)
        self.destroy()
