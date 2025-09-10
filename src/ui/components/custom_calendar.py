# src/ui/custom_calendar.py
# Layout original do utilizador, com as funcionalidades de posicionamento e
# seleção de mês/ano integradas.

import customtkinter as ctk
import calendar
from datetime import datetime


class CustomCalendar(ctk.CTkToplevel):
    def __init__(self, master, on_date_select_callback, anchor_widget=None):
        super().__init__(master)

        self.on_date_select_callback = on_date_select_callback
        self.anchor_widget = anchor_widget
        self.today = datetime.now()
        self.current_year = self.today.year
        self.current_month = self.today.month
        self.view = "days"  # Controla a vista atual: 'days', 'months', 'years'

        # --- Configuração da Janela (Layout sem borda) ---
        self.overrideredirect(True)
        self.lift()

        self.main_frame = ctk.CTkFrame(self, border_width=1)
        self.main_frame.pack(fill="both", expand=True, padx=1, pady=1)

        self._update_view()
        self._position_window()

        self.grab_set()
        self.bind("<FocusOut>", lambda e: self.destroy())
        self.bind("<Escape>", lambda e: self.destroy())
        self.focus_set()

    def _position_window(self):
        self.update_idletasks()
        width = self.winfo_reqwidth()
        height = self.winfo_reqheight()

        if self.anchor_widget:
            x = self.anchor_widget.winfo_rootx()
            y = self.anchor_widget.winfo_rooty() + self.anchor_widget.winfo_height() + 2

            screen_width = self.winfo_screenwidth()
            screen_height = self.winfo_screenheight()
            if x + width > screen_width:
                x = screen_width - width - 5
            if y + height > screen_height:
                y = self.anchor_widget.winfo_rooty() - height - 2
        else:
            x = self.master.winfo_x() + (self.master.winfo_width() / 2) - (width / 2)
            y = self.master.winfo_y() + (self.master.winfo_height() / 2) - (height / 2)

        self.geometry(f"+{int(x)}+{int(y)}")

    def _update_view(self):
        for widget in self.main_frame.winfo_children():
            widget.destroy()

        if self.view == "days":
            self._create_days_view()
        elif self.view == "months":
            self._create_months_view()
        elif self.view == "years":
            self._create_years_view()

        header_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        header_frame.pack(pady=5, padx=5)

        ctk.CTkButton(
            header_frame,
            text="<",
            command=lambda: setattr(self, "current_year", self.current_year - 1)
            or self._update_view(),
            width=30,
        ).pack(side="left")
        ctk.CTkButton(
            header_frame,
            text=str(self.current_year),
            command=lambda: self.switch_view("years"),
            font=ctk.CTkFont(size=14, weight="bold"),
        ).pack(side="left", padx=10, expand=True)
        ctk.CTkButton(
            header_frame,
            text=">",
            command=lambda: setattr(self, "current_year", self.current_year + 1)
            or self._update_view(),
            width=30,
        ).pack(side="left")

        month_grid = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        month_grid.pack(expand=True, fill="both", padx=5, pady=5)
        for i in range(1, 13):
            row, col = divmod(i - 1, 4)
            month_grid.grid_rowconfigure(row, weight=1)
            month_grid.grid_columnconfigure(col, weight=1)
            btn = ctk.CTkButton(
                month_grid,
                text=calendar.month_abbr[i],
                command=lambda m=i: self._select_month(m),
            )
            btn.grid(row=row, column=col, sticky="nsew", padx=2, pady=2)

    def _create_years_view(self):
        year_frame = ctk.CTkScrollableFrame(
            self.main_frame, label_text="Selecione o Ano", label_fg_color="transparent"
        )
        year_frame.pack(fill="both", expand=True, padx=5, pady=5)

        start_year, end_year = self.today.year - 10, self.today.year + 5
        current_year_index = 0
        for i, year in enumerate(range(start_year, end_year + 1)):
            if year == self.current_year:
                current_year_index = i
            ctk.CTkButton(
                year_frame, text=str(year), command=lambda y=year: self._select_year(y)
            ).pack(fill="x", pady=2)
        self.after(
            50,
            lambda: year_frame._parent_canvas.yview_moveto(
                float(current_year_index) / (end_year - start_year)
            ),
        )

    def switch_view(self, view):
        self.view = view
        self._update_view()

    def _select_month(self, month):
        self.current_month = month
        self.switch_view("days")

    def _select_year(self, year):
        self.current_year = year
        self.switch_view("months")

    def _prev_month(self):
        if self.current_month == 1:
            self.current_month, self.current_year = 12, self.current_year - 1
        else:
            self.current_month -= 1
        self._update_view()

    def _next_month(self):
        if self.current_month == 12:
            self.current_month, self.current_year = 1, self.current_year + 1
        else:
            self.current_month += 1
        self._update_view()

    def _on_date_select(self, day):
        selected_date = datetime(self.current_year, self.current_month, day)
        self.on_date_select_callback(selected_date.strftime("%d/%m/%Y"))
        self.grab_release()
        self.destroy()
