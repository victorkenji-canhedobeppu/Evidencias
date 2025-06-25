# src/ui/custom_calendar.py

import customtkinter as ctk
import calendar
from datetime import datetime


class CustomCalendar(ctk.CTkToplevel):
    def __init__(self, master, on_date_select_callback):
        super().__init__(master)

        self.on_date_select_callback = on_date_select_callback
        self.today = datetime.now()
        self.current_year = self.today.year
        self.current_month = self.today.month

        # Configuração da janela pop-up
        self.lift()  # Levanta a janela para o topo
        self.attributes("-topmost", True)  # Mantém no topo
        self.title("Selecione a Data")
        self.geometry("280x290")
        self.resizable(False, False)

        # Centraliza a janela em relação ao pai
        x = master.winfo_x() + (master.winfo_width() / 2) - (280 / 2)
        y = master.winfo_y() + (master.winfo_height() / 2) - (290 / 2)
        self.geometry(f"+{int(x)}+{int(y)}")

        # Frame principal
        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.pack(fill="both", expand=True, padx=5, pady=5)

        self._setup_widgets()
        self._update_calendar()

        # Garante que a janela capture o foco
        self.grab_set()

    def _setup_widgets(self):
        # Frame do cabeçalho (Mês, Ano e botões de navegação)
        header_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        header_frame.pack(pady=5)

        prev_button = ctk.CTkButton(
            header_frame, text="<", command=self._prev_month, width=30
        )
        prev_button.pack(side="left", padx=5)

        self.month_year_label = ctk.CTkLabel(
            header_frame, text="", font=ctk.CTkFont(size=14, weight="bold")
        )
        self.month_year_label.pack(side="left", padx=10)

        next_button = ctk.CTkButton(
            header_frame, text=">", command=self._next_month, width=30
        )
        next_button.pack(side="left", padx=5)

        # Frame dos dias da semana
        days_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        days_frame.pack(pady=5)
        days = ["Dom", "Seg", "Ter", "Qua", "Qui", "Sex", "Sáb"]
        for day in days:
            label = ctk.CTkLabel(
                days_frame, text=day, width=35, height=25, font=ctk.CTkFont(size=12)
            )
            label.pack(side="left", padx=2, pady=2)

        # Frame dos botões dos dias
        self.calendar_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.calendar_frame.pack()

    def _update_calendar(self):
        # Atualiza o rótulo do mês/ano
        month_name = calendar.month_name[self.current_month]
        self.month_year_label.configure(
            text=f"{month_name.capitalize()}, {self.current_year}"
        )

        # Limpa os botões de dias antigos
        for widget in self.calendar_frame.winfo_children():
            widget.destroy()

        # Cria a matriz do calendário para o mês e ano atuais
        month_calendar = calendar.monthcalendar(self.current_year, self.current_month)

        # Cria os botões para cada dia
        for week in month_calendar:
            week_frame = ctk.CTkFrame(self.calendar_frame, fg_color="transparent")
            week_frame.pack()
            for day in week:
                if day == 0:
                    # Cria um label vazio para dias que não pertencem ao mês
                    label = ctk.CTkLabel(week_frame, text="", width=35, height=35)
                    label.pack(side="left", padx=2, pady=2)
                else:
                    # Cria um botão para cada dia
                    btn = ctk.CTkButton(
                        week_frame,
                        text=str(day),
                        command=lambda d=day: self._on_date_select(d),
                        width=35,
                        height=35,
                        fg_color=("#3B8ED0", "#1F6AA5"),  # Cor padrão do botão
                    )
                    # Destaca o dia de hoje
                    if (
                        self.current_year == self.today.year
                        and self.current_month == self.today.month
                        and day == self.today.day
                    ):
                        btn.configure(
                            fg_color=("#2FA572", "#146C46")
                        )  # Cor de destaque

                    btn.pack(side="left", padx=2, pady=2)

    def _prev_month(self):
        self.current_month -= 1
        if self.current_month == 0:
            self.current_month = 12
            self.current_year -= 1
        self._update_calendar()

    def _next_month(self):
        self.current_month += 1
        if self.current_month == 13:
            self.current_month = 1
            self.current_year += 1
        self._update_calendar()

    def _on_date_select(self, day):
        selected_date = datetime(self.current_year, self.current_month, day)
        self.on_date_select_callback(selected_date.strftime("%d/%m/%Y"))
        self.grab_release()  # Libera o foco
        self.destroy()
