from builtins import getattr
import os
import sys
import logging
import threading
import traceback
from datetime import datetime

import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox
from openpyxl import load_workbook

if getattr(sys, "frozen", False):
    # Программа запущена как EXE
    BASE_DIR = os.path.dirname(sys.executable)
else:
    # Программа запущена через Python
    BASE_DIR = os.path.dirname(
        os.path.dirname(os.path.abspath(__file__))
    )

# ============================================================
# НАСТРОЙКИ
# ============================================================

TOUR_ROWS = {
    1: range(7, 13),
    2: range(15, 21),
    3: range(25, 31),
    4: range(33, 39),
    5: range(43, 49),
    6: range(51, 57),
    7: range(61, 67),
    8: range(69, 75),
}

TOTAL_TOURS = len(TOUR_ROWS)
MATCHES_PER_TOUR = 6

def get_tour_rows(tour_number):
    if tour_number not in TOUR_ROWS:
        raise ValueError(f"Некорректный номер тура: {tour_number}")

    return TOUR_ROWS[tour_number]

MAIN_FILE = os.path.join(BASE_DIR, "data", "main.xlsx")
PLAYERS_FILE = os.path.join(BASE_DIR, "data", "export_players.xlsx")

APPLICATION_TEMPLATE = os.path.join(
    BASE_DIR, "templates", "application.xlsx"
)

PROTOCOL_TEMPLATE = os.path.join(
    BASE_DIR, "templates", "protocol.xlsx"
)

MAIN_SHEET = "Расписание игр"
SCHEDULE_SHEET = "3т2026 Лига 1"



# ============================================================
# ПУТИ К ФАЙЛАМ
# ============================================================

if getattr(sys, "frozen", False):
    # Если программа запущена как EXE
    BASE_DIR = os.path.dirname(sys.executable)
    RESOURCE_DIR = sys._MEIPASS
else:
    # Если программа запущена через Python
    BASE_DIR = os.path.dirname(
        os.path.dirname(os.path.abspath(__file__))
    )
    RESOURCE_DIR = BASE_DIR


def resource_path(relative_path):
    """Путь к файлам ресурсов, в том числе внутри EXE."""
    return os.path.join(RESOURCE_DIR, relative_path)


OUTPUT_DIR = os.path.join(BASE_DIR, "outputs")
LOG_DIR = os.path.join(BASE_DIR, "logs")

os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(LOG_DIR, exist_ok=True)


# ============================================================
# ЛОГИРОВАНИЕ
# ============================================================

LOG_FILE = os.path.join(LOG_DIR, "superleague.log")

logger = logging.getLogger("SuperLeague")
logger.setLevel(logging.INFO)

if not logger.handlers:
    file_handler = logging.FileHandler(
        LOG_FILE,
        encoding="utf-8"
    )

    formatter = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(message)s"
    )

    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)


# ============================================================
# ДАТА И ВРЕМЯ ТУРА
# ============================================================

def get_tour_date(tour_number):
    """Возвращает дату проведения выбранного тура."""

    dates = {
        1: datetime(2026, 11, 5),
        2: datetime(2026, 11, 5),

        3: datetime(2026, 11, 6),
        4: datetime(2026, 11, 6),

        5: datetime(2026, 11, 7),
        6: datetime(2026, 11, 7),

        7: datetime(2026, 11, 8),
        8: datetime(2026, 11, 8),
    }

    return dates[tour_number]


def get_tour_time(tour_number):
    """Нечётные туры — 10:00, чётные — 16:00."""

    if tour_number % 2 == 1:
        return "10:00"

    return "16:00"


def get_tour_stage(tour_number):
    """Определяет этап соревнований."""

    if tour_number <= 5:
        return "Подгруппы"

    return "Финальные игры"


# ============================================================
# ГЕНЕРАЦИЯ ОДНОГО МАТЧА
# ============================================================

def generate_match(
    main,
    players,
    schedule,
    schedule_start,
    tour_number,
    output_dir,
    log
):
    """Создает заявку и протокол для одного матча."""

    # --------------------------------------------------------
    # Получение команд
    # --------------------------------------------------------

    team_1 = schedule[f"V{schedule_start}"].value
    team_2 = schedule[f"W{schedule_start}"].value

    if not team_1 or not team_2:
        raise ValueError(
            f"Не найдены команды в строке {schedule_start}"
        )

    # --------------------------------------------------------
    # Создание заявки
    # --------------------------------------------------------

    application = load_workbook(APPLICATION_TEMPLATE)

    app = application["app"]

    # Дата матча из расписания
    match_date = main[MAIN_SHEET][
        f"B{schedule_start}"
    ].value

    # Номер тура
    app["D5"] = f"{tour_number} тур"

    # Этап соревнований
    app["D6"] = get_tour_stage(tour_number)

    # Дата проведения тура
    app["O5"] = get_tour_date(tour_number)

    # Время проведения тура
    app["O6"] = get_tour_time(tour_number)

    # --------------------------------------------------------
    # Заполнение команд
    # --------------------------------------------------------

    app["T5"] = match_date
    app["AO5"] = match_date

    # Номер стола из ячейки T5 — используем его в названиях файлов.
    table_number = app["T5"].value
    if table_number is None or str(table_number).strip() == "":
        raise ValueError(
            f"Не указан номер стола в ячейке T5 для матча {team_1} — {team_2}"
        )
    table_prefix = f"Т{table_number}"

    app["C8"] = team_1
    app["X8"] = team_1

    app["M8"] = team_2
    app["AH8"] = team_2

    # --------------------------------------------------------
    # Игроки первой команды
    # --------------------------------------------------------

    team_1_players = (
        players[
            players["Название"] == team_1
        ]["Имя в заявку"]
        .dropna()
        .tolist()
    )

    application_start = 10

    for player in team_1_players:
        app[f"D{application_start}"] = player
        app[f"Y{application_start}"] = player

        application_start += 1

    # --------------------------------------------------------
    # Игроки второй команды
    # --------------------------------------------------------

    team_2_players = (
        players[
            players["Название"] == team_2
        ]["Имя в заявку"]
        .dropna()
        .tolist()
    )

    application_start = 10

    for player in team_2_players:
        app[f"N{application_start}"] = player
        app[f"AI{application_start}"] = player

        application_start += 1

    # --------------------------------------------------------
    # Названия команд для файлов
    # --------------------------------------------------------

    team_1_pseudonym = main[MAIN_SHEET][
        f"A{schedule_start}"
    ].value

    team_2_pseudonym = main[MAIN_SHEET][
        f"C{schedule_start}"
    ].value

    application_filename = (
        f"{table_prefix} - Заявка на матч "
        f"{team_1_pseudonym} - {team_2_pseudonym} "
        f"Суперлига 2026.xlsx"
    )

    application_path = os.path.join(
        output_dir,
        application_filename
    )

    application.save(application_path)

    # --------------------------------------------------------
    # Создание протокола
    # --------------------------------------------------------

    protocol_filename = (
        f"{table_prefix} - Протокол "
        f"{team_1_pseudonym} - {team_2_pseudonym} "
        f"Суперлига 2026.xlsx"
    )

    protocol_path = os.path.join(
        output_dir,
        protocol_filename
    )

    protocol = load_workbook(PROTOCOL_TEMPLATE)

    protocol.save(protocol_path)

    # --------------------------------------------------------
    # Заполнение формул протокола
    # --------------------------------------------------------

    protocol2 = load_workbook(protocol_path)

    prot2 = protocol2["prot"]

    source_file = application_filename

    # Верхняя часть протокола
    prot2["B2"] = f"='[{source_file}]app'!B4"

    prot2["D3"] = f"='[{source_file}]app'!D5"
    prot2["M3"] = f"='[{source_file}]app'!D6"

    prot2["S3"] = f"='[{source_file}]app'!O5"
    prot2["U3"] = f"='[{source_file}]app'!O6"

    prot2["W3"] = (
        f'="TABLE "&\'[{source_file}]app\'!T5'
    )

    prot2["D4"] = f"='[{source_file}]app'!E17"
    prot2["N4"] = f"='[{source_file}]app'!Z17"

    prot2["D5"] = f"='[{source_file}]app'!O17"
    prot2["N5"] = f"='[{source_file}]app'!AJ17"

    prot2["E7"] = f"='[{source_file}]app'!D25"
    prot2["N7"] = f"='[{source_file}]app'!Y25"

    # Игры
    prot2["C11"] = f"='[{source_file}]app'!F19"
    prot2["S11"] = f"='[{source_file}]app'!AA19"

    prot2["C18"] = f"='[{source_file}]app'!F20"
    prot2["S18"] = f"='[{source_file}]app'!AA20"

    prot2["D22"] = f"='[{source_file}]app'!F23"
    prot2["D23"] = f"='[{source_file}]app'!F24"

    prot2["Q22"] = f"='[{source_file}]app'!AA23"
    prot2["Q23"] = f"='[{source_file}]app'!AA24"

    # Формулы счета
    prot2["C26"] = (
        f'=IF(D25="+",'
        f'IF(J22="+",D22,'
        f'IF(J23="+",D23,\'[{source_file}]app\'!F21)),'
        f'\'[{source_file}]app\'!F21)'
    )

    prot2["S26"] = (
        f'=IF(V25="+",'
        f'IF(P22="+",Q22,'
        f'IF(P23="+",Q23,\'[{source_file}]app\'!AA21)),'
        f'\'[{source_file}]app\'!AA21)'
    )

    prot2["C33"] = (
        '=IF(D32="+",'
        'IF(J22="+",D22,'
        'IF(J23="+",D23,C11)),C11)'
    )

    prot2["S33"] = (
        '=IF(V32="+",'
        'IF(P22="+",Q22,'
        'IF(P23="+",Q23,S18)),S18)'
    )

    prot2["C40"] = (
        '=IF(D39="+",'
        'IF(J22="+",D22,'
        'IF(J23="+",D23,C18)),C18)'
    )

    prot2["S40"] = (
        '=IF(V39="+",'
        'IF(P22="+",Q22,'
        'IF(P23="+",Q23,S11)),S11)'
    )

    prot2["C41"] = (
        f'=IF(D39="++",'
        f'IF(J22="+",D22,'
        f'IF(J23="+",D23,'
        f'IF(\'[{source_file}]app\'!D22="",C26,'
        f'\'[{source_file}]app\'!F22))),'
        f'IF(\'[{source_file}]app\'!D22="",C26,'
        f'\'[{source_file}]app\'!F22))'
    )

    prot2["S41"] = (
        f'=IF(V39="++",'
        f'IF(P22="+",Q22,'
        f'IF(P23="+",Q23,'
        f'IF(\'[{source_file}]app\'!Y22="",S26,'
        f'\'[{source_file}]app\'!AA22))),'
        f'IF(\'[{source_file}]app\'!Y22="",S26,'
        f'\'[{source_file}]app\'!AA22))'
    )

    # Замена формул
    prot2["D22"] = f"='[{source_file}]app'!F23"
    prot2["D23"] = f"='[{source_file}]app'!F24"

    prot2["Q22"] = f"='[{source_file}]app'!AA23"
    prot2["Q23"] = f"='[{source_file}]app'!AA24"

    # Нижняя часть протокола
    prot2["C46"] = f"='[{source_file}]app'!E17"
    prot2["O46"] = f"='[{source_file}]app'!Z17"

    prot2["F54"] = f"='[{source_file}]app'!N20"
    prot2["F55"] = f"='[{source_file}]app'!N25"

    # --------------------------------------------------------
    # Сохранение протокола
    # --------------------------------------------------------

    protocol2.save(protocol_path)

    log(
        f"Созданы документы: "
        f"{team_1_pseudonym} — {team_2_pseudonym}"
    )


# ============================================================
# ГЕНЕРАЦИЯ ВСЕХ МАТЧЕЙ ТУРА
# ============================================================

def generate_documents(tour_number, log):
    """Генерирует документы для выбранного тура."""

    log("Загрузка основного Excel-файла...")

    main = load_workbook(
    MAIN_FILE,
    data_only=True
)

    

    log("Загрузка списка игроков...")

    players = pd.read_excel(PLAYERS_FILE)

    schedule = main[SCHEDULE_SHEET]

    tour_rows = get_tour_rows(tour_number)

    # Отдельная папка для каждого тура: outputs/R1, outputs/R2 и т.д.
    tour_output_dir = os.path.join(OUTPUT_DIR, f"R{tour_number}")
    os.makedirs(tour_output_dir, exist_ok=True)

    log(f"Папка сохранения: {tour_output_dir}")
    log(f"Начинается генерация {tour_number}-го тура.")

    total_matches = len(tour_rows)

    for index, schedule_start in enumerate(
        tour_rows,
        start=1
    ):

        log(
            f"Матч {index}/{total_matches}: "
            f"обработка строки {schedule_start}"
        )

        generate_match(
            main=main,
            players=players,
            schedule=schedule,
            schedule_start=schedule_start,
            tour_number=tour_number,
            output_dir=tour_output_dir,
            log=log
        )

    log(
        f"Генерация {tour_number}-го тура завершена."
    )


# ============================================================
# ГРАФИЧЕСКИЙ ИНТЕРФЕЙС
# ============================================================

class SuperLeagueApp:

    def __init__(self, root):

        self.root = root

        self.root.title(
            "Генерация протоколов и заявок"
        )

        self.root.geometry("650x450")
        self.root.minsize(600, 400)

        self.root.protocol(
            "WM_DELETE_WINDOW",
            self.on_close
        )

        self.is_generating = False

        self.setup_ui()

    # --------------------------------------------------------
    # Интерфейс
    # --------------------------------------------------------

    def setup_ui(self):

        main_frame = ttk.Frame(
            self.root,
            padding=15
        )

        main_frame.pack(
            fill="both",
            expand=True
        )

        # Выбор тура
        tour_frame = ttk.Frame(main_frame)

        tour_frame.pack(
            fill="x",
            pady=(0, 10)
        )

        tour_label = ttk.Label(
            tour_frame,
            text="Выберите тур:",
            font=("Arial", 11)
        )

        tour_label.pack(
            side="left",
            padx=(0, 10)
        )

        self.tour_combobox = ttk.Combobox(
            tour_frame,
            values=list(range(1, TOTAL_TOURS + 1)),
            state="readonly",
            width=10,
        )

        self.tour_combobox.current(0)

        self.tour_combobox.pack(
            side="left"
        )

        # Кнопка генерации
        self.generate_button = ttk.Button(
            main_frame,
            text="Сгенерировать документы",
            command=self.start_generation
        )

        self.generate_button.pack(
            fill="x",
            pady=(0, 10)
        )

        # Лог
        log_label = ttk.Label(
            main_frame,
            text="Журнал выполнения:",
            font=("Arial", 10)
        )

        log_label.pack(
            anchor="w",
            pady=(0, 5)
        )

        log_frame = ttk.Frame(main_frame)

        log_frame.pack(
            fill="both",
            expand=True
        )

        self.log_text = tk.Text(
            log_frame,
            height=15,
            wrap="word",
            state="disabled",
            font=("Consolas", 9)
        )

        scrollbar = ttk.Scrollbar(
            log_frame,
            orient="vertical",
            command=self.log_text.yview
        )

        self.log_text.configure(
            yscrollcommand=scrollbar.set
        )

        self.log_text.pack(
            side="left",
            fill="both",
            expand=True
        )

        scrollbar.pack(
            side="right",
            fill="y"
        )

        # Статус
        self.status_label = ttk.Label(
            main_frame,
            text="Готов к работе",
            anchor="w"
        )

        self.status_label.pack(
            fill="x",
            pady=(10, 0)
        )

    # --------------------------------------------------------
    # Вывод логов
    # --------------------------------------------------------

    def log(self, message):

        logger.info(message)

        def write_message():

            self.log_text.configure(
                state="normal"
            )

            self.log_text.insert(
                "end",
                message + "\n"
            )

            self.log_text.see("end")

            self.log_text.configure(
                state="disabled"
            )

        try:
            self.root.after(
                0,
                write_message
            )
        except RuntimeError:
            pass

    # --------------------------------------------------------
    # Запуск генерации
    # --------------------------------------------------------

    def start_generation(self):

        if self.is_generating:
            return

        selected_tour = self.tour_combobox.get()

        if not selected_tour:
            messagebox.showwarning(
                "Внимание",
                "Выберите тур."
            )

            return

        tour_number = int(selected_tour)

        self.is_generating = True

        self.generate_button.configure(
            state="disabled"
        )

        self.tour_combobox.configure(
            state="disabled"
        )

        self.status_label.configure(
            text=f"Генерация {tour_number}-го тура..."
        )

        self.log_text.configure(
            state="normal"
        )

        self.log_text.delete(
            "1.0",
            "end"
        )

        self.log_text.configure(
            state="disabled"
        )

        thread = threading.Thread(
            target=self.run_generation,
            args=(tour_number,),
            daemon=True
        )

        thread.start()

    # --------------------------------------------------------
    # Выполнение в отдельном потоке
    # --------------------------------------------------------

    def run_generation(self, tour_number):

        try:

            generate_documents(
                tour_number,
                self.log
            )

            self.root.after(
                0,
                self.generation_success
            )

        except Exception as error:

            error_traceback = traceback.format_exc()

            logger.error(
                "Ошибка генерации:\n%s",
                error_traceback
            )

            self.log(
                f"ОШИБКА: {error}"
            )

            self.root.after(
                0,
                self.generation_error,
                str(error),
                error_traceback
            )

    # --------------------------------------------------------
    # Успешное завершение
    # --------------------------------------------------------

    def generation_success(self):

        self.is_generating = False

        self.generate_button.configure(
            state="normal"
        )

        self.tour_combobox.configure(
            state="readonly"
        )

        self.status_label.configure(
            text="Генерация завершена успешно."
        )

        messagebox.showinfo(
            "Готово",
            "Документы успешно сгенерированы!\n\n"
            "Созданы 6 заявок и 6 протоколов.\n\n"
            f"Папка сохранения:\n{OUTPUT_DIR}"
        )

    # --------------------------------------------------------
    # Ошибка генерации
    # --------------------------------------------------------

    def generation_error(
        self,
        error_message,
        error_traceback
    ):

        self.is_generating = False

        self.generate_button.configure(
            state="normal"
        )

        self.tour_combobox.configure(
            state="readonly"
        )

        self.status_label.configure(
            text="Ошибка при генерации документов."
        )

        messagebox.showerror(
            "Ошибка",
            "Во время генерации произошла ошибка:\n\n"
            f"{error_message}\n\n"
            "Подробности записаны в logs/superleague.log"
        )

    # --------------------------------------------------------
    # Закрытие приложения
    # --------------------------------------------------------

    def on_close(self):

        if self.is_generating:

            result = messagebox.askyesno(
                "Генерация выполняется",
                "Документы еще генерируются.\n"
                "Вы действительно хотите закрыть программу?"
            )

            if not result:
                return

        self.root.destroy()


# ============================================================
# ЗАПУСК ПРОГРАММЫ
# ============================================================

def main():

    root = tk.Tk()

    app = SuperLeagueApp(root)

    root.mainloop()


if __name__ == "__main__":

    try:
        main()

    except Exception:

        error_traceback = traceback.format_exc()

        logger.critical(
            "Критическая ошибка запуска:\n%s",
            error_traceback
        )

        try:

            messagebox.showerror(
                "Критическая ошибка",
                "Не удалось запустить приложение.\n\n"
                "Подробности находятся в logs/superleague.log"
            )

        except Exception:
            pass