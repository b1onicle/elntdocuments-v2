
from pathlib import Path
import os
import re
import sys
import tempfile

import psycopg
from openpyxl import load_workbook


# ============================================================
# НАСТРОЙКИ
# ============================================================

DB_CONFIG = {
    "host": "localhost",
    "port": 5432,
    "dbname": "postgres",
    "user": "postgres",
    "password": "123",
}

# Структура:
# project/
# ├── outputs/
# │   ├── R1/
# │   └── R2/
# └── scripts/
#     └── import_protocols.py

SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_DIR = SCRIPT_DIR.parent
OUTPUT_DIR = PROJECT_DIR / "outputs"

SUPERLEAGUE_ID = 1

PLAYER_ROWS = [11, 18, 26, 33]
SET_START_ROWS = [10, 17, 25, 32]

APPLICATION_SHEET = "app"
PROTOCOL_SHEET = "prot"

APPLICATION_MATCH_ID_CELL = "O4"
PROTOCOL_MATCH_ID_CELL = "U2"

TEAM_1_CELL = "C8"
TEAM_2_CELL = "M8"


# ============================================================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ============================================================

def parse_optional_id(value, source):
    if value is None or str(value).strip() == "":
        return None

    try:
        return int(value)
    except (TypeError, ValueError):
        raise ValueError(
            f"Некорректный ID в {source}: {value!r}"
        )


def parse_score(value, cell_address):
    if value is None or str(value).strip() == "":
        return None

    try:
        return int(value)
    except (TypeError, ValueError):
        raise ValueError(
            f"Некорректный счёт в {cell_address}: {value!r}"
        )


# ============================================================
# ПОИСК ПРОТОКОЛОВ И ЗАЯВОК
# ============================================================

def find_protocol_files():

    if not OUTPUT_DIR.exists():
        raise FileNotFoundError(
            f"Папка outputs не найдена: {OUTPUT_DIR}"
        )

    protocol_files = []

    for file_path in OUTPUT_DIR.rglob("*.xlsx"):

        if file_path.name.startswith("~$"):
            continue

        if "протокол" in file_path.stem.casefold():
            protocol_files.append(file_path)

    return sorted(protocol_files)


def find_application_for_protocol(protocol_path):

    application_stem = re.sub(
        r"протокол",
        "Заявка на матч",
        protocol_path.stem,
        count=1,
        flags=re.IGNORECASE,
    )

    application_path = protocol_path.with_name(
        application_stem + protocol_path.suffix
    )

    if not application_path.exists():
        raise FileNotFoundError(
            "Не найдена соответствующая заявка.\n"
            f"Протокол: {protocol_path.name}\n"
            f"Ожидалась заявка: {application_path.name}"
        )

    return application_path


# ============================================================
# ПОИСК КОМАНД И ИГРОКОВ
# ============================================================

def get_team_id(cursor, team_name):

    if not isinstance(team_name, str) or not team_name.strip():
        raise ValueError(
            f"Пустое название команды: {team_name!r}"
        )

    cursor.execute(
        """
        SELECT id, name
        FROM teams
        WHERE LOWER(TRIM(name)) = LOWER(TRIM(%s))
          AND league_id = %s
          AND deleted_at IS NULL
        """,
        (team_name.strip(), SUPERLEAGUE_ID),
    )

    results = cursor.fetchall()

    if not results:
        raise ValueError(
            f"Команда не найдена в Суперлиге: {team_name}"
        )

    if len(results) > 1:
        raise ValueError(
            f"Несколько команд с названием {team_name}: {results}"
        )

    return results[0][0]


def get_player_id(cursor, player_name):

    if not isinstance(player_name, str) or not player_name.strip():
        raise ValueError(
            f"Некорректное имя игрока: {player_name!r}"
        )

    parts = player_name.strip().split()

    if len(parts) != 2:
        raise ValueError(
            "Ожидается формат «Фамилия Имя»: "
            f"{player_name!r}"
        )

    last_name, first_name = parts

    cursor.execute(
        """
        SELECT p.id, pe.last_name, pe.first_name, pe.middle_name
        FROM players p
        JOIN people pe ON pe.id = p.people_id
        WHERE LOWER(TRIM(pe.last_name)) = LOWER(%s)
          AND LOWER(TRIM(pe.first_name)) = LOWER(%s)
          AND p.deleted_at IS NULL
        """,
        (last_name, first_name),
    )

    results = cursor.fetchall()

    if not results:
        raise ValueError(
            f"Игрок не найден: {player_name}"
        )

    if len(results) > 1:
        raise ValueError(
            f"Несколько совпадений для {player_name}: {results}"
        )

    return results[0][0]


# ============================================================
# СОХРАНЕНИЕ ID В EXCEL
# ============================================================

def save_workbooks_with_id(
    application_wb,
    protocol_wb,
    application_path,
    protocol_path,
    match_id,
):

    app_temp = None
    protocol_temp = None

    try:

        application_wb["app"]["O4"] = match_id
        protocol_wb["prot"]["U2"] = match_id

        with tempfile.NamedTemporaryFile(
            suffix=application_path.suffix,
            prefix=application_path.stem + "_tmp_",
            dir=application_path.parent,
            delete=False,
        ) as temp_file:
            app_temp = Path(temp_file.name)

        with tempfile.NamedTemporaryFile(
            suffix=protocol_path.suffix,
            prefix=protocol_path.stem + "_tmp_",
            dir=protocol_path.parent,
            delete=False,
        ) as temp_file:
            protocol_temp = Path(temp_file.name)

        application_wb.save(app_temp)
        protocol_wb.save(protocol_temp)

        os.replace(app_temp, application_path)
        app_temp = None

        os.replace(protocol_temp, protocol_path)
        protocol_temp = None

    finally:

        for temp_path in (app_temp, protocol_temp):
            if temp_path and temp_path.exists():
                temp_path.unlink(missing_ok=True)


# ============================================================
# ИМПОРТ ОДНОГО МАТЧА
# ============================================================

def import_one_match(conn, protocol_path):

    application_path = find_application_for_protocol(
        protocol_path
    )

    application_wb = load_workbook(
        application_path,
        data_only=False
    )

    protocol_wb = load_workbook(
        protocol_path,
        data_only=False
    )

    protocol_values_wb = load_workbook(
        protocol_path,
        data_only=True,
        read_only=True
    )

    try:

        if "app" not in application_wb.sheetnames:
            raise ValueError("В заявке отсутствует лист app")

        if "prot" not in protocol_wb.sheetnames:
            raise ValueError("В протоколе отсутствует лист prot")

        app = application_wb["app"]
        prot = protocol_wb["prot"]
        prot_values = protocol_values_wb["prot"]

        team_1_name = app["C8"].value
        team_2_name = app["M8"].value

        if not team_1_name or not team_2_name:
            raise ValueError(
                "Не заполнены команды в app!C8 или app!M8"
            )

        app_match_id = parse_optional_id(
            app["O4"].value,
            f"{application_path.name}!O4"
        )

        protocol_match_id = parse_optional_id(
            prot["U2"].value,
            f"{protocol_path.name}!U2"
        )

        if (
            app_match_id is not None
            and protocol_match_id is not None
            and app_match_id != protocol_match_id
        ):
            raise ValueError(
                "ID в заявке и протоколе различаются: "
                f"{app_match_id} и {protocol_match_id}"
            )

        stored_match_id = (
            app_match_id
            if app_match_id is not None
            else protocol_match_id
        )

        with conn.transaction():

            with conn.cursor() as cursor:

                team_1_id = get_team_id(
                    cursor,
                    team_1_name
                )

                team_2_id = get_team_id(
                    cursor,
                    team_2_name
                )

                match_id = None

                # Проверяем ранее записанный ID.
                if stored_match_id is not None:

                    cursor.execute(
                        """
                        SELECT id, team_1_id, team_2_id
                        FROM matches
                        WHERE id = %s
                        """,
                        (stored_match_id,)
                    )

                    existing_match = cursor.fetchone()

                    if existing_match:

                        _, existing_team_1, existing_team_2 = (
                            existing_match
                        )

                        if (
                            existing_team_1 != team_1_id
                            or existing_team_2 != team_2_id
                        ):
                            raise ValueError(
                                "Команды существующего матча "
                                "не совпадают с заявкой."
                            )

                        match_id = stored_match_id

                # Создаём новый матч, если его ещё нет.
                if match_id is None:

                    cursor.execute(
                        """
                        INSERT INTO matches (
                            team_1_id,
                            team_2_id
                        )
                        VALUES (%s, %s)
                        RETURNING id
                        """,
                        (team_1_id, team_2_id)
                    )

                    match_id = cursor.fetchone()[0]

                    print(
                        f"Создан матч. ID: {match_id}"
                    )

                # Проверяем, не записаны ли игры ранее.
                cursor.execute(
                    """
                    SELECT EXISTS (
                        SELECT 1
                        FROM games
                        WHERE match_id = %s
                    )
                    """,
                    (match_id,)
                )

                games_exist = cursor.fetchone()[0]

                if games_exist:

                    protocol_values_wb.close()      
                    
                    save_workbooks_with_id(
                        application_wb,
                        protocol_wb,
                        application_path,
                        protocol_path,
                        match_id
                    )

                    return {
                        "status": "skipped",
                        "match_id": match_id,
                        "games": 0,
                        "sets": 0,
                    }

                total_games = 0
                total_sets = 0

                # Четыре игры матча.
                for game_index, player_row in enumerate(PLAYER_ROWS):

                    player_1_name = prot_values[
                        f"C{player_row}"
                    ].value

                    player_2_name = prot_values[
                        f"S{player_row}"
                    ].value

                    if not player_1_name and not player_2_name:
                        continue

                    if not player_1_name or not player_2_name:
                        raise ValueError(
                            f"Не заполнены оба игрока в строке "
                            f"{player_row}"
                        )

                    player_1_id = get_player_id(
                        cursor,
                        player_1_name
                    )

                    player_2_id = get_player_id(
                        cursor,
                        player_2_name
                    )

                    cursor.execute(
                        """
                        INSERT INTO games (
                            match_id,
                            player_1_id,
                            player_2_id
                        )
                        VALUES (%s, %s, %s)
                        RETURNING id
                        """,
                        (
                            match_id,
                            player_1_id,
                            player_2_id
                        )
                    )

                    game_id = cursor.fetchone()[0]
                    total_games += 1

                    set_start_row = SET_START_ROWS[game_index]

                    # Максимум пять сетов.
                    for set_index in range(5):

                        score_row = set_start_row + set_index

                        player_1_score = parse_score(
                            prot_values[f"L{score_row}"].value,
                            f"L{score_row}"
                        )

                        player_2_score = parse_score(
                            prot_values[f"N{score_row}"].value,
                            f"N{score_row}"
                        )

                        if (
                            player_1_score is None
                            and player_2_score is None
                        ):
                            break

                        if (
                            player_1_score is None
                            or player_2_score is None
                        ):
                            raise ValueError(
                                f"Неполный счёт сета "
                                f"в строке {score_row}"
                            )

                        cursor.execute(
                            """
                            INSERT INTO sets (
                                game_id,
                                player_1_score,
                                player_2_score,
                                set_number
                            )
                            VALUES (%s, %s, %s, %s)
                            """,
                            (
                                game_id,
                                player_1_score,
                                player_2_score,
                                set_index + 1
                            )
                        )

                        total_sets += 1

                if total_games == 0:
                    raise ValueError(
                        "В протоколе не найдено заполненных игр"
                    )

                # Записываем ID в Excel.
                protocol_values_wb.close()

                save_workbooks_with_id(
                    application_wb,
                    protocol_wb,
                    application_path,
                    protocol_path,
                    match_id
                )

                return {
                    "status": "imported",
                    "match_id": match_id,
                    "games": total_games,
                    "sets": total_sets,
                }

    finally:

        application_wb.close()
        protocol_wb.close()
        protocol_values_wb.close()


# ============================================================
# ГЛАВНАЯ ФУНКЦИЯ
# ============================================================

def main():

    protocol_files = find_protocol_files()

    if not protocol_files:
        print(
            f"Протоколы не найдены: {OUTPUT_DIR}"
        )
        return

    print(f"Папка поиска: {OUTPUT_DIR}")
    print(f"Найдено протоколов: {len(protocol_files)}")
    print("=" * 70)

    imported = 0
    skipped = 0
    failed = 0

    total_games = 0
    total_sets = 0

    with psycopg.connect(**DB_CONFIG) as conn:

        for index, protocol_path in enumerate(
            protocol_files,
            start=1
        ):

            print(
                f"\n[{index}/{len(protocol_files)}] "
                f"{protocol_path.relative_to(OUTPUT_DIR)}"
            )

            try:

                result = import_one_match(
                    conn,
                    protocol_path
                )

                if result["status"] == "imported":

                    imported += 1
                    total_games += result["games"]
                    total_sets += result["sets"]

                    print(
                        f"ИМПОРТИРОВАНО: "
                        f"match_id={result['match_id']}, "
                        f"игр={result['games']}, "
                        f"сетов={result['sets']}"
                    )

                else:

                    skipped += 1

                    print(
                        f"ПРОПУСК: матч ID {result['match_id']}. "
                        f"Игры уже есть в БД."
                    )

            except Exception as error:

                failed += 1

                print(f"ОШИБКА: {error}")
                print(
                    "Текущий матч не был импортирован."
                )

    print("\n" + "=" * 70)
    print("ИТОГ ИМПОРТА")
    print(f"Найдено протоколов: {len(protocol_files)}")
    print(f"Импортировано матчей: {imported}")
    print(f"Пропущено: {skipped}")
    print(f"Ошибок: {failed}")
    print(f"Добавлено игр: {total_games}")
    print(f"Добавлено сетов: {total_sets}")
    print("=" * 70)


if __name__ == "__main__":

    try:
        main()

    except KeyboardInterrupt:
        print("\nИмпорт прерван пользователем.")
        sys.exit(1)

    except Exception as error:
        print(f"\nКритическая ошибка: {error}")
        sys.exit(1)
