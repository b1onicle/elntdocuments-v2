
import os
from pathlib import Path
from collections import defaultdict
from copy import copy

import psycopg
import openpyxl
from dotenv import load_dotenv
from openpyxl.utils import get_column_letter


# =====================================================
# НАСТРОЙКИ
# =====================================================

load_dotenv()

BASE_DIR = Path(__file__).resolve().parent.parent
RESULT_FILE = BASE_DIR / "outputs" / "result.xlsx"

DB_CONFIG = {
    "host": os.getenv("DB_HOST", "localhost"),
    "port": os.getenv("DB_PORT", "5432"),
    "dbname": os.getenv("DB_NAME", "postgres"),
    "user": os.getenv("DB_USER", "postgres"),
    "password": os.getenv("DB_PASSWORD"),
}

LEAGUE_ID = 1

GROUPS = {
    1: [7, 9, 11, 13, 15, 17],
    2: [22, 24, 26, 28, 30, 32],
}


# =====================================================
# НОРМАЛИЗАЦИЯ И ПОИСК КОМАНД
# =====================================================

def normalize_name(name):
    if not name:
        return ""

    return " ".join(str(name).split()).casefold()


def find_team_id(cur, excel_name):
    """
    Ищет команду в PostgreSQL через ILIKE.

    Сначала проверяет полное совпадение без учёта регистра.
    Если его нет — ищет название как подстроку.

    Если найдено несколько вариантов, не выбирает случайный.
    """

    # 1. Полное совпадение без учёта регистра
    cur.execute("""
        SELECT id, name
        FROM teams
        WHERE name ILIKE %s
          AND league_id = %s
          AND deleted_at IS NULL
    """, (excel_name, LEAGUE_ID))

    results = cur.fetchall()

    if len(results) == 1:
        team_id, db_name = results[0]

        print(
            f"[НАЙДЕНА] {excel_name} -> "
            f"{db_name} (ID: {team_id})"
        )

        return team_id

    if len(results) > 1:
        print(
            f"[НЕОДНОЗНАЧНО] {excel_name}: "
            f"найдено {len(results)} совпадений"
        )

        for team_id, db_name in results:
            print(f"    ID {team_id}: {db_name}")

        return None

    # 2. Частичное совпадение через ILIKE
    cur.execute("""
        SELECT id, name
        FROM teams
        WHERE name ILIKE %s
          AND league_id = %s
          AND deleted_at IS NULL
    """, (f"%{excel_name}%", LEAGUE_ID))

    results = cur.fetchall()

    if len(results) == 1:
        team_id, db_name = results[0]

        print(
            f"[НАЙДЕНА ILIKE] {excel_name} -> "
            f"{db_name} (ID: {team_id})"
        )

        return team_id

    if len(results) == 0:
        print(
            f"[НЕ НАЙДЕНА] Команда: {excel_name}"
        )

    else:
        print(
            f"[НЕОДНОЗНАЧНО] {excel_name}: "
            f"найдено {len(results)} совпадений"
        )

        for team_id, db_name in results:
            print(f"    ID {team_id}: {db_name}")

    return None


# =====================================================
# ПОЛУЧЕНИЕ КОМАНД ИЗ EXCEL
# =====================================================

def get_team_positions(ws):
    """
    Получает названия команд из столбца C.

    Возвращает:
    название -> группа, строка, позиция.
    """

    positions = {}

    for group, rows in GROUPS.items():

        for index, row in enumerate(rows):

            team_name = ws.cell(row=row, column=3).value

            if not team_name:
                continue

            key = normalize_name(team_name)

            if key in positions:
                raise ValueError(
                    f"Команда повторяется в шаблоне: {team_name}"
                )

            positions[key] = {
                "name": team_name,
                "group": group,
                "row": row,
                "index": index,
            }

    return positions


# =====================================================
# ПОДГОТОВКА EXCEL
# =====================================================

def unmerge_block(ws, row, first_col):
    """
    Разъединяет блок 3x2, сохраняя оформление ячеек.
    """

    start = f"{get_column_letter(first_col)}{row}"
    end = f"{get_column_letter(first_col + 2)}{row + 1}"

    target_range = f"{start}:{end}"

    merged_ranges = {
        str(rng) for rng in ws.merged_cells.ranges
    }

    if target_range not in merged_ranges:
        return

    styles = {}

    for r in range(row, row + 2):
        for c in range(first_col, first_col + 3):
            styles[(r, c)] = copy(ws.cell(r, c)._style)

    ws.unmerge_cells(target_range)

    for (r, c), style in styles.items():
        ws.cell(r, c)._style = style


def prepare_excel(ws):
    """
    Подготавливает матрицу результатов.
    Удаляет предыдущие результаты.
    """

    for group, rows in GROUPS.items():

        for team_index, team_row in enumerate(rows):

            for opponent_index in range(6):

                if team_index == opponent_index:
                    continue

                first_col = 4 + opponent_index * 3

                unmerge_block(
                    ws,
                    team_row,
                    first_col
                )

                # Очищаем очки матча
                ws.cell(
                    row=team_row,
                    column=first_col + 1
                ).value = None

                # Очищаем счёт игр
                ws.cell(
                    row=team_row + 1,
                    column=first_col
                ).value = None


# =====================================================
# ЗАПИСЬ РЕЗУЛЬТАТОВ В EXCEL
# =====================================================

def write_result(
    ws,
    team_row,
    opponent_index,
    match_points,
    game_scores
):
    first_col = 4 + opponent_index * 3

    # Очки за матч
    ws.cell(
        row=team_row,
        column=first_col + 1
    ).value = match_points

    # Счёт отдельных игр
    ws.cell(
        row=team_row + 1,
        column=first_col
    ).value = ", ".join(game_scores)


# =====================================================
# ПОЛУЧЕНИЕ МАТЧЕЙ ИЗ POSTGRESQL
# =====================================================

def get_matches(cur, team_ids):

    if not team_ids:
        return []

    cur.execute("""
        SELECT
            m.id,
            m.team_1_id,
            t1.name,
            m.team_2_id,
            t2.name
        FROM matches m
        JOIN teams t1 ON t1.id = m.team_1_id
        JOIN teams t2 ON t2.id = m.team_2_id
        WHERE m.team_1_id = ANY(%s)
          AND m.team_2_id = ANY(%s)
          AND t1.league_id = %s
          AND t2.league_id = %s
          AND t1.deleted_at IS NULL
          AND t2.deleted_at IS NULL
        ORDER BY m.id
    """, (
        team_ids,
        team_ids,
        LEAGUE_ID,
        LEAGUE_ID
    ))

    return cur.fetchall()


# =====================================================
# РАСЧЁТ РЕЗУЛЬТАТОВ МАТЧА
# =====================================================

def get_match_result(
    cur,
    match_id,
    team_1_id,
    team_2_id
):

    cur.execute("""
        SELECT
            g.id,
            p1.team_id,
            p2.team_id,
            s.set_number,
            s.player_1_score,
            s.player_2_score
        FROM games g
        JOIN players p1 ON p1.id = g.player_1_id
        JOIN players p2 ON p2.id = g.player_2_id
        JOIN sets s ON s.game_id = g.id
        WHERE g.match_id = %s
        ORDER BY g.id, s.set_number
    """, (match_id,))

    rows = cur.fetchall()

    if not rows:
        return None

    games = defaultdict(list)

    for row in rows:

        game_id = row[0]
        player_1_team = row[1]
        player_2_team = row[2]
        set_number = row[3]
        score_1 = row[4]
        score_2 = row[5]

        if {
            player_1_team,
            player_2_team
        } != {
            team_1_id,
            team_2_id
        }:
            raise ValueError(
                f"Неверные команды игроков в игре {game_id}"
            )

        if score_1 == score_2:
            raise ValueError(
                f"Ничья в сете {set_number}, игра {game_id}"
            )

        games[game_id].append({
            "player_1_team": player_1_team,
            "player_2_team": player_2_team,
            "score_1": score_1,
            "score_2": score_2,
        })

    team_1_wins = 0
    team_2_wins = 0

    scores_team_1 = []
    scores_team_2 = []

    for game_id, sets in sorted(games.items()):

        player_1_sets = 0
        player_2_sets = 0

        for s in sets:

            if s["score_1"] > s["score_2"]:
                player_1_sets += 1

            else:
                player_2_sets += 1

        # Игра должна иметь определённого победителя
        if player_1_sets == player_2_sets:
            raise ValueError(
                f"Игра {game_id} не завершена: "
                f"{player_1_sets}:{player_2_sets}"
            )

        player_1_team = sets[0]["player_1_team"]

        if player_1_team == team_1_id:
            score_1 = player_1_sets
            score_2 = player_2_sets

        else:
            score_1 = player_2_sets
            score_2 = player_1_sets

        scores_team_1.append(f"{score_1}:{score_2}")
        scores_team_2.append(f"{score_2}:{score_1}")

        if score_1 > score_2:
            team_1_wins += 1
        else:
            team_2_wins += 1

    if team_1_wins > team_2_wins:
        points_1 = 2
        points_2 = 1

    elif team_2_wins > team_1_wins:
        points_1 = 1
        points_2 = 2

    else:
        points_1 = 1
        points_2 = 1

    return {
        "team_1_wins": team_1_wins,
        "team_2_wins": team_2_wins,
        "points_1": points_1,
        "points_2": points_2,
        "scores_1": scores_team_1,
        "scores_2": scores_team_2,
    }


# =====================================================
# ОСНОВНАЯ ПРОГРАММА
# =====================================================

def main():

    if not DB_CONFIG["password"]:
        raise ValueError(
            "Не задан DB_PASSWORD в файле .env"
        )

    if not RESULT_FILE.exists():
        raise FileNotFoundError(
            f"Файл не найден: {RESULT_FILE}"
        )

    print("=" * 55)
    print("ОБНОВЛЕНИЕ ИТОГОВОЙ ТАБЛИЦЫ")
    print("=" * 55)

    wb = openpyxl.load_workbook(RESULT_FILE)
    ws = wb["res"]

    team_positions = get_team_positions(ws)

    print(f"\nКоманд в Excel: {len(team_positions)}")

    processed = 0
    skipped = 0

    try:

        with psycopg.connect(**DB_CONFIG) as conn:

            with conn.cursor() as cur:

                # -------------------------------------
                # Сопоставляем Excel-команды с БД
                # -------------------------------------

                print("\nПОИСК КОМАНД")
                print("-" * 55)

                excel_to_db = {}
                db_to_excel = {}

                for key, position in team_positions.items():

                    team_name = position["name"]

                    team_id = find_team_id(
                        cur,
                        team_name
                    )

                    if team_id is None:
                        skipped += 1
                        continue

                    if team_id in db_to_excel:
                        raise ValueError(
                            f"Две команды Excel соответствуют "
                            f"одной команде PostgreSQL: "
                            f"{db_to_excel[team_id]} и {team_name}"
                        )

                    excel_to_db[key] = team_id
                    db_to_excel[team_id] = position

                team_ids = list(db_to_excel.keys())

                print(
                    f"\nУспешно сопоставлено: "
                    f"{len(excel_to_db)} из {len(team_positions)}"
                )

                if not team_ids:
                    raise ValueError(
                        "Не удалось сопоставить ни одной команды"
                    )

                # -------------------------------------
                # Получаем матчи
                # -------------------------------------

                matches = get_matches(
                    cur,
                    team_ids
                )

                print(f"\nМатчей в базе: {len(matches)}")

                # -------------------------------------
                # Подготавливаем Excel
                # -------------------------------------

                prepare_excel(ws)

                # -------------------------------------
                # Обрабатываем матчи
                # -------------------------------------

                print("\nОБРАБОТКА МАТЧЕЙ")
                print("-" * 55)

                processed_pairs = set()

                for match in matches:

                    (
                        match_id,
                        team_1_id,
                        team_1_name,
                        team_2_id,
                        team_2_name
                    ) = match

                    pos_1 = db_to_excel.get(team_1_id)
                    pos_2 = db_to_excel.get(team_2_id)

                    if pos_1 is None or pos_2 is None:
                        skipped += 1
                        continue

                    if pos_1["group"] != pos_2["group"]:
                        print(
                            f"[ПРОПУСК {match_id}] "
                            f"Команды из разных групп"
                        )
                        skipped += 1
                        continue

                    pair = frozenset((team_1_id, team_2_id))

                    if pair in processed_pairs:
                        raise ValueError(
                            f"Несколько матчей между "
                            f"{pos_1['name']} и {pos_2['name']}. "
                            f"Нельзя однозначно определить, "
                            f"какой результат записывать."
                        )

                    result = get_match_result(
                        cur,
                        match_id,
                        team_1_id,
                        team_2_id
                    )

                    if result is None:
                        print(
                            f"[ПРОПУСК {match_id}] "
                            f"Нет результатов сетов"
                        )
                        skipped += 1
                        continue

                    processed_pairs.add(pair)

                    # Результат первой команды
                    write_result(
                        ws,
                        pos_1["row"],
                        pos_2["index"],
                        result["points_1"],
                        result["scores_1"]
                    )

                    # Результат второй команды
                    write_result(
                        ws,
                        pos_2["row"],
                        pos_1["index"],
                        result["points_2"],
                        result["scores_2"]
                    )

                    processed += 1

                    print(
                        f"[OK {match_id}] "
                        f"{pos_1['name']} "
                        f"{result['team_1_wins']}:"
                        f"{result['team_2_wins']} "
                        f"{pos_2['name']}"
                    )

        # ---------------------------------------------
        # Сохранение
        # ---------------------------------------------

        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"

        wb.save(RESULT_FILE)

    finally:
        wb.close()

    print("\n" + "=" * 55)
    print("ОБРАБОТКА ЗАВЕРШЕНА")
    print("=" * 55)

    print(f"Обработано матчей: {processed}")
    print(f"Пропущено: {skipped}")
    print(f"Обновлён файл: {RESULT_FILE}")


if __name__ == "__main__":
    main()
