import os
import sqlite3
import logging
import re
from datetime import datetime, timedelta
from typing import Optional, List, Tuple

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.chart import LineChart, Reference

from telegram import Update, ReplyKeyboardRemove, InlineKeyboardMarkup, InlineKeyboardButton
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    CallbackQueryHandler,
    ContextTypes,
    ConversationHandler,
    MessageHandler,
    filters,
)

logging.basicConfig(format="%(asctime)s - %(levelname)s - %(message)s", level=logging.INFO)
logger = logging.getLogger(__name__)

DB_PATH = os.getenv("DB_PATH", "training_bot.db")
BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN","8653934626:AAHc_hYdOEBuskBwyyPa-uctWuTl1cwRHuE")

SESSION_DATE, CUSTOM_EXERCISE_NAME, SESSION_QUICK_INPUT, SESSION_MANUAL_INPUT, SESSION_RPE, SESSION_NOTES, SESSION_EDIT = range(7)

PROGRAMS = {
    "A": {
        "title": "ТРЕНИРОВКА A (СИЛА)",
        "notes": ["Отдых: 2–3 мин", "Интенсивность: RIR 1–2"],
        "review_weeks": 10,
        "exercises": [
            {"group": "Ноги", "exercise": "Присед со штангой", "sets": 4, "reps": "5–6", "rpe": "RIR 1–2", "step": "+2.5–5 кг"},
            {"group": "Грудь", "exercise": "Жим штанги лёжа", "sets": 4, "reps": "5–6", "rpe": "RIR 1–2", "step": "+2.5 кг"},
            {"group": "Спина", "exercise": "Подтягивания с весом", "sets": 4, "reps": "5–6", "rpe": "RIR 1–2", "step": "+1.25–2.5 кг"},
            {"group": "Спина", "exercise": "Тяга штанги в наклоне", "sets": 3, "reps": "6–8", "rpe": "RIR 1–2", "step": "+2.5 кг"},
            {"group": "Грудь", "exercise": "Брусья с весом (грудь)", "sets": 4, "reps": "5–6", "rpe": "RIR 1–2", "step": "+1.25–2.5 кг"},
            {"group": "Плечи", "exercise": "Жим гантелей сидя", "sets": 4, "reps": "6–8", "rpe": "RIR 1–2", "step": "+1–2 кг"},
            {"group": "Плечи", "exercise": "Face Pull", "sets": 3, "reps": "12–15", "rpe": "контроль", "step": "след. стек"},
            {"group": "Пресс", "exercise": "Подъём ног в висе", "sets": 3, "reps": "10–12", "rpe": "контроль", "step": "+1–2 повт"},
            {"group": "Пресс", "exercise": "Cable Crunch (тяжёлый)", "sets": 3, "reps": "10–12", "rpe": "контроль", "step": "след. стек"},
        ],
    },
    "B": {
        "title": "ТРЕНИРОВКА B (ГИПЕРТРОФИЯ)",
        "notes": ["Отдых: 60–90 сек"],
        "review_weeks": 10,
        "exercises": [
            {"group": "Грудь", "exercise": "Жим гантелей на наклонной скамье", "sets": 3, "reps": "8–10", "rpe": "7–9", "step": "+1–2 кг"},
            {"group": "Спина", "exercise": "Тяга горизонтального блока", "sets": 3, "reps": "10–12", "rpe": "7–9", "step": "след. стек"},
            {"group": "Грудь/Трицепс", "exercise": "Брусья (без веса)", "sets": 2, "reps": "10–12", "rpe": "без отказа", "step": "+1–2 повт"},
            {"group": "Ноги", "exercise": "Жим ногами", "sets": 3, "reps": "10–12", "rpe": "7–9", "step": "+5–10 кг"},
            {"group": "Ноги", "exercise": "Болгарские выпады", "sets": 3, "reps": "10", "rpe": "7–9", "step": "+1–2 кг"},
            {"group": "Плечи", "exercise": "Разведения гантелей в стороны", "sets": 4, "reps": "12–15", "rpe": "контроль", "step": "+1 кг"},
            {"group": "Плечи", "exercise": "Cable lateral raise", "sets": 3, "reps": "12–15", "rpe": "контроль", "step": "след. стек"},
            {"group": "Плечи", "exercise": "Face Pull / Reverse Pec Deck", "sets": 2, "reps": "12–15", "rpe": "контроль", "step": "след. стек"},
            {"group": "Бицепс", "exercise": "Бицепс (наклонная)", "sets": 3, "reps": "10–12", "rpe": "7–9", "step": "+1 кг"},
            {"group": "Трицепс", "exercise": "Трицепс (канат)", "sets": 2, "reps": "12–15", "rpe": "контроль", "step": "след. стек"},
            {"group": "Пресс", "exercise": "Скручивания с весом", "sets": 3, "reps": "12–15", "rpe": "контроль", "step": "след. стек"},
            {"group": "Кор", "exercise": "Pallof Press", "sets": 3, "reps": "12", "rpe": "контроль", "step": "+1–2 повт"},
        ],
    },
    "C": {
        "title": "ТРЕНИРОВКА C (ATHLETIC / METABOLIC)",
        "notes": ["Отдых: 30–60 сек"],
        "review_weeks": 10,
        "exercises": [
            {"group": "Грудь", "exercise": "Chest Press", "sets": 3, "reps": "15", "rpe": "7–8", "step": "след. стек"},
            {"group": "Спина", "exercise": "Тяга верхнего блока", "sets": 3, "reps": "15", "rpe": "7–8", "step": "след. стек"},
            {"group": "Трицепс", "exercise": "Брусья (трицепс)", "sets": 3, "reps": "15–20", "rpe": "7–8", "step": "+1–2 повт"},
            {"group": "Ноги", "exercise": "Выпады", "sets": 3, "reps": "15", "rpe": "7–8", "step": "+1–2 кг"},
            {"group": "Ягодицы", "exercise": "Ягодичный мост", "sets": 3, "reps": "12–15", "rpe": "7–8", "step": "+2.5–5 кг"},
            {"group": "Ноги", "exercise": "Сгибание ног лёжа", "sets": 2, "reps": "12–15", "rpe": "7–8", "step": "след. стек"},
            {"group": "Плечи", "exercise": "Mechanical Drop Set (разведения)", "sets": 3, "reps": "12 строгих / 10 читинг / 15 частичных", "rpe": "контроль", "step": "качество"},
            {"group": "Пресс", "exercise": "Подъём коленей в висе (медленно)", "sets": 3, "reps": "15", "rpe": "контроль", "step": "+1–2 повт"},
        ],
    },
}

EXERCISE_CATALOG = sorted(set(
    [ex["exercise"] for d in PROGRAMS.values() for ex in d["exercises"]] +
    ["Подтягивания", "Тяга штанги в наклоне", "Face Pull", "Reverse Pec Deck",
     "Жим гантелей на наклонной", "Тяга горизонтального блока", "Жим ногами",
     "Болгарские выпады", "Разведения гантелей в стороны", "Cable lateral raise",
     "Скручивания", "Планка", "Внешняя ротация резиной", "Y-Raise",
     "Подъём ног в висе", "Cable Crunch", "Подтягивания с весом"]
))
PAGE_SIZE = 8

def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def column_exists(cur, table_name, column_name):
    cur.execute(f"PRAGMA table_info({table_name})")
    return column_name in [row[1] for row in cur.fetchall()]

def init_db():
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS workouts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            workout_date TEXT NOT NULL,
            day_type TEXT NOT NULL,
            exercise TEXT NOT NULL,
            set1_reps REAL, set1_kg REAL,
            set2_reps REAL, set2_kg REAL,
            set3_reps REAL, set3_kg REAL,
            set4_reps REAL, set4_kg REAL,
            set5_reps REAL, set5_kg REAL,
            target_sets INTEGER,
            target_reps TEXT,
            target_guidance TEXT,
            step_rule TEXT,
            rpe REAL,
            notes TEXT,
            suggestion TEXT,
            coach_status TEXT,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        )
    """)
    extra_cols = {
        "target_sets": "INTEGER",
        "target_reps": "TEXT",
        "target_guidance": "TEXT",
        "step_rule": "TEXT",
        "suggestion": "TEXT",
        "coach_status": "TEXT",
    }
    for col, typ in extra_cols.items():
        if not column_exists(cur, "workouts", col):
            cur.execute(f"ALTER TABLE workouts ADD COLUMN {col} {typ}")
    conn.commit()
    conn.close()

def parse_date(s):
    for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(s.strip(), fmt).date().isoformat()
        except ValueError:
            pass
    raise ValueError("Используй дату в формате ДД.ММ.ГГГГ")

def parse_optional_float(s):
    s = s.strip().replace(",", ".")
    if s in ("", "-", "none", "null"):
        return None
    return float(s)

def parse_sets(text):
    parts = [p.strip() for p in text.split(",") if p.strip()]
    if not 1 <= len(parts) <= 5:
        raise ValueError("Нужно указать от 1 до 5 подходов.")
    result = []
    for part in parts:
        normalized = part.strip().replace("х", "x").replace("Х", "x").replace("×", "x").replace("*", "x")
        normalized = re.sub(r"\s+", "", normalized)
        if "x" not in normalized:
            raise ValueError("Формат: 8x60, 8х60, 8 * 60 или 8 x 60")
        reps_str, kg_str = normalized.split("x", 1)
        if not reps_str or not kg_str:
            raise ValueError("Формат: 8x60, 8х60, 8 * 60 или 8 x 60")
        result.append((parse_optional_float(reps_str), parse_optional_float(kg_str)))
    while len(result) < 5:
        result.append((None, None))
    return result

def build_repeated_sets(weight, reps, n_sets):
    result = [(reps, weight) for _ in range(max(1, min(5, n_sets)))]
    while len(result) < 5:
        result.append((None, None))
    return result

def calc_top_weight_and_reps(row):
    pairs = [(row["set1_reps"], row["set1_kg"]), (row["set2_reps"], row["set2_kg"]), (row["set3_reps"], row["set3_kg"]), (row["set4_reps"], row["set4_kg"]), (row["set5_reps"], row["set5_kg"])]
    valid = [(r, w) for r, w in pairs if r is not None and w is not None]
    if not valid:
        return None, None
    top_weight = max(w for r, w in valid)
    reps_at_top = max(r for r, w in valid if w == top_weight)
    return top_weight, reps_at_top

def calc_e1rm(top_weight, reps_at_top):
    if top_weight is None or reps_at_top is None:
        return None
    return round(float(top_weight) * (1 + float(reps_at_top) / 30), 1)

def calc_volume(row):
    total = 0.0
    for reps_col, kg_col in [("set1_reps","set1_kg"),("set2_reps","set2_kg"),("set3_reps","set3_kg"),("set4_reps","set4_kg"),("set5_reps","set5_kg")]:
        reps = row[reps_col]
        kg = row[kg_col]
        if reps is not None and kg is not None:
            total += float(reps) * float(kg)
    return round(total, 1)

def get_prev_best_e1rm(user_id, exercise, current_id=None):
    conn = get_conn()
    cur = conn.cursor()
    if current_id is None:
        cur.execute("SELECT * FROM workouts WHERE user_id=? AND exercise=? ORDER BY workout_date, id", (user_id, exercise))
    else:
        cur.execute("SELECT * FROM workouts WHERE user_id=? AND exercise=? AND id < ? ORDER BY workout_date, id", (user_id, exercise, current_id))
    rows = cur.fetchall()
    conn.close()
    vals = []
    for r in rows:
        tw, rr = calc_top_weight_and_reps(r)
        e1 = calc_e1rm(tw, rr)
        if e1 is not None:
            vals.append(e1)
    return max(vals) if vals else None

def get_last_same_exercise(user_id, exercise):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT * FROM workouts WHERE user_id=? AND exercise=? ORDER BY workout_date DESC, id DESC LIMIT 1", (user_id, exercise))
    row = cur.fetchone()
    conn.close()
    return row

def get_last_workout_date(user_id, day_type):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT workout_date FROM workouts WHERE user_id=? AND day_type=? ORDER BY workout_date DESC, id DESC LIMIT 1", (user_id, day_type))
    row = cur.fetchone()
    conn.close()
    return row["workout_date"] if row else None

def get_session_rows(user_id, day_type, workout_date):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT * FROM workouts WHERE user_id=? AND day_type=? AND workout_date=? ORDER BY id", (user_id, day_type, workout_date))
    rows = cur.fetchall()
    conn.close()
    return rows

def parse_rep_range(rep_text):
    txt = (rep_text or "-").replace(" ", "").replace("–", "-")
    if any(x in txt for x in ["сек", "sec", "/", "строгих", "читинг", "частичных"]) or txt == "-":
        return None, None
    if "-" in txt:
        a, b = txt.split("-", 1)
        try:
            return int(a), int(b)
        except ValueError:
            return None, None
    try:
        return int(txt), int(txt)
    except ValueError:
        return None, None

def find_exercise_config(day, exercise_name):
    for ex in PROGRAMS.get(day, {}).get("exercises", []):
        if ex["exercise"] == exercise_name:
            return ex
    return {"group": "Каталог", "exercise": exercise_name, "sets": None, "reps": "-", "rpe": "-", "step": "ручное решение"}

def format_last_performance(row):
    if not row:
        return "Нет прошлой записи."
    sets = []
    for i in range(1, 6):
        reps = row[f"set{i}_reps"]
        kg = row[f"set{i}_kg"]
        if reps is not None and kg is not None:
            sets.append(f"{reps:g}x{kg:g}")
    tw, rr = calc_top_weight_and_reps(row)
    e1 = calc_e1rm(tw, rr)
    return f"Прошлый раз: {', '.join(sets) if sets else '-'} | e1RM: {e1 if e1 else '-'}"

def analyze_progress(cfg, sets, rpe):
    low, high = parse_rep_range(cfg.get("reps"))
    valid = [(r, w) for r, w in sets if r is not None and w is not None]
    if not valid:
        return "Нет данных для подсказки.", "unknown"
    weights = [w for _, w in valid]
    top_weight = max(weights)
    reps_at_top = [r for r, w in valid if w == top_weight]
    min_reps_at_top = min(reps_at_top) if reps_at_top else None
    if low is None or high is None:
        if rpe is not None and rpe >= 9:
            return "Перегруз. Оставь или немного снизь вес.", "overload"
        if rpe is not None and rpe <= 7:
            return f"Легко. Можно ускорить прогрессию по правилу: {cfg['step']}.", "easy"
        return f"Свободное упражнение. Следующая цель: {cfg['step']}.", "stable"
    if min_reps_at_top is not None and min_reps_at_top >= high and (rpe is None or rpe <= 8.5):
        return f"Ты в верхней границе → пора добавлять вес. Правило: {cfg['step']}.", "progress"
    if min_reps_at_top is not None and min_reps_at_top < low:
        return "Ниже диапазона → вес пока рано увеличивать, добери повторы/технику.", "below"
    if rpe is not None and rpe >= 9:
        return "Высокий RPE → перегруз, оставь или немного снизь вес.", "overload"
    if rpe is not None and rpe <= 7:
        return "Легко → можно ускорить прогрессию.", "easy"
    return "Середина диапазона → оставь вес и добери повторы.", "stable"

def build_main_menu():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("Тренировка A", callback_data="start_A")],
        [InlineKeyboardButton("Тренировка B", callback_data="start_B")],
        [InlineKeyboardButton("Тренировка C", callback_data="start_C")],
        [InlineKeyboardButton("Dashboard", callback_data="dashboard_open"), InlineKeyboardButton("Экспорт Excel", callback_data="export_open")],
    ])

def build_day_menu(day, done_exercises=None, context_user_data=None):
    done = set(done_exercises or [])
    rows = []
    ex_map = {}  # idx -> exercise name
    for idx, ex in enumerate(PROGRAMS.get(day, {}).get("exercises", [])):
        ex_map[idx] = ex["exercise"]
        label = ("✅ " if ex["exercise"] in done else "") + ex["exercise"]
        rows.append([InlineKeyboardButton(label, callback_data=f"pick::{idx}")])
    if context_user_data is not None:
        context_user_data["day_ex_map"] = ex_map
    rows.append([InlineKeyboardButton("Каталог всех упражнений", callback_data="catalog::0")])
    rows.append([InlineKeyboardButton("Добавить своё упражнение", callback_data="custom_exercise")])
    rows.append([InlineKeyboardButton("✏️ Редактировать упражнение", callback_data="edit_exercise")])
    rows.append([InlineKeyboardButton("📊 Итог тренировки / Выйти", callback_data="finish_workout")])
    rows.append([InlineKeyboardButton("В меню", callback_data="go_menu")])
    return InlineKeyboardMarkup(rows)

def get_done_exercises(user_id, day, workout_date):
    rows = get_session_rows(user_id, day, workout_date)
    return set(r["exercise"] for r in rows)

def build_catalog_menu(page, context_user_data=None):
    start = page * PAGE_SIZE
    items = EXERCISE_CATALOG[start:start + PAGE_SIZE]
    rows = []
    cat_map = {}
    for local_idx, name in enumerate(items):
        global_idx = start + local_idx
        cat_map[global_idx] = name
        rows.append([InlineKeyboardButton(name, callback_data=f"cat_pick::{global_idx}")])
    if context_user_data is not None:
        existing = context_user_data.get("cat_ex_map", {})
        existing.update(cat_map)
        context_user_data["cat_ex_map"] = existing
    nav = []
    if page > 0:
        nav.append(InlineKeyboardButton("◀️ Назад", callback_data=f"catalog::{page-1}"))
    if start + PAGE_SIZE < len(EXERCISE_CATALOG):
        nav.append(InlineKeyboardButton("Вперед ▶️", callback_data=f"catalog::{page+1}"))
    if nav:
        rows.append(nav)
    rows.append([InlineKeyboardButton("К упражнениям дня", callback_data="back_to_day_menu")])
    rows.append([InlineKeyboardButton("Добавить своё упражнение", callback_data="custom_exercise")])
    rows.append([InlineKeyboardButton("📊 Итог тренировки / Выйти", callback_data="finish_workout")])
    return InlineKeyboardMarkup(rows)

def build_input_mode_menu(has_last):
    rows = []
    if has_last:
        rows.append([InlineKeyboardButton("Повторить прошлый вес", callback_data="use_last")])
    rows.append([InlineKeyboardButton("Быстрый ввод", callback_data="quick_input")])
    rows.append([InlineKeyboardButton("Ввести вручную", callback_data="manual_input")])
    rows.append([InlineKeyboardButton("Назад к упражнениям", callback_data="back_to_day_menu")])
    return InlineKeyboardMarkup(rows)

def build_edit_select_menu(session_rows):
    """Inline keyboard: pick which exercise from current session to edit."""
    rows = []
    for r in session_rows:
        sets_str = ", ".join(
            f"{r[f'set{i}_reps']:g}x{r[f'set{i}_kg']:g}"
            for i in range(1, 6)
            if r[f"set{i}_reps"] is not None and r[f"set{i}_kg"] is not None
        )
        label = f"{r['exercise']} [{sets_str}]"
        rows.append([InlineKeyboardButton(label, callback_data=f"edit_pick::{r['id']}")])
    rows.append([InlineKeyboardButton("Назад к упражнениям", callback_data="back_to_day_menu")])
    return InlineKeyboardMarkup(rows)

def build_edit_field_menu(row_id):
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("Изменить подходы (повторения×кг)", callback_data=f"edit_sets::{row_id}")],
        [InlineKeyboardButton("Изменить заметку", callback_data=f"edit_notes::{row_id}")],
        [InlineKeyboardButton("Назад к списку", callback_data="edit_exercise")],
    ])


def build_workout_summary(user_id, day, workout_date):
    rows = get_session_rows(user_id, day, workout_date)
    if not rows:
        return f"🏁 Тренировка {day} | {workout_date}\nНет записей."

    total_volume = round(sum(calc_volume(r) for r in rows), 1)
    progress_list = []
    overload_list = []
    scored = []

    for r in rows:
        tw, rr = calc_top_weight_and_reps(r)
        e1 = calc_e1rm(tw, rr)
        scored.append((r["exercise"], e1 if e1 is not None else 0, calc_volume(r)))
        status = r["coach_status"] or ""
        if status == "progress":
            progress_list.append(r["exercise"])
        elif status == "overload":
            overload_list.append(r["exercise"])

    best_by_volume = sorted(scored, key=lambda x: x[2], reverse=True)[:3]

    lines = [
        f"🏁 Тренировка {day} | {workout_date} завершена!",
        f"Упражнений: {len(rows)} | Тоннаж: {total_volume} кг",
        "",
        "Топ по тоннажу:",
    ]
    for ex, _, vol in best_by_volume:
        lines.append(f"• {ex}: {vol:g} кг")

    if progress_list:
        lines.append("")
        lines.append("Прогресс:")
        for ex in progress_list[:5]:
            lines.append(f"• {ex}")

    if overload_list:
        lines.append("")
        lines.append("Перегруз:")
        for ex in overload_list[:5]:
            lines.append(f"• {ex}")

    return "\n".join(lines)


async def start(update, context):
    text = "Привет. Это v6.1.\n\nТеперь после завершения тренировки бот выводит итог: тоннаж, лучшие упражнения, где был прогресс и где был перегруз."
    target = update.message if update.message else update.callback_query.message
    await target.reply_text(text, reply_markup=build_main_menu())

async def help_cmd(update, context):
    await update.message.reply_text(
        "Сценарий:\n1) Нажимаешь Тренировка A/B/C\n2) Вводишь дату\n3) Выбираешь упражнение\n4) Можно повторить прошлый вес, сделать быстрый ввод или ввести вручную\n5) После упражнения бот даёт авто-подсказку\n6) После завершения тренировки бот выдаёт итог\n7) /coach — недельный анализ"
    )

async def cancel(update, context):
    context.user_data.clear()
    target = update.message if update.message else update.callback_query.message
    if update.callback_query:
        await update.callback_query.answer()
    await target.reply_text("Ок, отменил.", reply_markup=ReplyKeyboardRemove())
    await target.reply_text("Меню:", reply_markup=build_main_menu())
    return ConversationHandler.END

async def menu_callback(update, context):
    q = update.callback_query
    await q.answer()
    data = q.data
    if data.startswith("start_"):
        day = data.split("_", 1)[1]
        context.user_data.clear()
        context.user_data["selected_day"] = day
        last_dt = get_last_workout_date(update.effective_user.id, day)
        review_msg = "📌 Контроль пересмотра тренировок: через 10 недель"
        if last_dt:
            try:
                last_date = datetime.fromisoformat(last_dt).date()
                review_msg = f"📌 Контроль пересмотра тренировок: {(last_date + timedelta(weeks=10)).isoformat()} (через 10 недель от последней {day})"
            except Exception:
                pass
        notes = "\n".join(f"• {x}" for x in PROGRAMS[day]["notes"])
        await q.message.reply_text(
            f"{PROGRAMS[day]['title']}\n\nРекомендации:\n{notes}\n\n{review_msg}\n\nДата тренировки? Формат ДД.ММ.ГГГГ",
            reply_markup=ReplyKeyboardRemove(),
        )
        return SESSION_DATE
    if data == "dashboard_open":
        await dashboard_text(q.message, update.effective_user.id)
        await q.message.reply_text("Меню:", reply_markup=build_main_menu())
        return ConversationHandler.END
    if data == "export_open":
        await export_file(q.message, update.effective_user.id)
        await q.message.reply_text("Меню:", reply_markup=build_main_menu())
        return ConversationHandler.END
    if data == "finish_workout":
        day = context.user_data.get("selected_day", "-")
        workout_date = context.user_data.get("workout_date", "-")
        summary = build_workout_summary(update.effective_user.id, day, workout_date)
        context.user_data.clear()
        await q.message.reply_text(summary, reply_markup=build_main_menu())
        return ConversationHandler.END
    if data == "go_menu":
        context.user_data.clear()
        await q.message.reply_text("Меню:", reply_markup=build_main_menu())
        return ConversationHandler.END
    if data == "back_to_day_menu":
        day = context.user_data.get("selected_day")
        if not day:
            await q.message.reply_text("Сначала начни тренировку.", reply_markup=build_main_menu())
            return ConversationHandler.END
        workout_date = context.user_data.get("workout_date", "")
        done = get_done_exercises(update.effective_user.id, day, workout_date)
        await q.message.reply_text(f"Упражнения дня {day}:", reply_markup=build_day_menu(day, done, context.user_data))
        return SESSION_MANUAL_INPUT
    if data.startswith("catalog::"):
        page = int(data.split("::", 1)[1])
        await q.message.reply_text("Каталог упражнений:", reply_markup=build_catalog_menu(page, context.user_data))
        return SESSION_MANUAL_INPUT
    if data == "custom_exercise":
        await q.message.reply_text("Напиши название своего упражнения:")
        return CUSTOM_EXERCISE_NAME
    if data.startswith("pick::"):
        idx = int(data.split("::", 1)[1])
        ex_map = context.user_data.get("day_ex_map", {})
        ex_name = ex_map.get(idx)
        if not ex_name:
            # fallback: rebuild map from PROGRAMS
            day = context.user_data.get("selected_day", "")
            exercises = PROGRAMS.get(day, {}).get("exercises", [])
            ex_name = exercises[idx]["exercise"] if idx < len(exercises) else None
        if not ex_name:
            await q.message.reply_text("Ошибка: упражнение не найдено.")
            return SESSION_MANUAL_INPUT
        context.user_data["current_exercise"] = ex_name
        day = context.user_data.get("selected_day", "")
        cfg = find_exercise_config(day, ex_name)
        prev_row = get_last_same_exercise(update.effective_user.id, ex_name)
        await q.message.reply_text(
            f"Выбрано: {ex_name}\nГруппа: {cfg['group']}\nЦель: {cfg['sets'] if cfg['sets'] else '-'} подходов × {cfg['reps']}\nОриентир: {cfg['rpe']}\nШаг прогрессии: {cfg['step']}\n{format_last_performance(prev_row)}\n\nКак хочешь занести данные?",
            reply_markup=build_input_mode_menu(prev_row is not None),
        )
        return SESSION_MANUAL_INPUT
    if data.startswith("cat_pick::"):
        idx = int(data.split("::", 1)[1])
        cat_map = context.user_data.get("cat_ex_map", {})
        ex_name = cat_map.get(idx)
        if not ex_name:
            ex_name = EXERCISE_CATALOG[idx] if idx < len(EXERCISE_CATALOG) else None
        if not ex_name:
            await q.message.reply_text("Ошибка: упражнение не найдено.")
            return SESSION_MANUAL_INPUT
        context.user_data["current_exercise"] = ex_name
        day = context.user_data.get("selected_day", "")
        cfg = find_exercise_config(day, ex_name)
        prev_row = get_last_same_exercise(update.effective_user.id, ex_name)
        await q.message.reply_text(
            f"Выбрано: {ex_name}\nГруппа: {cfg['group']}\nЦель: {cfg['sets'] if cfg['sets'] else '-'} подходов × {cfg['reps']}\nОриентир: {cfg['rpe']}\nШаг прогрессии: {cfg['step']}\n{format_last_performance(prev_row)}\n\nКак хочешь занести данные?",
            reply_markup=build_input_mode_menu(prev_row is not None),
        )
        return SESSION_MANUAL_INPUT
    if data == "use_last":
        ex_name = context.user_data.get("current_exercise")
        if not ex_name:
            await q.message.reply_text("Сначала выбери упражнение.")
            return ConversationHandler.END
        prev_row = get_last_same_exercise(update.effective_user.id, ex_name)
        if not prev_row:
            await q.message.reply_text("Нет прошлой записи. Используй быстрый ввод или ручной.")
            return SESSION_MANUAL_INPUT
        sets = []
        for i in range(1, 6):
            reps = prev_row[f"set{i}_reps"]
            kg = prev_row[f"set{i}_kg"]
            if reps is not None and kg is not None:
                sets.append((reps, kg))
        while len(sets) < 5:
            sets.append((None, None))
        context.user_data["current_sets"] = sets[:5]
        await q.message.reply_text(f"Подставил прошлый вариант: {format_last_performance(prev_row)}\nТеперь введи RPE или '-'")
        return SESSION_RPE
    if data == "quick_input":
        ex_name = context.user_data.get("current_exercise")
        if not ex_name:
            await q.message.reply_text("Сначала выбери упражнение.")
            return ConversationHandler.END
        day = context.user_data.get("selected_day", "")
        cfg = find_exercise_config(day, ex_name)
        default_sets = cfg["sets"] or 4
        await q.message.reply_text(
            f"Быстрый ввод в формате:\nвес повторения количество_подходов\nНапример: 60 8 {default_sets}\n\nЕсли количество подходов не укажешь, бот возьмёт целевое число подходов."
        )
        return SESSION_QUICK_INPUT
    if data == "manual_input":
        await q.message.reply_text("Введи подходы вручную. Пример:\n8x60, 8x60, 7x62.5, 6x62.5")
        return SESSION_MANUAL_INPUT
    if data == "edit_exercise":
        day = context.user_data.get("selected_day")
        workout_date = context.user_data.get("workout_date", "")
        if not day or not workout_date:
            await q.message.reply_text("Нет активной тренировки.")
            return ConversationHandler.END
        session_rows = get_session_rows(update.effective_user.id, day, workout_date)
        if not session_rows:
            await q.message.reply_text("В этой тренировке ещё нет записей для редактирования.")
            return SESSION_MANUAL_INPUT
        await q.message.reply_text("Выбери упражнение для редактирования:", reply_markup=build_edit_select_menu(session_rows))
        return SESSION_MANUAL_INPUT
    if data.startswith("edit_pick::"):
        row_id = int(data.split("::", 1)[1])
        context.user_data["edit_row_id"] = row_id
        await q.message.reply_text("Что хочешь изменить?", reply_markup=build_edit_field_menu(row_id))
        return SESSION_MANUAL_INPUT
    if data.startswith("edit_sets::"):
        row_id = int(data.split("::", 1)[1])
        context.user_data["edit_row_id"] = row_id
        context.user_data["edit_field"] = "sets"
        await q.message.reply_text(
            "Введи новые подходы в формате:\n8x60, 8x60, 7x62.5\n\nМожно указать от 1 до 5 подходов."
        )
        return SESSION_EDIT
    if data.startswith("edit_notes::"):
        row_id = int(data.split("::", 1)[1])
        context.user_data["edit_row_id"] = row_id
        context.user_data["edit_field"] = "notes"
        await q.message.reply_text("Введи новую заметку (или '-' чтобы очистить):")
        return SESSION_EDIT
    return ConversationHandler.END

async def session_date(update, context):
    try:
        context.user_data["workout_date"] = parse_date(update.message.text)
    except Exception as e:
        await update.message.reply_text(str(e))
        return SESSION_DATE
    day = context.user_data["selected_day"]
    done = get_done_exercises(update.effective_user.id, day, context.user_data["workout_date"])
    await update.message.reply_text(
        f"Готово. День {day} | дата: {context.user_data['workout_date']}\nТеперь выбирай упражнения в любом порядке:",
        reply_markup=build_day_menu(day, done, context.user_data),
    )
    return SESSION_MANUAL_INPUT

async def custom_exercise_name(update, context):
    ex_name = update.message.text.strip()
    if not ex_name:
        await update.message.reply_text("Название не может быть пустым.")
        return CUSTOM_EXERCISE_NAME
    context.user_data["current_exercise"] = ex_name
    prev_row = get_last_same_exercise(update.effective_user.id, ex_name)
    await update.message.reply_text(
        f"Добавлено своё упражнение: {ex_name}\n{format_last_performance(prev_row)}\n\nКак хочешь занести данные?",
        reply_markup=build_input_mode_menu(prev_row is not None),
    )
    return SESSION_MANUAL_INPUT

async def session_quick_input(update, context):
    ex_name = context.user_data.get("current_exercise")
    if not ex_name:
        await update.message.reply_text("Сначала выбери упражнение.")
        return ConversationHandler.END
    parts = update.message.text.strip().replace(",", ".").split()
    if len(parts) not in {2, 3}:
        await update.message.reply_text("Формат быстрого ввода: вес повторения количество_подходов\nПример: 60 8 4")
        return SESSION_QUICK_INPUT
    try:
        weight = float(parts[0])
        reps = float(parts[1])
        n_sets = int(float(parts[2])) if len(parts) == 3 else (find_exercise_config(context.user_data.get("selected_day", ""), ex_name)["sets"] or 4)
    except Exception:
        await update.message.reply_text("Не понял числа. Пример: 60 8 4")
        return SESSION_QUICK_INPUT
    context.user_data["current_sets"] = build_repeated_sets(weight, reps, n_sets)
    await update.message.reply_text(f"Ок, собрал {n_sets} подход(а/ов) по шаблону: {reps:g}x{weight:g}\nТеперь введи RPE или '-'")
    return SESSION_RPE

async def session_manual_input(update, context):
    if "current_exercise" not in context.user_data:
        day = context.user_data.get("selected_day")
        if day:
            workout_date = context.user_data.get("workout_date", "")
            done = get_done_exercises(update.effective_user.id, day, workout_date)
            await update.message.reply_text("Сначала выбери упражнение кнопкой ниже:", reply_markup=build_day_menu(day, done, context.user_data))
            return SESSION_MANUAL_INPUT
        await update.message.reply_text("Сначала начни тренировку.", reply_markup=build_main_menu())
        return ConversationHandler.END
    try:
        context.user_data["current_sets"] = parse_sets(update.message.text)
    except Exception as e:
        await update.message.reply_text(f"Ошибка: {e}\n\nПримеры: 15x40, 8х67.5, 6 * 67.5, 5 x 67.5")
        return SESSION_MANUAL_INPUT
    await update.message.reply_text("RPE? Если не используешь — '-'")
    return SESSION_RPE

async def session_rpe(update, context):
    try:
        context.user_data["current_rpe"] = parse_optional_float(update.message.text)
    except Exception:
        await update.message.reply_text("RPE должен быть числом или '-'.")
        return SESSION_RPE
    await update.message.reply_text("Заметка по упражнению? Если не нужна — '-'")
    return SESSION_NOTES

async def session_notes(update, context):
    user_id = update.effective_user.id
    ex_name = context.user_data["current_exercise"]
    day = context.user_data.get("selected_day", "")
    workout_date = context.user_data.get("workout_date", "")
    cfg = find_exercise_config(day, ex_name)
    notes = update.message.text.strip()
    notes = None if notes == "-" else notes
    sets = context.user_data["current_sets"]
    rpe = context.user_data["current_rpe"]
    suggestion, coach_status = analyze_progress(cfg, sets, rpe)

    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        "INSERT INTO workouts (user_id, workout_date, day_type, exercise, set1_reps, set1_kg, set2_reps, set2_kg, set3_reps, set3_kg, set4_reps, set4_kg, set5_reps, set5_kg, target_sets, target_reps, target_guidance, step_rule, rpe, notes, suggestion, coach_status) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
        (user_id, workout_date, day, ex_name, sets[0][0], sets[0][1], sets[1][0], sets[1][1], sets[2][0], sets[2][1], sets[3][0], sets[3][1], sets[4][0], sets[4][1], cfg["sets"], cfg["reps"], cfg["rpe"], cfg["step"], rpe, notes, suggestion, coach_status)
    )
    row_id = cur.lastrowid
    conn.commit()
    cur.execute("SELECT * FROM workouts WHERE id=?", (row_id,))
    row = cur.fetchone()
    conn.close()

    tw, rr = calc_top_weight_and_reps(row)
    e1rm = calc_e1rm(tw, rr)
    prev_best = get_prev_best_e1rm(user_id, ex_name, row_id)
    pr = "да" if (e1rm is not None and (prev_best is None or e1rm > prev_best)) else "нет"
    delta = None if (e1rm is None or prev_best is None) else round(e1rm - prev_best, 1)

    for key in ["current_exercise", "current_sets", "current_rpe"]:
        context.user_data.pop(key, None)

    done = get_done_exercises(user_id, day, workout_date)
    await update.message.reply_text(
        f"Сохранил: {ex_name} ✅\ne1RM: {e1rm if e1rm is not None else '-'}\nPR: {pr}\nΔ к прошлому лучшему: {delta if delta is not None else '-'}\n\nПодсказка тренера:\n{suggestion}\n\nВыбирай следующее упражнение:",
        reply_markup=build_day_menu(day, done, context.user_data),
    )
    return SESSION_MANUAL_INPUT

async def session_edit(update, context):
    """Handle text input when editing an existing exercise record."""
    user_id = update.effective_user.id
    row_id = context.user_data.get("edit_row_id")
    field = context.user_data.get("edit_field")
    day = context.user_data.get("selected_day", "")
    workout_date = context.user_data.get("workout_date", "")

    if not row_id or not field:
        await update.message.reply_text("Ошибка: нет данных для редактирования.")
        return ConversationHandler.END

    conn = get_conn()
    cur = conn.cursor()

    if field == "sets":
        try:
            sets = parse_sets(update.message.text)
        except Exception as e:
            await update.message.reply_text(f"Ошибка: {e}\n\nПримеры: 15x40, 8х67.5, 6 * 67.5, 5 x 67.5")
            return SESSION_EDIT
        cur.execute(
            """UPDATE workouts SET
               set1_reps=?, set1_kg=?,
               set2_reps=?, set2_kg=?,
               set3_reps=?, set3_kg=?,
               set4_reps=?, set4_kg=?,
               set5_reps=?, set5_kg=?
               WHERE id=? AND user_id=?""",
            (sets[0][0], sets[0][1], sets[1][0], sets[1][1],
             sets[2][0], sets[2][1], sets[3][0], sets[3][1],
             sets[4][0], sets[4][1], row_id, user_id)
        )
        conn.commit()
        # Recalculate suggestion/coach_status with updated sets
        cur.execute("SELECT * FROM workouts WHERE id=?", (row_id,))
        row = cur.fetchone()
        if row:
            cfg = find_exercise_config(day, row["exercise"])
            suggestion, coach_status = analyze_progress(cfg, sets, row["rpe"])
            cur.execute(
                "UPDATE workouts SET suggestion=?, coach_status=? WHERE id=?",
                (suggestion, coach_status, row_id)
            )
            conn.commit()
        conn.close()
        context.user_data.pop("edit_row_id", None)
        context.user_data.pop("edit_field", None)
        done = get_done_exercises(user_id, day, workout_date)
        await update.message.reply_text(
            "✅ Подходы обновлены.\n\nВыбирай следующее упражнение:",
            reply_markup=build_day_menu(day, done, context.user_data)
        )
        return SESSION_MANUAL_INPUT

    elif field == "notes":
        text = update.message.text.strip()
        new_notes = None if text == "-" else text
        cur.execute("UPDATE workouts SET notes=? WHERE id=? AND user_id=?", (new_notes, row_id, user_id))
        conn.commit()
        conn.close()
        context.user_data.pop("edit_row_id", None)
        context.user_data.pop("edit_field", None)
        done = get_done_exercises(user_id, day, workout_date)
        await update.message.reply_text(
            "✅ Заметка обновлена.\n\nВыбирай следующее упражнение:",
            reply_markup=build_day_menu(day, done, context.user_data)
        )
        return SESSION_MANUAL_INPUT

    conn.close()
    return SESSION_MANUAL_INPUT

async def dashboard_text(message, user_id):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT * FROM workouts WHERE user_id=?", (user_id,))
    rows = cur.fetchall()
    conn.close()

    workouts_cnt = len(set((r["workout_date"], r["day_type"]) for r in rows))
    total_volume = round(sum(calc_volume(r) for r in rows), 1)
    best_by_ex = {}
    for r in rows:
        tw, rr = calc_top_weight_and_reps(r)
        e1 = calc_e1rm(tw, rr)
        if e1 is not None:
            best_by_ex[r["exercise"]] = max(best_by_ex.get(r["exercise"], 0), e1)
    lines = ["📊 Dashboard", f"Тренировок: {workouts_cnt}", f"Тоннаж: {total_volume} кг", f"Упражнений с прогрессом: {len(best_by_ex)}"]
    if best_by_ex:
        lines += ["", "Лучшие e1RM:"]
        for ex, e1 in sorted(best_by_ex.items(), key=lambda x: x[1], reverse=True)[:8]:
            lines.append(f"• {ex}: {e1}")
    await message.reply_text("\n".join(lines))

async def dashboard(update, context):
    await dashboard_text(update.message, update.effective_user.id)

async def coach(update, context):
    user_id = update.effective_user.id
    cutoff = (datetime.utcnow().date() - timedelta(days=7)).isoformat()
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT * FROM workouts WHERE user_id=? AND workout_date>=? ORDER BY workout_date, id", (user_id, cutoff))
    rows = cur.fetchall()
    conn.close()
    if not rows:
        await update.message.reply_text("За последние 7 дней записей пока нет.")
        return

    workouts_cnt = len(set((r["workout_date"], r["day_type"]) for r in rows))
    total_volume = round(sum(calc_volume(r) for r in rows), 1)
    status_counts = {"progress": 0, "stable": 0, "below": 0, "overload": 0, "easy": 0, "unknown": 0}
    for r in rows:
        status = r["coach_status"] or "unknown"
        if status not in status_counts:
            status = "unknown"
        status_counts[status] += 1

    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT * FROM workouts WHERE user_id=? ORDER BY workout_date, id", (user_id,))
    all_rows = cur.fetchall()
    conn.close()

    by_ex = {}
    for r in all_rows:
        tw, rr = calc_top_weight_and_reps(r)
        e1 = calc_e1rm(tw, rr)
        if e1 is None:
            continue
        by_ex.setdefault(r["exercise"], []).append((r["workout_date"], e1))

    progress_lines = []
    regress = 0
    stagnation = 0
    improvements = 0
    for ex, vals in by_ex.items():
        if len(vals) >= 2:
            prev = vals[-2][1]
            last = vals[-1][1]
            diff = round(last - prev, 1)
            if diff > 0:
                improvements += 1
                progress_lines.append(f"+ {ex} → +{diff} кг e1RM")
            elif diff < 0:
                regress += 1
                progress_lines.append(f"- {ex} → {diff} кг e1RM")
            else:
                stagnation += 1
                progress_lines.append(f"= {ex} → без изменений")

    total_tracked = improvements + regress + stagnation
    balance_line = "Недостаточно данных."
    if total_tracked > 0:
        p = round(improvements / total_tracked * 100)
        s = round(stagnation / total_tracked * 100)
        r = round(regress / total_tracked * 100)
        balance_line = f"Баланс: {p}% прогресс / {s}% плато / {r}% регресс"

    recommendation = "Продолжаем."
    if status_counts["overload"] >= max(2, workouts_cnt):
        recommendation = "⚠️ Есть признаки перегруза: высокий RPE и/или просадка повторений. Оставь веса на 1–2 тренировки или чуть снизь объём."
    elif status_counts["easy"] >= max(2, workouts_cnt):
        recommendation = "⚠️ Есть признаки недогруза: упражнения ощущаются слишком легко. Можно ускорить прогрессию."
    elif regress > improvements:
        recommendation = "⚠️ Регресса больше, чем роста. Проверь восстановление, сон и объём."
    elif improvements >= regress and improvements > 0:
        recommendation = "✅ Вектор хороший. Большинство движений идут в нужную сторону."

    msg = [
        "🧠 Coach — анализ за последние 7 дней",
        f"Тренировок: {workouts_cnt}",
        f"Общий тоннаж: {total_volume} кг",
        "",
        "Статусы:",
        f"• прогресс: {status_counts['progress']}",
        f"• стабильно: {status_counts['stable']}",
        f"• ниже диапазона: {status_counts['below']}",
        f"• перегруз: {status_counts['overload']}",
        f"• легко: {status_counts['easy']}",
        "",
        balance_line,
        "",
        "Тренд по упражнениям:",
    ]
    msg.extend(progress_lines[:12] if progress_lines else ["Пока мало данных по повторным записям."])
    msg.extend(["", recommendation])
    await update.message.reply_text("\n".join(msg))

def create_excel_export(user_id, output_path):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT * FROM workouts WHERE user_id=? ORDER BY workout_date, id", (user_id,))
    workouts = cur.fetchall()
    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.title = "Log"
    headers = ["Date", "Week", "Day", "Exercise", "Set1 Reps", "Set1 Kg", "Set2 Reps", "Set2 Kg", "Set3 Reps", "Set3 Kg", "Set4 Reps", "Set4 Kg", "Set5 Reps", "Set5 Kg", "Target Sets", "Target Reps", "Target Guidance", "Step Rule", "RPE", "Volume", "Top Weight", "Top Reps@Top", "e1RM", "PR", "Suggestion", "Coach Status", "Notes"]
    ws.append(headers)

    navy = PatternFill("solid", fgColor="1F4E78")
    a_fill = PatternFill("solid", fgColor="DCEEFF")
    b_fill = PatternFill("solid", fgColor="E6F4DE")
    c_fill = PatternFill("solid", fgColor="FFF2D8")
    yellow = PatternFill("solid", fgColor="FFF2CC")

    for c in ws[1]:
        c.fill = navy
        c.font = Font(color="FFFFFF", bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center")

    first_date = None
    best_so_far = {}
    for r in workouts:
        if first_date is None:
            first_date = datetime.fromisoformat(r["workout_date"]).date()
        week = ((datetime.fromisoformat(r["workout_date"]).date() - first_date).days // 7) + 1
        tw, rr = calc_top_weight_and_reps(r)
        e1 = calc_e1rm(tw, rr)
        prev = best_so_far.get(r["exercise"])
        is_pr = e1 is not None and (prev is None or e1 > prev)
        if is_pr and e1 is not None:
            best_so_far[r["exercise"]] = e1
        ws.append([r["workout_date"], week, r["day_type"], r["exercise"], r["set1_reps"], r["set1_kg"], r["set2_reps"], r["set2_kg"], r["set3_reps"], r["set3_kg"], r["set4_reps"], r["set4_kg"], r["set5_reps"], r["set5_kg"], r["target_sets"], r["target_reps"], r["target_guidance"], r["step_rule"], r["rpe"], calc_volume(r), tw, rr, e1, "PR" if is_pr else "", r["suggestion"], r["coach_status"], r["notes"]])

    for row in range(2, ws.max_row + 1):
        day = ws.cell(row=row, column=3).value
        fill = a_fill if day == "A" else b_fill if day == "B" else c_fill if day == "C" else None
        if fill:
            for col in range(1, ws.max_column + 1):
                ws.cell(row=row, column=col).fill = fill
        if ws.cell(row=row, column=24).value == "PR":
            ws.cell(row=row, column=24).fill = yellow

    widths = [12,8,6,30,10,9,10,9,10,9,10,9,10,9,10,16,16,14,8,12,11,12,10,7,28,14,24]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[chr(64+i) if i <= 26 else "AA"].width = w
    ws.freeze_panes = "A2"

    prog = wb.create_sheet("Progress")
    prog.append(["Date", "Жим штанги лёжа", "Присед со штангой", "Подтягивания с весом", "Жим гантелей на наклонной скамье"])
    exercise_map = {"Жим штанги лёжа": 2, "Присед со штангой": 3, "Подтягивания с весом": 4, "Жим гантелей на наклонной скамье": 5}
    by_date = {}
    for r in workouts:
        tw, rr = calc_top_weight_and_reps(r)
        e1 = calc_e1rm(tw, rr)
        if e1 is None:
            continue
        dt = r["workout_date"]
        by_date.setdefault(dt, {})
        col = exercise_map.get(r["exercise"])
        if col:
            prev = by_date[dt].get(col)
            by_date[dt][col] = max(prev, e1) if prev is not None else e1
    for dt in sorted(by_date):
        row = [dt, None, None, None, None]
        for col, val in by_date[dt].items():
            row[col - 1] = val
        prog.append(row)
    for c in prog[1]:
        c.fill = navy
        c.font = Font(color="FFFFFF", bold=True)
    chart = LineChart()
    chart.title = "Strength Progress (e1RM)"
    chart.y_axis.title = "kg"
    chart.x_axis.title = "Date"
    data = Reference(prog, min_col=2, max_col=5, min_row=1, max_row=prog.max_row)
    cats = Reference(prog, min_col=1, min_row=2, max_row=prog.max_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.height = 8
    chart.width = 16
    prog.add_chart(chart, "G2")
    wb.save(output_path)
    return output_path

async def export_file(message, user_id):
    path = f"training_export_v61_{user_id}.xlsx"
    create_excel_export(user_id, path)
    with open(path, "rb") as f:
        await message.reply_document(document=f, filename=path, caption="Вот твоя выгрузка Excel 📊")

async def export_cmd(update, context):
    await export_file(update.message, update.effective_user.id)

def build_application():
    if not BOT_TOKEN:
        raise RuntimeError("Нужно задать TELEGRAM_BOT_TOKEN")
    init_db()
    app = ApplicationBuilder().token(BOT_TOKEN).build()

    conv = ConversationHandler(
        entry_points=[
            CallbackQueryHandler(menu_callback, pattern=r"^(start_[ABC]|dashboard_open|export_open|finish_workout|go_menu|back_to_day_menu|catalog::\d+|pick::\d+|cat_pick::\d+|custom_exercise|use_last|quick_input|manual_input|edit_exercise|edit_pick::\d+|edit_sets::\d+|edit_notes::\d+)$"),
            CommandHandler("trainer", start),
        ],
        states={
            SESSION_DATE: [MessageHandler(filters.TEXT & ~filters.COMMAND, session_date)],
            CUSTOM_EXERCISE_NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, custom_exercise_name)],
            SESSION_QUICK_INPUT: [MessageHandler(filters.TEXT & ~filters.COMMAND, session_quick_input)],
            SESSION_MANUAL_INPUT: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, session_manual_input),
                CallbackQueryHandler(menu_callback, pattern=r"^(finish_workout|go_menu|back_to_day_menu|catalog::\d+|pick::\d+|cat_pick::\d+|custom_exercise|use_last|quick_input|manual_input|edit_exercise|edit_pick::\d+|edit_sets::\d+|edit_notes::\d+)$"),
            ],
            SESSION_RPE: [MessageHandler(filters.TEXT & ~filters.COMMAND, session_rpe)],
            SESSION_NOTES: [MessageHandler(filters.TEXT & ~filters.COMMAND, session_notes)],
            SESSION_EDIT: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, session_edit),
                CallbackQueryHandler(menu_callback, pattern=r"^(edit_exercise|edit_pick::\d+|edit_sets::\d+|edit_notes::\d+|back_to_day_menu)$"),
            ],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
        allow_reentry=True,
    )

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("menu", start))
    app.add_handler(CommandHandler("help", help_cmd))
    app.add_handler(CommandHandler("dashboard", dashboard))
    app.add_handler(CommandHandler("coach", coach))
    app.add_handler(CommandHandler("export", export_cmd))
    app.add_handler(conv)
    app.add_handler(CommandHandler("cancel", cancel))
    return app

def main():
    app = build_application()
    app.run_polling()

if __name__ == "__main__":
    main()
