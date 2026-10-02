from dotenv import load_dotenv
load_dotenv()

import os
import io
import re
import html
import logging
import secrets
import string
import asyncio
from decimal import Decimal, InvalidOperation

import asyncpg
from telegram import (
    Update, InlineKeyboardButton, InlineKeyboardMarkup,
)
from telegram.constants import ParseMode
from telegram.ext import (
    Application, CommandHandler, CallbackQueryHandler,
    MessageHandler, filters, ContextTypes,
)

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer,
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

from yookassa import Configuration, Payment

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
)
log = logging.getLogger("smeta")

BOT_TOKEN = os.environ["BOT_TOKEN"]
DATABASE_URL = os.environ["DATABASE_URL"]

YOOKASSA_SHOP_ID = os.environ.get("YOOKASSA_SHOP_ID", "")
YOOKASSA_SECRET_KEY = os.environ.get("YOOKASSA_SECRET_KEY", "")
YOOKASSA_RETURN_URL = os.environ.get("YOOKASSA_RETURN_URL", "https://t.me")

Configuration.account_id = YOOKASSA_SHOP_ID
Configuration.secret_key = YOOKASSA_SECRET_KEY

pool = None
bot_username = None

ADMIN_IDS = [
    778239050,
]

CATEGORIES = [
    "🎬 Аренда",
    "📍 Локация",
    "💡 Свет",
    "📸 Камера",
    "🎙 Звук",
    "🎨 Реквизит и декор",
    "👕 Одежда",
    "💄 Макияж",
    "🚕 Транспорт и логистика",
    "🍔 Кейтеринг",
    "💼 Админка",
    "👥 Команда",
    "📦 Другое",
]

CATEGORY_ORDER = {name: i for i, name in enumerate(CATEGORIES)}
FALLBACK_ORDER = len(CATEGORIES) + 1

ROLE_OWNER = "owner"
ROLE_MEMBER = "member"
ROLE_VIEWER = "viewer"


SCHEMA = """
CREATE TABLE IF NOT EXISTS users (
    user_id     BIGINT PRIMARY KEY,
    username    TEXT,
    first_name  TEXT,
    last_name   TEXT,
    updated_at  TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS projects (
    id          SERIAL PRIMARY KEY,
    name        TEXT NOT NULL,
    budget      NUMERIC(14,2) NOT NULL DEFAULT 0,
    creator_id  BIGINT NOT NULL,
    is_personal BOOLEAN NOT NULL DEFAULT FALSE,
    created_at  TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS project_members (
    project_id  INTEGER NOT NULL REFERENCES projects(id) ON DELETE CASCADE,
    user_id     BIGINT NOT NULL,
    role        TEXT NOT NULL DEFAULT 'member',
    added_at    TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    PRIMARY KEY (project_id, user_id)
);

CREATE TABLE IF NOT EXISTS invites (
    project_id  INTEGER NOT NULL REFERENCES projects(id) ON DELETE CASCADE,
    code        TEXT NOT NULL UNIQUE,
    created_at  TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    PRIMARY KEY (project_id)
);

CREATE TABLE IF NOT EXISTS expenses (
    id               SERIAL PRIMARY KEY,
    project_id       INTEGER NOT NULL REFERENCES projects(id) ON DELETE CASCADE,
    category         TEXT NOT NULL DEFAULT '📦 Другое',
    name             TEXT NOT NULL,
    qty              NUMERIC(14,3) NOT NULL DEFAULT 1,
    price            NUMERIC(14,2) NOT NULL DEFAULT 0,
    comment          TEXT,
    author_id        BIGINT,
    author_username  TEXT,
    receipt_file_id  TEXT,
    created_at       TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_expenses_project ON expenses(project_id);
CREATE INDEX IF NOT EXISTS idx_members_user     ON project_members(user_id);
CREATE INDEX IF NOT EXISTS idx_members_project  ON project_members(project_id);

CREATE TABLE IF NOT EXISTS payments (
    id            SERIAL PRIMARY KEY,
    user_id       BIGINT NOT NULL,
    payment_id    TEXT NOT NULL UNIQUE,
    amount        NUMERIC(14,2) NOT NULL,
    status        TEXT NOT NULL DEFAULT 'pending',
    created_at    TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS subscriptions (
    user_id          BIGINT PRIMARY KEY,
    expires_at       TIMESTAMPTZ NOT NULL,
    updated_at       TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    notified_3days   BOOLEAN NOT NULL DEFAULT FALSE,
    notified_expired BOOLEAN NOT NULL DEFAULT FALSE
);
"""


async def migrate_db():
    async with pool.acquire() as c:
        await c.execute("ALTER TABLE projects ADD COLUMN IF NOT EXISTS is_personal BOOLEAN NOT NULL DEFAULT FALSE")
        await c.execute("ALTER TABLE projects ADD COLUMN IF NOT EXISTS created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()")
        await c.execute("ALTER TABLE expenses ADD COLUMN IF NOT EXISTS comment TEXT")
        await c.execute("ALTER TABLE expenses ADD COLUMN IF NOT EXISTS author_id BIGINT")
        await c.execute("ALTER TABLE expenses ADD COLUMN IF NOT EXISTS author_username TEXT")
        await c.execute("ALTER TABLE expenses ADD COLUMN IF NOT EXISTS created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()")
        await c.execute("ALTER TABLE expenses ADD COLUMN IF NOT EXISTS receipt_file_id TEXT")
        await c.execute("ALTER TABLE subscriptions ADD COLUMN IF NOT EXISTS notified_3days BOOLEAN NOT NULL DEFAULT FALSE")
        await c.execute("ALTER TABLE subscriptions ADD COLUMN IF NOT EXISTS notified_expired BOOLEAN NOT NULL DEFAULT FALSE")
        await c.execute("UPDATE projects SET is_personal=FALSE WHERE is_personal IS NULL")
        await c.execute("INSERT INTO project_members(project_id, user_id, role) SELECT id, creator_id, 'owner' FROM projects WHERE creator_id IS NOT NULL ON CONFLICT (project_id, user_id) DO NOTHING")
        await c.execute("INSERT INTO project_members(project_id, user_id, role) SELECT DISTINCT e.project_id, e.author_id, 'member' FROM expenses e WHERE e.author_id IS NOT NULL ON CONFLICT (project_id, user_id) DO NOTHING")
    log.info("Migration done")


async def init_db():
    global pool
    pool = await asyncpg.create_pool(DATABASE_URL, min_size=1, max_size=5)
    async with pool.acquire() as c:
        await c.execute(SCHEMA)
    await migrate_db()
    log.info("DB schema ready")


async def upsert_user(u):
    if u is None:
        return
    async with pool.acquire() as c:
        await c.execute(
            """
            INSERT INTO users(user_id, username, first_name, last_name, updated_at)
            VALUES($1,$2,$3,$4,NOW())
            ON CONFLICT (user_id) DO UPDATE
            SET username   = EXCLUDED.username,
                first_name = EXCLUDED.first_name,
                last_name  = EXCLUDED.last_name,
                updated_at = NOW()
            """,
            u.id, u.username, u.first_name, u.last_name,
        )


async def get_user_role(project_id, user_id):
    async with pool.acquire() as c:
        row = await c.fetchrow(
            "SELECT role FROM project_members WHERE project_id=$1 AND user_id=$2",
            project_id, user_id,
        )
    return row["role"] if row else None


async def user_can_view_project(project_id, user_id):
    async with pool.acquire() as c:
        p = await c.fetchrow(
            "SELECT id, creator_id, is_personal FROM projects WHERE id=$1",
            project_id,
        )
        if not p:
            return False
        if p["is_personal"]:
            return p["creator_id"] == user_id
        role = await c.fetchval(
            "SELECT role FROM project_members WHERE project_id=$1 AND user_id=$2",
            project_id, user_id,
        )
        return role is not None


async def user_can_edit_project(project_id, user_id):
    role = await get_user_role(project_id, user_id)
    return role == ROLE_OWNER


def _gen_invite_code(n=8):
    alphabet = string.ascii_letters + string.digits
    return "".join(secrets.choice(alphabet) for _ in range(n))


async def get_or_create_invite_code(project_id):
    async with pool.acquire() as c:
        row = await c.fetchrow(
            "SELECT code FROM invites WHERE project_id=$1", project_id
        )
        if row:
            return row["code"]
        for _ in range(5):
            code = _gen_invite_code()
            try:
                await c.execute(
                    "INSERT INTO invites(project_id, code) VALUES($1,$2)",
                    project_id, code,
                )
                return code
            except asyncpg.UniqueViolationError:
                continue
        raise RuntimeError("Не удалось создать код приглашения")


def invite_link(code):
    if not bot_username:
        return f"https://t.me/?start=join_{code}"
    return f"https://t.me/{bot_username}?start=join_{code}"


async def get_subscription_info(user_id):
    async with pool.acquire() as c:
        sub = await c.fetchrow(
            "SELECT expires_at FROM subscriptions WHERE user_id=$1 AND expires_at > NOW()",
            user_id,
        )
    if sub:
        return True, sub["expires_at"]
    return False, None


async def user_has_subscription(user_id):
    is_active, _ = await get_subscription_info(user_id)
    return is_active


async def create_yookassa_payment(user_id, amount, description="Доступ к боту"):
    if not Configuration.account_id or not Configuration.secret_key:
        return None, "Платежи временно недоступны."
    try:
        payment = Payment.create({
            "amount": {"value": f"{amount:.2f}", "currency": "RUB"},
            "confirmation": {
                "type": "redirect",
                "return_url": YOOKASSA_RETURN_URL,
            },
            "capture": True,
            "description": description,
            "metadata": {"user_id": str(user_id)},
        }, str(user_id) + str(amount) + str(int(asyncio.get_event_loop().time())))
        async with pool.acquire() as c:
            await c.execute(
                "INSERT INTO payments(user_id, payment_id, amount, status) VALUES($1,$2,$3,$4)",
                user_id, payment.id, amount, payment.status,
            )
        return payment.confirmation.confirmation_url, None
    except Exception as e:
        log.error("YooKassa create payment failed: %s", e)
        return None, "Не удалось создать платёж. Попробуйте позже."


async def check_yookassa_payment(user_id, payment_id):
    if not Configuration.account_id or not Configuration.secret_key:
        return False, "Платежи недоступны."
    try:
        payment = Payment.find_one(payment_id)
        if payment.status == "succeeded":
            amount = float(payment.amount.value)
            days = 90 if amount >= 1300 else 30
            async with pool.acquire() as c:
                await c.execute(
                    "UPDATE payments SET status='succeeded' WHERE payment_id=$1",
                    payment_id,
                )
                exists = await c.fetchval(
                    "SELECT 1 FROM subscriptions WHERE user_id=$1", user_id
                )
                if exists:
                    await c.execute(
                        "UPDATE subscriptions SET expires_at = GREATEST(expires_at, NOW()) + ($1 || ' days')::interval, notified_3days=FALSE, notified_expired=FALSE, updated_at=NOW() WHERE user_id=$2",
                        str(days), user_id,
                    )
                else:
                    await c.execute(
                        "INSERT INTO subscriptions(user_id, expires_at, updated_at) VALUES($1, NOW() + ($2 || ' days')::interval, NOW())",
                        user_id, str(days),
                    )
            return True, f"Оплата прошла! Доступ активирован на {days} дней."
        else:
            return False, f"Статус платежа: {payment.status}. Попробуйте позже."
    except Exception as e:
        log.error("YooKassa check payment failed: %s", e)
        return False, "Не удалось проверить платёж."


def money(v) -> str:
    if v is None:
        v = 0
    d = Decimal(str(v))
    sign = "-" if d < 0 else ""
    d = abs(d).quantize(Decimal("0.01"))
    s = f"{d:,.2f}".replace(",", " ").replace(".", ",")
    return f"{sign}{s} ₽"


def fmt_qty(q) -> str:
    d = Decimal(str(q)).normalize()
    if d == d.to_integral_value():
        return str(int(d))
    return str(d).replace(".", ",")


def esc(s) -> str:
    return html.escape(str(s)) if s is not None else ""


def _safe_name(name):
    s = re.sub(r"[^\w\-]+", "_", name or "").strip("_")
    return (s[:40] or "smeta")


def parse_money(text):
    if not text:
        return None
    t = text.strip().lower()
    for token in ("₽", "руб.", "руб", "р.", "rub", "rur"):
        t = t.replace(token, "")
    t = t.replace(" ", "").replace("\u00a0", "")
    if "," in t and "." in t:
        if t.rfind(",") > t.rfind("."):
            t = t.replace(".", "").replace(",", ".")
        else:
            t = t.replace(",", "")
    else:
        t = t.replace(",", ".")
    try:
        d = Decimal(t)
        if d < 0:
            return None
        return d.quantize(Decimal("0.01"))
    except InvalidOperation:
        return None


def parse_number(text):
    if not text:
        return None
    t = text.strip().replace(",", ".").replace(" ", "")
    try:
        d = Decimal(t)
        if d < 0:
            return None
        return d
    except InvalidOperation:
        return None


MEASURE_WORDS = {
    "шт", "штук", "штука", "штуки",
    "банок", "банка", "банки",
    "кг", "г", "гр", "л", "мл",
    "м", "см", "мм",
    "упаковок", "упаковка", "упаковки",
    "пачек", "пачка", "пачки",
    "бутылок", "бутылка", "бутылки",
    "коробок", "коробка", "коробки",
    "рулон", "рулона", "рулонов",
}


def _num(s):
    return Decimal(s.replace(",", "."))


def _clean_name(s):
    parts = [p for p in s.strip().split() if p]
    while parts and parts[0].lower() in MEASURE_WORDS:
        parts.pop(0)
    name = " ".join(parts) if parts else s.strip()
    if name:
        name = name[0].upper() + name[1:]
    return name


def parse_quick(text):
    if not text:
        return None
    t = text.strip()
    low = t.lower()
    m = re.match(
        r"^(\d+(?:[.,]\d+)?)\s+(.+?)\s+по\s+(\d+(?:[.,]\d+)?)\s*"
        r"(?:₽|руб\.?|р\.?)?$",
        low,
    )
    if m:
        return _clean_name(m.group(2)), _num(m.group(1)), _num(m.group(3))
    m = re.match(
        r"^(\d+(?:[.,]\d+)?)\s+(.+?)\s+(\d+(?:[.,]\d+)?)\s*"
        r"(?:₽|руб\.?|р\.?)?$",
        low,
    )
    if m:
        return _clean_name(m.group(2)), _num(m.group(1)), _num(m.group(3))
    m = re.match(
        r"^(.+?)\s+(\d+(?:[.,]\d+)?)\s*[x×*]\s*(\d+(?:[.,]\d+)?)\s*"
        r"(?:₽|руб\.?|р\.?)?$",
        low,
    )
    if m:
        return _clean_name(m.group(1)), _num(m.group(2)), _num(m.group(3))
    m = re.match(
        r"^([^\d].*?)\s+(\d+(?:[.,]\d+)?)\s*(?:₽|руб\.?|р\.?)?$",
        low,
    )
    if m:
        return _clean_name(m.group(1)), Decimal(1), _num(m.group(2))
    return None


def _sort_rows_by_category(rows):
    return sorted(
        rows,
        key=lambda r: (
            CATEGORY_ORDER.get(r["category"], FALLBACK_ORDER),
            r["id"],
        ),
    )


def _register_cyrillic_fonts():
    regular_candidates = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/dejavu/DejaVuSans.ttf",
    ]
    bold_candidates = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "/usr/share/fonts/dejavu/DejaVuSans-Bold.ttf",
    ]
    reg_path = next((p for p in regular_candidates if os.path.exists(p)), None)
    bold_path = next((p for p in bold_candidates if os.path.exists(p)), None)
    if reg_path:
        pdfmetrics.registerFont(TTFont("DejaVu", reg_path))
        font = "DejaVu"
    else:
        log.warning("DejaVuSans.ttf not found, fallback to Helvetica")
        font = "Helvetica"
    if bold_path:
        pdfmetrics.registerFont(TTFont("DejaVu-Bold", bold_path))
        font_bold = "DejaVu-Bold"
    else:
        font_bold = "Helvetica-Bold"
    return font, font_bold


def kb_main():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("👤 Личные проекты",    callback_data="projects:personal")],
        [InlineKeyboardButton("👥 Командные проекты", callback_data="projects:team")],
        [InlineKeyboardButton("➕ Новый проект",       callback_data="project:new")],
        [InlineKeyboardButton("💎 Подписка",           callback_data="buy:access")],
    ])


def kb_back(cb="menu"):
    return InlineKeyboardMarkup([[InlineKeyboardButton("◀️ Назад", callback_data=cb)]])


def kb_project(pid, role=None):
    rows = [
        [InlineKeyboardButton("➕ Добавить расход", callback_data=f"exp:add:{pid}")],
        [
            InlineKeyboardButton("📋 Расходы", callback_data=f"exp:list:{pid}"),
            InlineKeyboardButton("📊 Excel",  callback_data=f"excel:{pid}"),
            InlineKeyboardButton("📄 PDF",    callback_data=f"pdf:{pid}"),
        ],
        [
            InlineKeyboardButton("📎 Прикрепить чек", callback_data=f"receipt:choose:{pid}"),
        ],
    ]
    if role == ROLE_OWNER:
        rows.append([
            InlineKeyboardButton("👥 Пригласить",  callback_data=f"project:invite:{pid}"),
            InlineKeyboardButton("⚙️ Настройки",   callback_data=f"project:settings:{pid}"),
        ])
        rows.append([
            InlineKeyboardButton("🗑 Удалить проект", callback_data=f"project:delete:{pid}"),
        ])
    rows.append([InlineKeyboardButton("◀️ К списку", callback_data="menu")])
    return InlineKeyboardMarkup(rows)


def kb_categories(pid):
    rows = [[InlineKeyboardButton(c, callback_data=f"cat:{pid}:{i}")] for i, c in enumerate(CATEGORIES)]
    rows.append([InlineKeyboardButton("❌ Отмена", callback_data=f"exp:cancel:{pid}")])
    return InlineKeyboardMarkup(rows)


def kb_comment(pid):
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("⏭ Пропустить", callback_data=f"exp:skip_comment:{pid}")],
        [InlineKeyboardButton("❌ Отмена",     callback_data=f"exp:cancel:{pid}")],
    ])


def kb_confirm(pid):
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("✅ Сохранить", callback_data=f"exp:save:{pid}")],
        [InlineKeyboardButton("❌ Отмена",     callback_data=f"exp:cancel:{pid}")],
    ])


def kb_project_type():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("👤 Личный",    callback_data="newp:type:personal")],
        [InlineKeyboardButton("👥 Командный", callback_data="newp:type:team")],
        [InlineKeyboardButton("❌ Отмена",    callback_data="menu")],
    ])


def kb_invite_menu(pid):
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("📝 По @username",  callback_data=f"project:invite:user:{pid}")],
        [InlineKeyboardButton("🔗 Ссылка",        callback_data=f"project:invite:link:{pid}")],
        [InlineKeyboardButton("◀️ Назад к проекту", callback_data=f"project:view:{pid}")],
    ])


def kb_change_type_menu(pid, current_is_personal):
    if current_is_personal:
        return InlineKeyboardMarkup([
            [InlineKeyboardButton("👥 Сделать командным", callback_data=f"project:type:team:{pid}")],
            [InlineKeyboardButton("◀️ Назад",             callback_data=f"project:settings:{pid}")],
        ])
    else:
        return InlineKeyboardMarkup([
            [InlineKeyboardButton("👤 Сделать личным",    callback_data=f"project:type:personal:{pid}")],
            [InlineKeyboardButton("◀️ Назад",             callback_data=f"project:settings:{pid}")],
        ])


async def project_view(pid, user_id=None):
    async with pool.acquire() as c:
        p = await c.fetchrow("SELECT * FROM projects WHERE id=$1", pid)
        if not p:
            return None, None
        t = await c.fetchrow(
            "SELECT COALESCE(SUM(qty*price),0) AS total FROM expenses WHERE project_id=$1",
            pid,
        )
    total = Decimal(str(t["total"]))
    budget = Decimal(str(p["budget"]))
    left = budget - total
    status = "🟢" if left >= 0 else "🔴"
    type_label = "👤 Личный" if p["is_personal"] else "👥 Командный"
    text = (
        f"📁 <b>{esc(p['name'])}</b>\n"
        f"<i>{type_label}</i>\n\n"
        f"💰 Бюджет: {money(budget)}\n"
        f"💸 Потрачено: {money(total)}\n"
        f"{status} Осталось: {money(left)}"
    )
    role = None
    if user_id is not None:
        role = await get_user_role(pid, user_id)
        if p["is_personal"] and p["creator_id"] == user_id:
            role = ROLE_OWNER
    return text, kb_project(pid, role)


async def send_project_msg(target, pid, edit=False, user_id=None):
    text, kb = await project_view(pid, user_id=user_id)
    if text is None:
        msg = "Проект не найден."
        if edit and hasattr(target, "edit_message_text"):
            try:
                return await target.edit_message_text(msg)
            except Exception:
                pass
        return await target.reply_text(msg)
    if edit and hasattr(target, "edit_message_text"):
        try:
            return await target.edit_message_text(
                text, parse_mode=ParseMode.HTML, reply_markup=kb
            )
        except Exception as e:
            log.warning("edit_message_text failed: %s", e)
    return await target.reply_text(text, parse_mode=ParseMode.HTML, reply_markup=kb)


async def projects_list_view(user_id, mode):
    async with pool.acquire() as c:
        if mode == "personal":
            rows = await c.fetch(
                """
                SELECT p.id, p.name, p.budget, p.created_at,
                       COALESCE(SUM(e.qty*e.price),0) AS spent,
                       COUNT(e.id) AS cnt
                FROM projects p
                LEFT JOIN expenses e ON e.project_id = p.id
                WHERE p.is_personal = TRUE AND p.creator_id = $1
                GROUP BY p.id
                ORDER BY p.created_at DESC
                """,
                user_id,
            )
            title = "👤 <b>Личные проекты</b>"
            empty = "Личных проектов пока нет."
        else:
            rows = await c.fetch(
                """
                SELECT p.id, p.name, p.budget, p.created_at,
                       COALESCE(SUM(e.qty*e.price),0) AS spent,
                       COUNT(e.id) AS cnt,
                       pm.role AS my_role
                FROM projects p
                JOIN project_members pm ON pm.project_id = p.id
                LEFT JOIN expenses e ON e.project_id = p.id
                WHERE p.is_personal = FALSE AND pm.user_id = $1
                GROUP BY p.id, pm.role
                ORDER BY p.created_at DESC
                """,
                user_id,
            )
            title = "👥 <b>Командные проекты</b>"
            empty = "Командных проектов пока нет."
    if not rows:
        text = f"{title}\n\n{empty}"
        kb = InlineKeyboardMarkup([
            [InlineKeyboardButton("➕ Новый проект", callback_data="project:new")],
            [InlineKeyboardButton("🏠 Меню",         callback_data="menu")],
        ])
        return text, kb
    lines = [title + "\n"]
    for r in rows:
        total = Decimal(str(r["spent"]))
        left = Decimal(str(r["budget"])) - total
        mark = "🟢" if left >= 0 else "🔴"
        role_icon = ""
        if mode == "team":
            role = r["my_role"]
            role_icon = " 👑" if role == ROLE_OWNER else ""
        lines.append(
            f"• <b>{esc(r['name'])}</b>{role_icon} — {money(total)} / {money(r['budget'])} {mark}\n"
            f"  расходов: {r['cnt']}"
        )
    kb_rows = []
    for r in rows:
        icon = "👤" if mode == "personal" else "👥"
        kb_rows.append([
            InlineKeyboardButton(
                f"{icon} {r['name'][:40]}",
                callback_data=f"project:view:{r['id']}",
            )
        ])
    kb_rows.append([InlineKeyboardButton("➕ Новый проект", callback_data="project:new")])
    kb_rows.append([InlineKeyboardButton("🏠 Меню",         callback_data="menu")])
    return "\n".join(lines), InlineKeyboardMarkup(kb_rows)


def _exp_summary(exp):
    qv = Decimal(str(exp.get("qty", "1")))
    pv = Decimal(str(exp.get("price", "0")))
    s = qv * pv
    lines = [
        "<b>Проверьте расход:</b>",
        f"• Категория: {esc(exp.get('category', '—'))}",
        f"• Название: {esc(exp.get('name', '—'))}",
        f"• Количество: {fmt_qty(qv)}",
        f"• Цена: {money(pv)}",
        f"• Сумма: {money(s)}",
    ]
    if exp.get("comment"):
        lines.append(f"• Комментарий: {esc(exp['comment'])}")
    return "\n".join(lines)


async def make_excel(update, pid):
    q = update.callback_query
    user_id = update.effective_user.id
    if not await user_can_view_project(pid, user_id):
        return await q.message.reply_text("У вас нет доступа к этому проекту.")
    async with pool.acquire() as c:
        p = await c.fetchrow("SELECT * FROM projects WHERE id=$1", pid)
        rows = await c.fetch("SELECT * FROM expenses WHERE project_id=$1", pid)
    if not p:
        return await q.message.reply_text("Проект не найден.")
    rows = _sort_rows_by_category(rows)
    total = sum((Decimal(str(r["qty"])) * Decimal(str(r["price"])) for r in rows), Decimal(0))
    budget = Decimal(str(p["budget"]))
    left = budget - total
    wb = Workbook()
    ws = wb.active
    ws.title = "Смета"
    ws["A1"] = "СМЕТА"
    ws["A1"].font = Font(bold=True, size=16)
    ws.merge_cells("A1:I1")
    ws["A2"] = "Проект";     ws["B2"] = p["name"]
    ws["A3"] = "Тип";        ws["B3"] = "Личный" if p["is_personal"] else "Командный"
    ws["A4"] = "Бюджет";     ws["B4"] = float(budget)
    ws["A5"] = "Потрачено";  ws["B5"] = float(total)
    ws["A6"] = "Осталось";   ws["B6"] = float(left)
    headers = [
        "Категория", "Наименование", "Количество", "Цена", "Сумма",
        "Комментарий", "Кто добавил", "Дата", "Чек",
    ]
    ws.append([])
    ws.append(headers)
    header_row = ws.max_row
    fill = PatternFill("solid", fgColor="DDDDDD")
    bold = Font(bold=True)
    for col in range(1, len(headers) + 1):
        cell = ws.cell(row=header_row, column=col)
        cell.font = bold
        cell.fill = fill
        cell.alignment = Alignment(horizontal="center")
    for r in rows:
        qv = Decimal(str(r["qty"]))
        pv = Decimal(str(r["price"]))
        s = qv * pv
        dt = r["created_at"].strftime("%d.%m.%Y %H:%M") if r["created_at"] else ""
        author = ("@" + r["author_username"]) if r["author_username"] else (
            f"id{r['author_id']}" if r["author_id"] else ""
        )
        receipt_mark = "📎 Есть чек" if r.get("receipt_file_id") else ""
        ws.append([
            r["category"], r["name"], float(qv), float(pv), float(s),
            r["comment"] or "", author, dt, receipt_mark,
        ])
    total_row = ws.max_row + 1
    ws.cell(row=total_row, column=4, value="ИТОГО").font = bold
    ws.cell(row=total_row, column=5, value=float(total)).font = bold
    for i, w in enumerate([18, 28, 12, 12, 14, 30, 18, 18, 12], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    await q.message.reply_document(buf, filename=f"smeta_{_safe_name(p['name'])}.xlsx")


async def make_pdf(update, pid):
    q = update.callback_query
    user_id = update.effective_user.id
    if not await user_can_view_project(pid, user_id):
        return await q.message.reply_text("У вас нет доступа к этому проекту.")
    async with pool.acquire() as c:
        p = await c.fetchrow("SELECT * FROM projects WHERE id=$1", pid)
        rows = await c.fetch("SELECT * FROM expenses WHERE project_id=$1", pid)
    if not p:
        return await q.message.reply_text("Проект не найден.")
    rows = _sort_rows_by_category(rows)
    total = sum((Decimal(str(r["qty"])) * Decimal(str(r["price"])) for r in rows), Decimal(0))
    budget = Decimal(str(p["budget"]))
    left = budget - total
    FONT, FONT_BOLD = _register_cyrillic_fonts()
    out = io.BytesIO()
    doc = SimpleDocTemplate(
        out, pagesize=A4,
        leftMargin=30, rightMargin=30, topMargin=30, bottomMargin=30,
    )
    title_style = ParagraphStyle("T",  fontName=FONT_BOLD, fontSize=22, leading=26, alignment=1)
    h2_style    = ParagraphStyle("H2", fontName=FONT_BOLD, fontSize=14, leading=18, spaceAfter=6)
    body_style  = ParagraphStyle("B",  fontName=FONT,      fontSize=11, leading=15)
    small_style = ParagraphStyle("S",  fontName=FONT,      fontSize=8,  leading=10)
    small_bold  = ParagraphStyle("SB", fontName=FONT_BOLD, fontSize=8,  leading=10)
    cat_style   = ParagraphStyle("C",  fontName=FONT_BOLD, fontSize=10, leading=13)
    type_label = "Личный" if p["is_personal"] else "Командный"
    story = [
        Paragraph("СМЕТА", title_style),
        Paragraph(esc(p["name"]), h2_style),
        Paragraph(f"Тип: {type_label}", body_style),
        Spacer(1, 6),
        Paragraph(f"Бюджет: {money(budget)}", body_style),
        Paragraph(f"Потрачено: {money(total)}", body_style),
        Paragraph(f"Осталось: {money(left)}", body_style),
        Spacer(1, 12),
    ]
    def P(t, st=None):
        return Paragraph(esc(t), st or small_style)
    # Группируем по категориям
    from collections import OrderedDict
    groups = OrderedDict()
    for r in rows:
        groups.setdefault(r["category"], []).append(r)
    data = [[
        Paragraph("Наименование", small_bold),
        Paragraph("Кол-во", small_bold),
        Paragraph("Цена", small_bold),
        Paragraph("Сумма", small_bold),
        Paragraph("Чек", small_bold),
    ]]
    cat_header_rows = []
    subtotal_rows = []
    for cat_name, items in groups.items():
        cat_header_rows.append(len(data))
        data.append([
            Paragraph(f"<b>{esc(cat_name)}</b>", cat_style),
            "", "", "", "",
        ])
        cat_total = Decimal(0)
        for r in items:
            qv = Decimal(str(r["qty"]))
            pv = Decimal(str(r["price"]))
            s = qv * pv
            cat_total += s
            receipt_mark = "📎" if r.get("receipt_file_id") else ""
            data.append([
                P(r["name"]),
                P(fmt_qty(qv)),
                P(money(pv)),
                P(money(s)),
                P(receipt_mark),
            ])
        subtotal_rows.append(len(data))
        data.append([
            Paragraph("<b>Итого по категории:</b>", small_bold),
            "",
            "",
            Paragraph(f"<b>{money(cat_total)}</b>", small_bold),
            "",
        ])
    # Общий итог
    data.append([
        Paragraph("<b>ВСЕГО:</b>", cat_style),
        "",
        "",
        Paragraph(f"<b>{money(total)}</b>", cat_style),
        "",
    ])
    table = Table(data, repeatRows=1, colWidths=[260, 55, 75, 85, 30])
    style_cmds = [
        ("GRID", (0, 0), (-1, -1), .4, colors.grey),
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("ALIGN", (1, 1), (-1, -1), "RIGHT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]
    for idx in cat_header_rows:
        style_cmds.append(("BACKGROUND", (0, idx), (-1, idx), colors.Color(0.94, 0.94, 0.94)))
        style_cmds.append(("SPAN", (0, idx), (-1, idx)))
    for idx in subtotal_rows:
        style_cmds.append(("BACKGROUND", (0, idx), (-1, idx), colors.Color(0.97, 0.97, 0.92)))
    # Общая итоговая строка
    last = len(data) - 1
    style_cmds.append(("BACKGROUND", (0, last), (-1, last), colors.Color(0.88, 0.88, 0.88)))
    table.setStyle(TableStyle(style_cmds))
    story.append(table)
    doc.build(story)
    out.seek(0)
    await q.message.reply_document(out, filename=f"smeta_{_safe_name(p['name'])}.pdf")


async def check_access_or_paywall(update, ctx, u):
    is_active, expires = await get_subscription_info(u.id)
    if is_active:
        return True
    kb = InlineKeyboardMarkup([
        [InlineKeyboardButton("💎 Оформить подписку", callback_data="buy:access")],
    ])
    text = (
        "🔒 <b>Доступ к боту по подписке</b>\n\n"
        "Оформите подписку от <b>500 ₽</b>, чтобы пользоваться ботом."
    )
    if update.callback_query:
        await update.callback_query.edit_message_text(
            text, parse_mode=ParseMode.HTML, reply_markup=kb
        )
    else:
        await update.message.reply_text(
            text, parse_mode=ParseMode.HTML, reply_markup=kb
        )
    return False


async def cmd_start(update, ctx):
    u = update.effective_user
    await upsert_user(u)
    if not await check_access_or_paywall(update, ctx, u):
        return
    args = ctx.args or []
    if args and args[0].startswith("join_"):
        return await handle_join_link(update, ctx, args[0][5:])
    text = (
        f"👋 Привет, {esc(u.first_name or 'друг')}!\n\n"
        "Это бот для командной работы со сметами.\n"
        "Выберите раздел:"
    )
    await update.message.reply_text(text, parse_mode=ParseMode.HTML, reply_markup=kb_main())


async def handle_join_link(update, ctx, code):
    u = update.effective_user
    async with pool.acquire() as c:
        row = await c.fetchrow("SELECT project_id FROM invites WHERE code=$1", code)
    if not row:
        return await update.message.reply_text(
            "Ссылка-приглашение недействительна или устарела.",
            reply_markup=kb_main(),
        )
    pid = row["project_id"]
    async with pool.acquire() as c:
        p = await c.fetchrow("SELECT id, name, is_personal FROM projects WHERE id=$1", pid)
    if not p:
        return await update.message.reply_text("Проект больше не существует.", reply_markup=kb_main())
    if p["is_personal"]:
        return await update.message.reply_text(
            "Этот проект сейчас личный — создатель отключил приглашения.",
            reply_markup=kb_main(),
        )
    async with pool.acquire() as c:
        await c.execute(
            "INSERT INTO project_members(project_id, user_id, role) VALUES($1,$2,$3) "
            "ON CONFLICT (project_id, user_id) DO NOTHING",
            pid, u.id, ROLE_MEMBER,
        )
    await update.message.reply_text(
        f"✅ Вы присоединились к проекту «{esc(p['name'])}».",
        parse_mode=ParseMode.HTML,
    )
    return await send_project_msg(update.message, pid, edit=False, user_id=u.id)


async def cmd_menu(update, ctx):
    u = update.effective_user
    await upsert_user(u)
    if not await check_access_or_paywall(update, ctx, u):
        return
    await update.message.reply_text("Меню:", reply_markup=kb_main())


async def cmd_admin(update, ctx):
    u = update.effective_user
    if u.id not in ADMIN_IDS:
        return await update.message.reply_text("⛔ Нет доступа.")
    await upsert_user(u)
    async with pool.acquire() as c:
        total_users = await c.fetchval("SELECT COUNT(*) FROM users")
        paid_users = await c.fetchval(
            "SELECT COUNT(DISTINCT user_id) FROM payments WHERE status='succeeded'"
        )
        active_subs = await c.fetchval(
            "SELECT COUNT(*) FROM subscriptions WHERE expires_at > NOW()"
        )
        total_income = await c.fetchval(
            "SELECT COALESCE(SUM(amount),0) FROM payments WHERE status='succeeded'"
        )
        projects_cnt = await c.fetchval("SELECT COUNT(*) FROM projects")
        expenses_cnt = await c.fetchval("SELECT COUNT(*) FROM expenses")
        last_payments = await c.fetch(
            "SELECT user_id, amount, status, created_at FROM payments "
            "ORDER BY created_at DESC LIMIT 10"
        )
    lines = [
        "👑 <b>Админ-панель</b>\n",
        f"👥 Пользователей всего: <b>{total_users}</b>",
        f"💎 Платящих: <b>{paid_users}</b>",
        f"✅ Активных подписок: <b>{active_subs}</b>",
        f"💰 Заработано: <b>{money(total_income)}</b>\n",
        f"📁 Проектов: <b>{projects_cnt}</b>",
        f"📋 Расходов: <b>{expenses_cnt}</b>\n",
        "🧾 <b>Последние платежи:</b>",
    ]
    for p in last_payments:
        dt = p["created_at"].strftime("%d.%m %H:%M") if p["created_at"] else ""
        status_icon = "✅" if p["status"] == "succeeded" else "⏳"
        lines.append(
            f"{status_icon} id{p['user_id']} — {money(p['amount'])} ({dt})"
        )
    await update.message.reply_text(
        "\n".join(lines),
        parse_mode=ParseMode.HTML,
    )


async def cmd_cancel(update, ctx):
    for k in ("state", "exp", "new_project", "budget_pid", "invite_pid",
              "settings_pid", "receipt_exp_id", "receipt_pid"):
        ctx.user_data.pop(k, None)
    await update.message.reply_text("Отменено.", reply_markup=kb_main())


async def on_photo(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    u = update.effective_user
    await upsert_user(u)
    state = ctx.user_data.get("state")
    if state != "receipt:wait_photo":
        return await update.message.reply_text(
            "Сначала нажмите «📎 Прикрепить чек» в проекте, потом отправьте фото."
        )
    exp_id = ctx.user_data.get("receipt_exp_id")
    pid = ctx.user_data.get("receipt_pid")
    if not exp_id:
        ctx.user_data.pop("state", None)
        return await update.message.reply_text("Что-то пошло не так. Начните заново.")
    photo = update.message.photo[-1]
    file_id = photo.file_id
    async with pool.acquire() as c:
        await c.execute(
            "UPDATE expenses SET receipt_file_id=$1 WHERE id=$2",
            file_id, exp_id,
        )
    ctx.user_data.pop("state", None)
    ctx.user_data.pop("receipt_exp_id", None)
    ctx.user_data.pop("receipt_pid", None)
    await update.message.reply_text("✅ Чек прикреплён.")
    return await send_project_msg(update.message, pid, edit=False, user_id=u.id)


async def on_text(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    u = update.effective_user
    await upsert_user(u)
    if not await check_access_or_paywall(update, ctx, u):
        return
    state = ctx.user_data.get("state")
    text = (update.message.text or "").strip()
    if not state:
        pid = ctx.user_data.get("current_project")
        if pid and await user_can_view_project(pid, u.id):
            parsed = parse_quick(text)
            if parsed:
                name, qty, price = parsed
                ctx.user_data["exp"] = {
                    "project_id": pid,
                    "name": name, "qty": str(qty), "price": str(price),
                }
                ctx.user_data["state"] = "exp:category"
                total = qty * price
                return await update.message.reply_text(
                    "Разобрал:\n"
                    f"• Название: {esc(name)}\n"
                    f"• Количество: {fmt_qty(qty)}\n"
                    f"• Цена: {money(price)}\n"
                    f"• Сумма: {money(total)}\n\n"
                    "Выберите категорию:",
                    parse_mode=ParseMode.HTML,
                    reply_markup=kb_categories(pid),
                )
        return await update.message.reply_text(
            "Используйте /start для меню.",
            reply_markup=kb_main(),
        )
    if text.lower() in ("/cancel", "отмена"):
        for k in ("state", "exp", "new_project", "budget_pid", "invite_pid",
                  "settings_pid", "receipt_exp_id", "receipt_pid"):
            ctx.user_data.pop(k, None)
        return await update.message.reply_text("Отменено.", reply_markup=kb_main())
    if state == "newp:name":
        if not text:
            return
        ctx.user_data["new_project"] = {"name": text}
        ctx.user_data["state"] = "newp:budget"
        return await update.message.reply_text(
            f"Проект: <b>{esc(text)}</b>\n\n"
            "Шаг 2 из 3. Отправьте бюджет числом. Например: 100000.",
            parse_mode=ParseMode.HTML,
        )
    if state == "newp:budget":
        val = parse_money(text)
        if val is None:
            return await update.message.reply_text("Не понял. Отправьте число, например 100000.")
        np = ctx.user_data.get("new_project") or {}
        np["budget"] = str(val)
        ctx.user_data["new_project"] = np
        ctx.user_data["state"] = "newp:type"
        return await update.message.reply_text(
            f"Проект: <b>{esc(np.get('name',''))}</b>\n"
            f"Бюджет: <b>{money(val)}</b>\n\n"
            "Шаг 3 из 3. Выберите тип проекта:",
            parse_mode=ParseMode.HTML,
            reply_markup=kb_project_type(),
        )
    if state == "invite:username":
        pid = ctx.user_data.get("invite_pid")
        if not pid:
            ctx.user_data.pop("state", None)
            return await update.message.reply_text("Что-то пошло не так.", reply_markup=kb_main())
        uname = text.lstrip("@").strip()
        if not uname:
            return await update.message.reply_text("Отправьте username, например: @vasya")
        async with pool.acquire() as c:
            target = await c.fetchrow(
                "SELECT user_id, first_name, username FROM users WHERE LOWER(username)=LOWER($1)",
                uname,
            )
        if not target:
            return await update.message.reply_text(
                f"Пользователь @{esc(uname)} не найден.\n"
                "Попросите его сначала написать боту /start.",
                parse_mode=ParseMode.HTML,
            )
        async with pool.acquire() as c:
            await c.execute(
                "INSERT INTO project_members(project_id, user_id, role) VALUES($1,$2,$3) "
                "ON CONFLICT (project_id, user_id) DO NOTHING",
                pid, target["user_id"], ROLE_MEMBER,
            )
        ctx.user_data.pop("state", None)
        ctx.user_data.pop("invite_pid", None)
        name = target["first_name"] or ("@" + (target["username"] or uname))
        await update.message.reply_text(
            f"✅ {esc(name)} добавлен(а) в проект.",
            parse_mode=ParseMode.HTML,
        )
        return await send_project_msg(update.message, pid, edit=False, user_id=u.id)
    if state == "exp:name":
        exp = ctx.user_data.get("exp") or {}
        parsed = parse_quick(text)
        if parsed:
            name, qty, price = parsed
            exp.update({"name": name, "qty": str(qty), "price": str(price)})
            ctx.user_data["exp"] = exp
            ctx.user_data["state"] = "exp:category"
            return await update.message.reply_text(
                "Разобрал:\n"
                f"• Название: {esc(name)}\n"
                f"• Количество: {fmt_qty(qty)}\n"
                f"• Цена: {money(price)}\n"
                f"• Сумма: {money(qty*price)}\n\n"
                "Выберите категорию:",
                parse_mode=ParseMode.HTML,
                reply_markup=kb_categories(exp["project_id"]),
            )
        exp["name"] = text
        ctx.user_data["exp"] = exp
        ctx.user_data["state"] = "exp:qty"
        return await update.message.reply_text(
            f"Название: <b>{esc(text)}</b>\n\nСколько? (например 1, 3, 2.5)",
            parse_mode=ParseMode.HTML,
        )
    if state == "exp:qty":
        val = parse_number(text)
        if val is None:
            return await update.message.reply_text("Нужно число. Например: 1")
        exp = ctx.user_data["exp"]
        exp["qty"] = str(val)
        ctx.user_data["state"] = "exp:price"
        return await update.message.reply_text("Цена за единицу? (например 850)")
    if state == "exp:price":
        val = parse_money(text)
        if val is None:
            return await update.message.reply_text("Нужно число. Например: 850")
        exp = ctx.user_data["exp"]
        exp["price"] = str(val)
        ctx.user_data["state"] = "exp:category"
        total = Decimal(exp["qty"]) * Decimal(exp["price"])
        return await update.message.reply_text(
            f"Итого: <b>{money(total)}</b>\n\nВыберите категорию:",
            parse_mode=ParseMode.HTML,
            reply_markup=kb_categories(exp["project_id"]),
        )
    if state == "exp:comment":
        exp = ctx.user_data["exp"]
        exp["comment"] = text
        ctx.user_data["state"] = "exp:confirm"
        return await update.message.reply_text(
            _exp_summary(exp),
            parse_mode=ParseMode.HTML,
            reply_markup=kb_confirm(exp["project_id"]),
        )
    if state == "project:budget":
        val = parse_money(text)
        if val is None:
            return await update.message.reply_text("Нужно число.")
        pid = ctx.user_data.get("budget_pid")
        if not pid or not await user_can_edit_project(pid, u.id):
            ctx.user_data.pop("state", None)
            ctx.user_data.pop("budget_pid", None)
            return await update.message.reply_text("Нет прав на изменение бюджета.")
        async with pool.acquire() as c:
            await c.execute("UPDATE projects SET budget=$1 WHERE id=$2", val, pid)
        ctx.user_data.pop("state", None)
        ctx.user_data.pop("budget_pid", None)
        await update.message.reply_text("✅ Бюджет обновлён.")
        return await send_project_msg(update.message, pid, edit=False, user_id=u.id)
    await update.message.reply_text("Не понял. /cancel — отмена.", reply_markup=kb_main())


async def callbacks(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    data = q.data or ""
    u = update.effective_user
    await upsert_user(u)
    if not (data.startswith("buy:") or data == "menu"):
        if not await check_access_or_paywall(update, ctx, u):
            return
    if data == "menu":
        return await q.edit_message_text("Меню:", reply_markup=kb_main())
    if data == "projects:personal":
        text, kb = await projects_list_view(u.id, "personal")
        return await q.edit_message_text(text, parse_mode=ParseMode.HTML, reply_markup=kb)
    if data == "projects:team":
        text, kb = await projects_list_view(u.id, "team")
        return await q.edit_message_text(text, parse_mode=ParseMode.HTML, reply_markup=kb)
    if data == "project:new":
        ctx.user_data["state"] = "newp:name"
        ctx.user_data.pop("new_project", None)
        return await q.edit_message_text(
            "➕ <b>Новый проект</b>\n\n"
            "Шаг 1 из 3. Отправьте название проекта.\n\n"
            "/cancel — отмена",
            parse_mode=ParseMode.HTML,
            reply_markup=kb_back("menu"),
        )
    if data.startswith("newp:type:"):
        if ctx.user_data.get("state") != "newp:type":
            return await q.edit_message_text("Сессия истекла. Начните заново.")
        t = data.split(":")[2]
        is_personal = (t == "personal")
        np = ctx.user_data.get("new_project") or {}
        name = np.get("name", "Без названия")
        budget = Decimal(np.get("budget", "0"))
        async with pool.acquire() as c:
            row = await c.fetchrow(
                "INSERT INTO projects(name, budget, creator_id, is_personal) "
                "VALUES($1,$2,$3,$4) RETURNING id",
                name, budget, u.id, is_personal,
            )
            pid = row["id"]
            await c.execute(
                "INSERT INTO project_members(project_id, user_id, role) VALUES($1,$2,$3) "
                "ON CONFLICT (project_id, user_id) DO NOTHING",
                pid, u.id, ROLE_OWNER,
            )
        ctx.user_data.pop("state", None)
        ctx.user_data.pop("new_project", None)
        ctx.user_data["current_project"] = pid
        type_label = "👤 Личный" if is_personal else "👥 Командный"
        await q.edit_message_text(
            f"✅ Проект «{esc(name)}» создан.\nТип: {type_label}",
            parse_mode=ParseMode.HTML,
        )
        return await send_project_msg(q.message, pid, edit=False, user_id=u.id)
    if data.startswith("project:view:"):
        try:
            pid = int(data.split(":")[2])
        except (IndexError, ValueError):
            return await q.edit_message_text("Ошибка: не понял ID проекта.")
        if not await user_can_view_project(pid, u.id):
            return await q.edit_message_text("У вас нет доступа к этому проекту.")
        ctx.user_data["current_project"] = pid
        return await send_project_msg(q, pid, edit=True, user_id=u.id)
    if data.startswith("project:settings:"):
        pid = int(data.split(":")[2])
        if not await user_can_view_project(pid, u.id):
            return await q.edit_message_text("У вас нет доступа к этому проекту.")
        async with pool.acquire() as c:
            p = await c.fetchrow("SELECT * FROM projects WHERE id=$1", pid)
            owner = await c.fetchrow(
                "SELECT username, first_name FROM users WHERE user_id=$1",
                p["creator_id"] if p else 0,
            ) if p else None
        if not p:
            return await q.edit_message_text("Проект не найден.")
        owner_name = (
            ("@" + owner["username"]) if owner and owner["username"]
            else (owner["first_name"] if owner and owner["first_name"] else f"id{p['creator_id']}")
        )
        created = p["created_at"].strftime("%d.%m.%Y") if p["created_at"] else ""
        type_label = "👤 Личный" if p["is_personal"] else "👥 Командный"
        text = (
            "⚙️ <b>Настройки проекта</b>\n\n"
            f"📁 {esc(p['name'])}\n"
            f"Тип: {type_label}\n"
            f"💰 Бюджет: {money(p['budget'])}\n"
            f"👑 Создатель: {esc(owner_name)}\n"
            f"📅 Создан: {created}"
        )
        is_owner = await user_can_edit_project(pid, u.id)
        if is_owner:
            kb = InlineKeyboardMarkup([
                [InlineKeyboardButton("✏️ Изменить бюджет", callback_data=f"project:budget:{pid}")],
                [InlineKeyboardButton("🔄 Изменить тип",     callback_data=f"project:type:{pid}")],
                [InlineKeyboardButton("👥 Пригласить",       callback_data=f"project:invite:{pid}")],
                [InlineKeyboardButton("◀️ Назад",            callback_data=f"project:view:{pid}")],
            ])
        else:
            kb = InlineKeyboardMarkup([
                [InlineKeyboardButton("◀️ Назад", callback_data=f"project:view:{pid}")],
            ])
        return await q.edit_message_text(text, parse_mode=ParseMode.HTML, reply_markup=kb)
    if data.startswith("project:budget:"):
        pid = int(data.split(":")[2])
        if not await user_can_edit_project(pid, u.id):
            return await q.edit_message_text("Нет прав на изменение бюджета.")
        ctx.user_data["state"] = "project:budget"
        ctx.user_data["budget_pid"] = pid
        return await q.edit_message_text(
            "Отправьте новый бюджет числом (например 150000).\n\n/cancel — отмена."
        )
    if data.startswith("project:type:"):
        pid = int(data.split(":")[2])
        if not await user_can_edit_project(pid, u.id):
            return await q.edit_message_text("Нет прав.")
        async with pool.acquire() as c:
            p = await c.fetchrow("SELECT is_personal, name FROM projects WHERE id=$1", pid)
        if not p:
            return await q.edit_message_text("Проект не найден.")
        cur = p["is_personal"]
        cur_label = "👤 Личный" if cur else "👥 Командный"
        new_label = "👥 Командный" if cur else "👤 Личный"
        warn = ""
        if not cur:
            warn = "\n\n⚠️ <b>Внимание:</b> приглашённые потеряют доступ. Их расходы сохранятся."
        text = (
            f"Проект: <b>{esc(p['name'])}</b>\n"
            f"Сейчас: {cur_label}\n\n"
            f"Сменить на: {new_label}?{warn}"
        )
        return await q.edit_message_text(text, parse_mode=ParseMode.HTML,
                                         reply_markup=kb_change_type_menu(pid, cur))
    if data.startswith("project:type:team:") or data.startswith("project:type:personal:"):
        parts = data.split(":")
        new_type = parts[2]
        pid = int(parts[3])
        if not await user_can_edit_project(pid, u.id):
            return await q.edit_message_text("Нет прав.")
        is_personal = (new_type == "personal")
        async with pool.acquire() as c:
            await c.execute("UPDATE projects SET is_personal=$1 WHERE id=$2", is_personal, pid)
        label = "👤 Личный" if is_personal else "👥 Командный"
        await q.edit_message_text(f"✅ Тип проекта изменён на {label}.")
        return await send_project_msg(q.message, pid, edit=False, user_id=u.id)
    if data.startswith("project:invite:"):
        parts = data.split(":")
        if len(parts) == 3:
            pid = int(parts[2])
            if not await user_can_edit_project(pid, u.id):
                return await q.edit_message_text("Приглашать может только создатель.")
            async with pool.acquire() as c:
                p = await c.fetchrow("SELECT is_personal FROM projects WHERE id=$1", pid)
            if not p:
                return await q.edit_message_text("Проект не найден.")
            if p["is_personal"]:
                return await q.edit_message_text(
                    "Это личный проект. Сначала смените тип на «👥 Командный»."
                )
            return await q.edit_message_text(
                "👥 <b>Пригласить в проект</b>\n\nВыберите способ:",
                parse_mode=ParseMode.HTML,
                reply_markup=kb_invite_menu(pid),
            )
        if len(parts) == 4 and parts[2] == "user":
            pid = int(parts[3])
            if not await user_can_edit_project(pid, u.id):
                return await q.edit_message_text("Приглашать может только создатель.")
            ctx.user_data["state"] = "invite:username"
            ctx.user_data["invite_pid"] = pid
            return await q.edit_message_text(
                "Отправьте username, например: <code>@vasya</code>\n\n"
                "⚠️ Пользователь должен хотя бы раз написать боту /start.\n\n"
                "/cancel — отмена.",
                parse_mode=ParseMode.HTML,
            )
        if len(parts) == 4 and parts[2] == "link":
            pid = int(parts[3])
            if not await user_can_edit_project(pid, u.id):
                return await q.edit_message_text("Приглашать может только создатель.")
            async with pool.acquire() as c:
                p = await c.fetchrow("SELECT is_personal, name FROM projects WHERE id=$1", pid)
            if not p:
                return await q.edit_message_text("Проект не найден.")
            if p["is_personal"]:
                return await q.edit_message_text(
                    "Это личный проект. Сначала смените тип на «👥 Командный»."
                )
            code = await get_or_create_invite_code(pid)
            link = invite_link(code)
            return await q.edit_message_text(
                f"🔗 <b>Ссылка-приглашение</b>\n\n"
                f"<code>{esc(link)}</code>\n\n"
                "Отправьте её другу. Он нажмёт, и сразу попадёт в проект.",
                parse_mode=ParseMode.HTML,
                reply_markup=InlineKeyboardMarkup([
                    [InlineKeyboardButton("◀️ Назад", callback_data=f"project:invite:{pid}")],
                ]),
            )
    if data.startswith("project:delete:"):
        pid = int(data.split(":")[2])
        if not await user_can_edit_project(pid, u.id):
            return await q.edit_message_text("Удалять проект может только создатель.")
        kb = InlineKeyboardMarkup([
            [InlineKeyboardButton("🗑 Да, удалить", callback_data=f"project:delete_yes:{pid}")],
            [InlineKeyboardButton("◀️ Отмена",      callback_data=f"project:view:{pid}")],
        ])
        return await q.edit_message_text(
            "Удалить проект и все его расходы? Действие необратимо.",
            reply_markup=kb,
        )
    if data.startswith("project:delete_yes:"):
        pid = int(data.split(":")[2])
        if not await user_can_edit_project(pid, u.id):
            return await q.edit_message_text("Удалять проект может только создатель.")
        async with pool.acquire() as c:
            await c.execute("DELETE FROM projects WHERE id=$1", pid)
        ctx.user_data.pop("current_project", None)
        text, kb = await projects_list_view(u.id, "team")
        return await q.edit_message_text(text, parse_mode=ParseMode.HTML, reply_markup=kb)
    if data.startswith("exp:add:"):
        pid = int(data.split(":")[2])
        if not await user_can_view_project(pid, u.id):
            return await q.edit_message_text("У вас нет доступа к этому проекту.")
        ctx.user_data["state"] = "exp:name"
        ctx.user_data["exp"] = {"project_id": pid}
        ctx.user_data["current_project"] = pid
        return await q.edit_message_text(
            "➕ <b>Новый расход</b>\n\n"
            "Напишите одной строкой:\n"
            "• <code>такси 1200</code>\n"
            "• <code>3 банки краски по 850</code>\n"
            "• <code>свет 2 x 500</code>\n\n"
            "Или отправьте название — введу по шагам.\n\n"
            "/cancel — отмена",
            parse_mode=ParseMode.HTML,
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("❌ Отмена", callback_data=f"exp:cancel:{pid}")],
            ]),
        )
    if data.startswith("exp:list:"):
        pid = int(data.split(":")[2])
        if not await user_can_view_project(pid, u.id):
            return await q.edit_message_text("У вас нет доступа к этому проекту.")
        async with pool.acquire() as c:
            p = await c.fetchrow("SELECT * FROM projects WHERE id=$1", pid)
            rows = await c.fetch(
                "SELECT * FROM expenses WHERE project_id=$1 ORDER BY id DESC LIMIT 100",
                pid,
            )
            check_rows = await c.fetch(
                "SELECT id, name FROM expenses WHERE project_id=$1 AND receipt_file_id IS NOT NULL ORDER BY id DESC LIMIT 5",
                pid,
            )
        if not p:
            return await q.edit_message_text("Проект не найден.")
        if not rows:
            body = "Расходов пока нет."
        else:
            rows = _sort_rows_by_category(rows)
            blocks = []
            for r in rows:
                qv = Decimal(str(r["qty"]))
                pv = Decimal(str(r["price"]))
                s = qv * pv
                author = (
                    ("@" + r["author_username"]) if r["author_username"]
                    else (f"id{r['author_id']}" if r["author_id"] else "—")
                )
                dt = r["created_at"].strftime("%d.%m.%Y %H:%M") if r["created_at"] else ""
                block = [
                    f"{esc(r['category'])}",
                    f"<b>{esc(r['name'])}</b>",
                    f"{fmt_qty(qv)} × {money(pv)} = {money(s)}",
                ]
                if r["comment"]:
                    block.append(f"💬 {esc(r['comment'])}")
                if r["receipt_file_id"]:
                    block.append("📎 Чек прикреплён")
                block.append(f"👤 {esc(author)} · {dt}")
                blocks.append("\n".join(block))
            body = "\n\n".join(blocks)
        text = f"📋 <b>Расходы проекта «{esc(p['name'])}»</b>\n\n{body}"
        if len(text) > 4000:
            text = text[:3900] + "\n\n… (показаны не все)"
        kb_rows = []
        for r in check_rows:
            kb_rows.append([
                InlineKeyboardButton(
                    f"📎 Чек: {r['name'][:25]}",
                    callback_data=f"receipt:show:{r['id']}",
                )
            ])
        kb_rows.append([InlineKeyboardButton("➕ Добавить расход", callback_data=f"exp:add:{pid}")])
        kb_rows.append([InlineKeyboardButton("◀️ Назад к проекту", callback_data=f"project:view:{pid}")])
        kb = InlineKeyboardMarkup(kb_rows)
        return await q.edit_message_text(text, parse_mode=ParseMode.HTML, reply_markup=kb)
    if data.startswith("cat:"):
        _, pid_s, idx_s = data.split(":")
        pid = int(pid_s)
        cat = CATEGORIES[int(idx_s)]
        exp = ctx.user_data.get("exp") or {"project_id": pid}
        exp["category"] = cat
        ctx.user_data["exp"] = exp
        ctx.user_data["state"] = "exp:comment"
        return await q.edit_message_text(
            f"Категория: {esc(cat)}\n\nКомментарий? Отправьте текст или нажмите «Пропустить».",
            parse_mode=ParseMode.HTML,
            reply_markup=kb_comment(pid),
        )
    if data.startswith("exp:skip_comment:"):
        pid = int(data.split(":")[2])
        exp = ctx.user_data.get("exp") or {"project_id": pid}
        exp.setdefault("category", "📦 Другое")
        exp["comment"] = None
        ctx.user_data["exp"] = exp
        ctx.user_data["state"] = "exp:confirm"
        return await q.edit_message_text(
            _exp_summary(exp),
            parse_mode=ParseMode.HTML,
            reply_markup=kb_confirm(pid),
        )
    if data.startswith("exp:save:"):
        pid = int(data.split(":")[2])
        if not await user_can_view_project(pid, u.id):
            return await q.edit_message_text("У вас нет доступа к этому проекту.")
        exp = ctx.user_data.get("exp") or {}
        name = exp.get("name")
        if not name:
            return await q.edit_message_text("Что-то пошло не так. Начните заново.")
        qty = Decimal(str(exp.get("qty", "1")))
        price = Decimal(str(exp.get("price", "0")))
        cat = exp.get("category", "📦 Другое")
        comment = exp.get("comment")
        async with pool.acquire() as c:
            await c.execute(
                "INSERT INTO expenses(project_id, category, name, qty, price, "
                "comment, author_id, author_username) VALUES($1,$2,$3,$4,$5,$6,$7,$8)",
                pid, cat, name, qty, price, comment, u.id, u.username,
            )
        ctx.user_data.pop("state", None)
        ctx.user_data.pop("exp", None)
        ctx.user_data["current_project"] = pid
        text, kb = await project_view(pid, user_id=u.id)
        return await q.edit_message_text(
            "✅ Расход добавлен.\n\n" + text,
            parse_mode=ParseMode.HTML,
            reply_markup=kb,
        )
    if data.startswith("exp:cancel:"):
        pid = int(data.split(":")[2])
        ctx.user_data.pop("state", None)
        ctx.user_data.pop("exp", None)
        if not await user_can_view_project(pid, u.id):
            return await q.edit_message_text("Отменено.", reply_markup=kb_main())
        text, kb = await project_view(pid, user_id=u.id)
        return await q.edit_message_text(
            "Отменено.\n\n" + text,
            parse_mode=ParseMode.HTML,
            reply_markup=kb,
        )
    if data.startswith("receipt:choose:"):
        pid = int(data.split(":")[2])
        if not await user_can_view_project(pid, u.id):
            return await q.edit_message_text("У вас нет доступа к этому проекту.")
        async with pool.acquire() as c:
            rows = await c.fetch(
                "SELECT id, name, category, qty, price FROM expenses "
                "WHERE project_id=$1 ORDER BY id DESC LIMIT 20",
                pid,
            )
        if not rows:
            return await q.edit_message_text(
                "В проекте пока нет расходов.",
                reply_markup=InlineKeyboardMarkup([
                    [InlineKeyboardButton("◀️ Назад", callback_data=f"project:view:{pid}")],
                ]),
            )
        kb_rows = []
        for r in rows:
            s = Decimal(str(r["qty"])) * Decimal(str(r["price"]))
            label = f"{r['category'][:2]} {r['name'][:25]} — {money(s)}"
            kb_rows.append([
                InlineKeyboardButton(label, callback_data=f"receipt:pick:{r['id']}")
            ])
        kb_rows.append([InlineKeyboardButton("◀️ Назад", callback_data=f"project:view:{pid}")])
        return await q.edit_message_text(
            "📎 <b>Выберите расход, к которому прикрепить чек:</b>",
            parse_mode=ParseMode.HTML,
            reply_markup=InlineKeyboardMarkup(kb_rows),
        )
    if data.startswith("receipt:pick:"):
        exp_id = int(data.split(":")[2])
        async with pool.acquire() as c:
            exp = await c.fetchrow(
                "SELECT id, project_id, name FROM expenses WHERE id=$1",
                exp_id,
            )
        if not exp:
            return await q.edit_message_text("Расход не найден.")
        pid = exp["project_id"]
        if not await user_can_view_project(pid, u.id):
            return await q.edit_message_text("Нет доступа к этому проекту.")
        ctx.user_data["state"] = "receipt:wait_photo"
        ctx.user_data["receipt_exp_id"] = exp_id
        ctx.user_data["receipt_pid"] = pid
        return await q.edit_message_text(
            f"📎 <b>Отправьте фото чека</b>\n\n"
            f"Расход: {esc(exp['name'])}\n\n"
            "Отправьте фото как <b>фотографию</b> (не как файл).\n\n"
            "/cancel — отмена.",
            parse_mode=ParseMode.HTML,
        )
    if data.startswith("receipt:show:"):
        exp_id = int(data.split(":")[2])
        async with pool.acquire() as c:
            exp = await c.fetchrow(
                "SELECT id, project_id, name, receipt_file_id FROM expenses WHERE id=$1",
                exp_id,
            )
        if not exp or not exp["receipt_file_id"]:
            return await q.edit_message_text("Чек не найден.")
        if not await user_can_view_project(exp["project_id"], u.id):
            return await q.edit_message_text("Нет доступа.")
        try:
            await q.message.reply_photo(
                exp["receipt_file_id"],
                caption=f"📎 Чек: {esc(exp['name'])}",
                parse_mode=ParseMode.HTML,
                reply_markup=InlineKeyboardMarkup([
                    [InlineKeyboardButton("❌ Открепить чек", callback_data=f"receipt:detach:{exp_id}")],
                ]),
            )
        except Exception as e:
            log.error("Send receipt failed: %s", e)
            await q.message.reply_text("Не удалось загрузить фото.")
        return

    if data.startswith("receipt:detach:"):
        exp_id = int(data.split(":")[2])
        async with pool.acquire() as c:
            exp = await c.fetchrow(
                "SELECT id, project_id FROM expenses WHERE id=$1", exp_id
            )
        if not exp:
            return await q.edit_message_text("Расход не найден.")
        if not await user_can_view_project(exp["project_id"], u.id):
            return await q.edit_message_text("Нет доступа.")
        async with pool.acquire() as c:
            await c.execute(
                "UPDATE expenses SET receipt_file_id=NULL WHERE id=$1", exp_id
            )
        await q.edit_message_text("✅ Чек откреплён.")
        return
    if data.startswith("excel:"):
        try:
            pid = int(data.split(":")[1])
        except (IndexError, ValueError):
            return await q.edit_message_text("Ошибка: не понял, какой проект.")
        return await make_excel(update, pid)
    if data.startswith("pdf:"):
        try:
            pid = int(data.split(":")[1])
        except (IndexError, ValueError):
            return await q.edit_message_text("Ошибка: не понял, какой проект.")
        return await make_pdf(update, pid)
    if data == "buy:access":
        is_active, expires = await get_subscription_info(u.id)
        header = ""
        if is_active:
            exp_str = expires.strftime("%d.%m.%Y") if expires else "—"
            header = f"✅ <b>Ваша подписка активна</b>\nДействует до: <b>{exp_str}</b>\n\n"
        return await q.edit_message_text(
            header +
            "💎 <b>Подписка на бота</b>\n\n"
            "Выберите тариф:",
            parse_mode=ParseMode.HTML,
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("📅 30 дней — 500 ₽", callback_data="buy:pay:30")],
                [InlineKeyboardButton("🔥 90 дней — 1300 ₽ (выгодно)", callback_data="buy:pay:90")],
                [InlineKeyboardButton("◀️ Назад", callback_data="menu")],
            ]),
        )
    if data.startswith("buy:pay"):
        parts = data.split(":")
        days = int(parts[2]) if len(parts) > 2 else 30
        amount = 500 if days == 30 else 1300
        await q.edit_message_text("⏳ Создаю платёж...")
        link, err = await create_yookassa_payment(
            u.id, amount, description=f"Подписка на {days} дней"
        )
        if not link:
            return await q.edit_message_text(
                f"❌ {err}",
                reply_markup=InlineKeyboardMarkup([
                    [InlineKeyboardButton("◀️ Назад", callback_data="menu")],
                ]),
            )
        return await q.edit_message_text(
            f"💎 <b>Оплата подписки на {days} дней</b>\n\n"
            f"Сумма: <b>{money(amount)}</b>\n\n"
            "Нажмите кнопку ниже, чтобы перейти к оплате.\n"
            "После оплаты вернитесь и нажмите «Я оплатил».",
            parse_mode=ParseMode.HTML,
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("🔗 Перейти к оплате", url=link)],
                [InlineKeyboardButton("✅ Я оплатил", callback_data="buy:check")],
                [InlineKeyboardButton("◀️ Назад", callback_data="menu")],
            ]),
        )
    if data == "buy:check":
        async with pool.acquire() as c:
            row = await c.fetchrow(
                "SELECT payment_id FROM payments WHERE user_id=$1 AND status='pending' "
                "ORDER BY created_at DESC LIMIT 1",
                u.id,
            )
        if not row:
            return await q.edit_message_text(
                "❌ Не нашли активный платёж. Попробуйте ещё раз.",
                reply_markup=InlineKeyboardMarkup([
                    [InlineKeyboardButton("◀️ Назад", callback_data="menu")],
                ]),
            )
        ok, msg = await check_yookassa_payment(u.id, row["payment_id"])
        return await q.edit_message_text(
            f"{'✅' if ok else '⏳'} {msg}",
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("◀️ В меню", callback_data="menu")],
            ]),
        )
    log.warning("Unhandled callback: %s", data)


async def send_expiry_notifications(app):
    async with pool.acquire() as c:
        soon = await c.fetch(
            "SELECT user_id, expires_at FROM subscriptions "
            "WHERE expires_at > NOW() AND expires_at < NOW() + INTERVAL '3 days' "
            "AND notified_3days = FALSE"
        )
    for r in soon:
        try:
            await app.bot.send_message(
                r["user_id"],
                f"⏰ Ваша подписка заканчивается {r['expires_at'].strftime('%d.%m.%Y')}.\n\n"
                "Продлите, чтобы не потерять доступ — напишите /start",
            )
            async with pool.acquire() as c:
                await c.execute(
                    "UPDATE subscriptions SET notified_3days=TRUE WHERE user_id=$1",
                    r["user_id"],
                )
        except Exception as e:
            log.warning("Notify 3days failed for %s: %s", r["user_id"], e)
    async with pool.acquire() as c:
        expired = await c.fetch(
            "SELECT user_id FROM subscriptions "
            "WHERE expires_at < NOW() AND notified_expired = FALSE"
        )
    for r in expired:
        try:
            await app.bot.send_message(
                r["user_id"],
                "🔒 Ваша подписка закончилась.\n\n"
                "Чтобы продолжить пользоваться ботом — напишите /start",
            )
            async with pool.acquire() as c:
                await c.execute(
                    "UPDATE subscriptions SET notified_expired=TRUE WHERE user_id=$1",
                    r["user_id"],
                )
        except Exception as e:
            log.warning("Notify expired failed for %s: %s", r["user_id"], e)


async def expiry_notifications_loop(app):
    await asyncio.sleep(90)
    while True:
        try:
            await send_expiry_notifications(app)
        except Exception as e:
            log.error("Notify loop failed: %s", e)
        await asyncio.sleep(6 * 3600)


async def post_init(app):
    global bot_username
    await init_db()
    me = await app.bot.get_me()
    bot_username = me.username
    log.info("Bot @%s started", me.username)
    asyncio.create_task(expiry_notifications_loop(app))


def main():
    app = (
        Application.builder()
        .token(BOT_TOKEN)
        .post_init(post_init)
        .build()
    )
    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("menu", cmd_menu))
    app.add_handler(CommandHandler("admin", cmd_admin))
    app.add_handler(CommandHandler("cancel", cmd_cancel))
    app.add_handler(CallbackQueryHandler(callbacks))
    app.add_handler(MessageHandler(filters.PHOTO, on_photo))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, on_text))
    app.run_polling()


if __name__ == "__main__":
    main()
