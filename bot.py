from dotenv import load_dotenv
load_dotenv()

import os
import io
import re
import html
import logging
import secrets
import string
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

from async_yookassa import YooKassaClient
from async_yookassa.models.payment import PaymentRequest, Amount, RedirectConfirmationRequest
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
)
log = logging.getLogger("smeta")

BOT_TOKEN = os.environ["BOT_TOKEN"]
DATABASE_URL = os.environ["DATABASE_URL"]
YOOKASSA_SHOP_ID = os.environ.get("1481007", "")
YOOKASSA_SECRET_KEY = os.environ.get("test_7o5ITQKUQZRSgRLYKmiQp1mlEdMEfirT5bWQlitaz9A", "")
YOOKASSA_RETURN_URL = os.environ.get("https://t.me@smetafashion_bot", "https://t.me")
pool = None
bot_username = None

# Белый список Telegram ID (кому разрешён доступ)
ALLOWED_USER_IDS = []

CATEGORIES = [
    "🎬 Аренда",
    "📍 Локация",
    "💡 Свет",
    "🎥 Камера",
    "🎤 Звук",
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
    id              SERIAL PRIMARY KEY,
    project_id      INTEGER NOT NULL REFERENCES projects(id) ON DELETE CASCADE,
    category        TEXT NOT NULL DEFAULT '📦 Другое',
    name            TEXT NOT NULL,
    qty             NUMERIC(14,3) NOT NULL DEFAULT 1,
    price           NUMERIC(14,2) NOT NULL DEFAULT 0,
    comment         TEXT,
    author_id       BIGINT,
    author_username TEXT,
    created_at      TIMESTAMPTZ NOT NULL DEFAULT NOW()
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
    user_id       BIGINT PRIMARY KEY,
    expires_at    TIMESTAMPTZ NOT NULL,
    updated_at    TIMESTAMPTZ NOT NULL DEFAULT NOW()
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
        await c.execute("UPDATE projects SET is_personal=FALSE WHERE is_personal IS NULL")
        await c.execute(
            """
            INSERT INTO project_members(project_id, user_id, role)
            SELECT id, creator_id, 'owner'
            FROM projects
            WHERE creator_id IS NOT NULL
            ON CONFLICT (project_id, user_id) DO NOTHING
            """
        )
        await c.execute(
            """
            INSERT INTO project_members(project_id, user_id, role)
            SELECT DISTINCT e.project_id, e.author_id, 'member'
            FROM expenses e
            WHERE e.author_id IS NOT NULL
            ON CONFLICT (project_id, user_id) DO NOTHING
            """
        )
        await c.execute("UPDATE expenses SET category='💄 Макияж' WHERE category='💄 Грим'")
    log.info("Migration done")


async def init_db():
    global pool
    pool = await asyncpg.create_pool(DATABASE_URL, min_size=1, max_size=5)
    async with pool.acquire() as c:
        await c.execute(SCHEMA)
    await migrate_db()
    log.info("DB schema ready")



# ---------- ЮKassa ----------

async def create_yookassa_payment(user_id, amount, description="Доступ к боту"):
    """Создаёт платёж и возвращает ссылку на оплату."""
    if not YOOKASSA_SHOP_ID or not YOOKASSA_SECRET_KEY:
        return None, "Платежи временно недоступны."

    try:
        async with YooKassaClient(
            account_id=YOOKASSA_SHOP_ID,
            secret_key=YOOKASSA_SECRET_KEY,
        ) as client:
            request = PaymentRequest(
                amount=Amount(value=f"{amount:.2f}", currency="RUB"),
                confirmation=RedirectConfirmationRequest(
                    type="redirect",
                    return_url=YOOKASSA_RETURN_URL,
                ),
                description=description,
                capture=True,
                metadata={"user_id": str(user_id)},
            )
            payment = await client.payment.create(request)

            # сохраняем в БД
            async with pool.acquire() as c:
                await c.execute(
                    "INSERT INTO payments(user_id, payment_id, amount, status) "
                    "VALUES($1,$2,$3,$4)",
                    user_id, payment.id, amount, payment.status,
                )

            return payment.confirmation.confirmation_url, None
    except Exception as e:
        log.error("YooKassa create payment failed: %s", e)
        return None, "Не удалось создать платёж. Попробуйте позже."


async def check_yookassa_payment(user_id, payment_id):
    """Проверяет статус платежа и выдаёт подписку на 30 дней, если оплачен."""
    if not YOOKASSA_SHOP_ID or not YOOKASSA_SECRET_KEY:
        return False, "Платежи недоступны."

    try:
        async with YooKassaClient(
            account_id=YOOKASSA_SHOP_ID,
            secret_key=YOOKASSA_SECRET_KEY,
        ) as client:
            payment = await client.payment.get(payment_id)

            if payment.status == "succeeded":
                # обновляем статус в БД
                async with pool.acquire() as c:
                    await c.execute(
                        "UPDATE payments SET status='succeeded' WHERE payment_id=$1",
                        payment_id,
                    )
                    # продлеваем подписку на 30 дней от текущей даты или от конца текущей
                    await c.execute(
                        """
                        INSERT INTO subscriptions(user_id, expires_at, updated_at)
                        VALUES($1, NOW() + INTERVAL '30 days', NOW())
                        ON CONFLICT (user_id) DO UPDATE
                        SET expires_at = GREATEST(subscriptions.expires_at, NOW()) + INTERVAL '30 days',
                            updated_at = NOW()
                        """,
                        user_id,
                    )
                return True, "Оплата прошла! Доступ активирован на 30 дней."
            else:
                return False, f"Статус платежа: {payment.status}. Попробуйте позже."

    except Exception as e:
        log.error("YooKassa check payment failed: %s", e)
        return False, "Не удалось проверить платёж."


async def user_has_subscription(user_id):
    """Проверяет, активна ли подписка."""
    async with pool.acquire() as c:
        row = await c.fetchrow(
            "SELECT expires_at FROM subscriptions WHERE user_id=$1 AND expires_at > NOW()",
            user_id,
        )
    return row is not None
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
        row = await c.fetchrow("SELECT code FROM invites WHERE project_id=$1", project_id)
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

# ---------- Клавиатуры ----------

def kb_main():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("👤 Личные проекты",   callback_data="projects:personal")],
        [InlineKeyboardButton("👥 Командные проекты", callback_data="projects:team")],
        [InlineKeyboardButton("➕ Новый проект",      callback_data="project:new")],
        [InlineKeyboardButton("💎 Купить доступ",    callback_data="buy:access")],
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
    ]
    if role == ROLE_OWNER:
        rows.append([
            InlineKeyboardButton("👥 Пригласить", callback_data=f"project:invite:{pid}"),
            InlineKeyboardButton("⚙️ Настройки",  callback_data=f"project:settings:{pid}"),
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
        [InlineKeyboardButton("📝 По @username", callback_data=f"project:invite:user:{pid}")],
        [InlineKeyboardButton("🔗 Ссылка",       callback_data=f"project:invite:link:{pid}")],
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
            [InlineKeyboardButton("👤 Сделать личным", callback_data=f"project:type:personal:{pid}")],
            [InlineKeyboardButton("◀️ Назад",          callback_data=f"project:settings:{pid}")],
        ])


# ---------- Отображение ----------

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


async def show_delete_choices(target, ctx, pid, user_id, query=None):
    async with pool.acquire() as c:
        if query:
            rows = await c.fetch(
                """
                SELECT id, name, qty, price, category
                FROM expenses
                WHERE project_id=$1 AND LOWER(name) LIKE $2
                ORDER BY id DESC
                LIMIT 25
                """,
                pid, f"%{query.lower()}%",
            )
        else:
            rows = await c.fetch(
                """
                SELECT id, name, qty, price, category
                FROM expenses
                WHERE project_id=$1
                ORDER BY id DESC
                LIMIT 25
                """,
                pid,
            )
    if not rows:
        return await target.edit_message_text(
            "Ничего не найдено." if query else "Расходов нет.",
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("◀️ К расходам", callback_data=f"exp:list:{pid}")],
            ]),
        )
    kb_rows = []
    for r in rows:
        qv = Decimal(str(r["qty"]))
        pv = Decimal(str(r["price"]))
        s = qv * pv
        label = f"{r['category']} {r['name']} · {money(s)}"[:60]
        kb_rows.append([InlineKeyboardButton(label, callback_data=f"exp:delete_pick:{r['id']}")])
    kb_rows.append([InlineKeyboardButton("❌ Отмена", callback_data=f"exp:list:{pid}")])
    header = f"Найдено: {len(rows)}" if query else f"Всего расходов: {len(rows)}. Выберите, что удалить:"
    return await target.edit_message_text(header, reply_markup=InlineKeyboardMarkup(kb_rows))


async def show_edit_choices(target, pid, user_id, query=None):
    async with pool.acquire() as c:
        if query:
            rows = await c.fetch(
                """
                SELECT id, name, qty, price, category
                FROM expenses
                WHERE project_id=$1 AND LOWER(name) LIKE $2
                ORDER BY id DESC
                LIMIT 25
                """,
                pid, f"%{query.lower()}%",
            )
        else:
            rows = await c.fetch(
                """
                SELECT id, name, qty, price, category
                FROM expenses
                WHERE project_id=$1
                ORDER BY id DESC
                LIMIT 25
                """,
                pid,
            )
    if not rows:
        return await target.edit_message_text(
            "Ничего не найдено." if query else "Расходов нет.",
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("◀️ К расходам", callback_data=f"exp:list:{pid}")],
            ]),
        )
    kb_rows = []
    for r in rows:
        qv = Decimal(str(r["qty"]))
        pv = Decimal(str(r["price"]))
        s = qv * pv
        label = f"{r['category']} {r['name']} · {money(s)}"[:60]
        kb_rows.append([InlineKeyboardButton(label, callback_data=f"exp:edit_pick:{r['id']}")])
    kb_rows.append([InlineKeyboardButton("❌ Отмена", callback_data=f"exp:list:{pid}")])
    header = f"Найдено: {len(rows)}" if query else f"Всего расходов: {len(rows)}. Выберите, что редактировать:"
    return await target.edit_message_text(header, reply_markup=InlineKeyboardMarkup(kb_rows))


async def finish_edit(update, ctx):
    edit = ctx.user_data.get("edit_exp")
    if not edit:
        ctx.user_data.pop("state", None)
        if update.callback_query:
            return await update.callback_query.edit_message_text("Что-то пошло не так.")
        return await update.message.reply_text("Что-то пошло не так.", reply_markup=kb_main())
    async with pool.acquire() as c:
        await c.execute(
            """
            UPDATE expenses
            SET name=$1, qty=$2, price=$3, category=$4, comment=$5
            WHERE id=$6
            """,
            edit["name"], Decimal(edit["qty"]), Decimal(edit["price"]),
            edit["category"], edit.get("comment"), edit["id"],
        )
    pid = edit["project_id"]
    uid = update.effective_user.id
    ctx.user_data.pop("state", None)
    ctx.user_data.pop("edit_exp", None)
    if update.callback_query:
        await update.callback_query.edit_message_text("✅ Расход обновлён.")
        text, kb = await project_view(pid, user_id=uid)
        return await update.callback_query.message.reply_text(text, parse_mode=ParseMode.HTML, reply_markup=kb)
    await update.message.reply_text("✅ Расход обновлён.")
    text, kb = await project_view(pid, user_id=uid)
    return await update.message.reply_text(text, parse_mode=ParseMode.HTML, reply_markup=kb)

# ---------- Excel ----------

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
    ws.merge_cells("A1:H1")
    ws["A2"] = "Проект";    ws["B2"] = p["name"]
    ws["A3"] = "Тип";       ws["B3"] = "Личный" if p["is_personal"] else "Командный"
    ws["A4"] = "Бюджет";    ws["B4"] = float(budget)
    ws["A5"] = "Потрачено"; ws["B5"] = float(total)
    ws["A6"] = "Осталось";  ws["B6"] = float(left)

    headers = ["Категория", "Наименование", "Количество", "Цена", "Сумма", "Комментарий", "Кто добавил", "Дата"]
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
        author = ("@" + r["author_username"]) if r["author_username"] else (f"id{r['author_id']}" if r["author_id"] else "")
        ws.append([r["category"], r["name"], float(qv), float(pv), float(s), r["comment"] or "", author, dt])

    total_row = ws.max_row + 1
    ws.cell(row=total_row, column=4, value="ИТОГО").font = bold
    ws.cell(row=total_row, column=5, value=float(total)).font = bold
    for i, w in enumerate([22, 28, 12, 12, 14, 30, 18, 18], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    await q.message.reply_document(buf, filename=f"smeta_{_safe_name(p['name'])}.xlsx")


# ---------- PDF ----------

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
    doc = SimpleDocTemplate(out, pagesize=A4, leftMargin=30, rightMargin=30, topMargin=30, bottomMargin=30)
    title_style = ParagraphStyle("T",  fontName=FONT_BOLD, fontSize=22, leading=26, alignment=1)
    h2_style    = ParagraphStyle("H2", fontName=FONT_BOLD, fontSize=14, leading=18, spaceAfter=6)
    body_style  = ParagraphStyle("B",  fontName=FONT,      fontSize=11, leading=15)
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

    data = [["Категория", "Наименование", "Кол-во", "Цена", "Сумма"]]
    for r in rows:
        qv = Decimal(str(r["qty"]))
        pv = Decimal(str(r["price"]))
        s = qv * pv
        data.append([r["category"], r["name"], fmt_qty(qv), money(pv), money(s)])
    data.append(["", "", "", "ИТОГО", money(total)])

    table = Table(data, repeatRows=1, colWidths=[95, 170, 55, 75, 85])
    table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), .4, colors.grey),
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("FONTNAME", (0, 0), (-1, 0), FONT_BOLD),
        ("FONTNAME", (0, 1), (-1, -2), FONT),
        ("FONTNAME", (0, -1), (-1, -1), FONT_BOLD),
        ("ALIGN", (2, 1), (-1, -1), "RIGHT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    story.append(table)
    doc.build(story)
    out.seek(0)
    await q.message.reply_document(out, filename=f"smeta_{_safe_name(p['name'])}.pdf")


# ---------- Команды ----------

async def cmd_start(update, ctx):
    u = update.effective_user
  
    await upsert_user(u)
    args = ctx.args or []
    if args and args[0].startswith("join_"):
        return await handle_join_link(update, ctx, args[0][5:])
    text = (
        f"👋 Привет, {esc(u.first_name or 'друг')}!\n\n"
        "Это бот для <b>командной</b> работы со сметами.\n"
        "Создавайте проекты — личные или командные. "
        "Приглашайте друзей по @username или ссылке.\n\n"
        "Выберите раздел:"
    )
    await update.message.reply_text(text, parse_mode=ParseMode.HTML, reply_markup=kb_main())


async def handle_join_link(update, ctx, code):
    u = update.effective_user
    async with pool.acquire() as c:
        row = await c.fetchrow("SELECT project_id FROM invites WHERE code=$1", code)
    if not row:
        return await update.message.reply_text("Ссылка недействительна.", reply_markup=kb_main())
    pid = row["project_id"]
    async with pool.acquire() as c:
        p = await c.fetchrow("SELECT id, name, is_personal FROM projects WHERE id=$1", pid)
    if not p:
        return await update.message.reply_text("Проект больше не существует.", reply_markup=kb_main())
    if p["is_personal"]:
        return await update.message.reply_text("Этот проект сейчас личный.", reply_markup=kb_main())
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
    await update.message.reply_text("Меню:", reply_markup=kb_main())


async def cmd_cancel(update, ctx):
    for k in ("state", "exp", "new_project", "budget_pid", "invite_pid", "delete_pid", "edit_pid", "edit_exp", "search_pid"):
        ctx.user_data.pop(k, None)
    await update.message.reply_text("Отменено.", reply_markup=kb_main())

async def on_text(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    u = update.effective_user
  
    await upsert_user(u)
    state = ctx.user_data.get("state")
    text = (update.message.text or "").strip()

    if not state:
        pid = ctx.user_data.get("current_project")
        if pid and await user_can_view_project(pid, u.id):
            parsed = parse_quick(text)
            if parsed:
                name, qty, price = parsed
                ctx.user_data["exp"] = {"project_id": pid, "name": name, "qty": str(qty), "price": str(price)}
                ctx.user_data["state"] = "exp:category"
                total = qty * price
                return await update.message.reply_text(
                    "Разобрал:\n"
                    f"• Название: {esc(name)}\n"
                    f"• Количество: {fmt_qty(qty)}\n"
                    f"• Цена: {money(price)}\n"
                    f"• Сумма: {money(total)}\n\nВыберите категорию:",
                    parse_mode=ParseMode.HTML,
                    reply_markup=kb_categories(pid),
                )
        return await update.message.reply_text("Используйте /start для меню.", reply_markup=kb_main())

    if text.lower() in ("/cancel", "отмена"):
        for k in ("state", "exp", "new_project", "budget_pid", "invite_pid", "delete_pid", "edit_pid", "edit_exp", "search_pid"):
            ctx.user_data.pop(k, None)
        return await update.message.reply_text("Отменено.", reply_markup=kb_main())

    # --- создание проекта ---
    if state == "newp:name":
        if not text:
            return
        ctx.user_data["new_project"] = {"name": text}
        ctx.user_data["state"] = "newp:budget"
        return await update.message.reply_text(
            f"Проект: <b>{esc(text)}</b>\n\nОтправьте бюджет числом (₽). Например: 100000.\nЕсли бюджет неизвестен — отправьте 0.",
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
            f"Проект: <b>{esc(np.get('name',''))}</b>\nБюджет: <b>{money(val)}</b>\n\nВыберите тип проекта:",
            parse_mode=ParseMode.HTML,
            reply_markup=kb_project_type(),
        )

    # --- поиск для удаления ---
    if state == "delete:search":
        pid = ctx.user_data.get("delete_pid")
        if not pid:
            ctx.user_data.pop("state", None)
            return await update.message.reply_text("Что-то пошло не так.", reply_markup=kb_main())
        word = text.strip()
        if not word:
            return await update.message.reply_text("Введите слово, например: такси")
        ctx.user_data.pop("state", None)
        ctx.user_data.pop("delete_pid", None)
        async with pool.acquire() as c:
            rows = await c.fetch(
                "SELECT id, name, qty, price, category FROM expenses WHERE project_id=$1 AND LOWER(name) LIKE $2 ORDER BY id DESC LIMIT 25",
                pid, f"%{word.lower()}%",
            )
        if not rows:
            return await update.message.reply_text(f"По слову «{esc(word)}» ничего не найдено.", parse_mode=ParseMode.HTML)
        kb_rows = []
        for r in rows:
            qv = Decimal(str(r["qty"])); pv = Decimal(str(r["price"])); s = qv * pv
            label = f"{r['category']} {r['name']} · {money(s)}"[:60]
            kb_rows.append([InlineKeyboardButton(label, callback_data=f"exp:delete_pick:{r['id']}")])
        kb_rows.append([InlineKeyboardButton("❌ Отмена", callback_data=f"exp:list:{pid}")])
        return await update.message.reply_text(f"Найдено: {len(rows)}. Выберите, что удалить:", reply_markup=InlineKeyboardMarkup(kb_rows))

    # --- поиск для редактирования ---
    if state == "edit:search":
        pid = ctx.user_data.get("edit_pid")
        if not pid:
            ctx.user_data.pop("state", None)
            return await update.message.reply_text("Что-то пошло не так.", reply_markup=kb_main())
        word = text.strip()
        if not word:
            return await update.message.reply_text("Введите слово, например: такси")
        ctx.user_data.pop("state", None)
        ctx.user_data.pop("edit_pid", None)
        async with pool.acquire() as c:
            rows = await c.fetch(
                "SELECT id, name, qty, price, category FROM expenses WHERE project_id=$1 AND LOWER(name) LIKE $2 ORDER BY id DESC LIMIT 25",
                pid, f"%{word.lower()}%",
            )
        if not rows:
            return await update.message.reply_text(f"По слову «{esc(word)}» ничего не найдено.", parse_mode=ParseMode.HTML)
        kb_rows = []
        for r in rows:
            qv = Decimal(str(r["qty"])); pv = Decimal(str(r["price"])); s = qv * pv
            label = f"{r['category']} {r['name']} · {money(s)}"[:60]
            kb_rows.append([InlineKeyboardButton(label, callback_data=f"exp:edit_pick:{r['id']}")])
        kb_rows.append([InlineKeyboardButton("❌ Отмена", callback_data=f"exp:list:{pid}")])
        return await update.message.reply_text(f"Найдено: {len(rows)}. Выберите, что редактировать:", reply_markup=InlineKeyboardMarkup(kb_rows))

    # --- ПОИСК (общий) ---
    if state == "search:query":
        pid = ctx.user_data.get("search_pid")
        if not pid:
            ctx.user_data.pop("state", None)
            return await update.message.reply_text("Что-то пошло не так.", reply_markup=kb_main())
        word = text.strip()
        if not word:
            return await update.message.reply_text("Введите слово для поиска.")
        ctx.user_data.pop("state", None)
        ctx.user_data["search_pid"] = pid
        return await _send_search_results_msg(update, pid, word)

    # --- редактирование ---
    if state == "edit:name":
        edit = ctx.user_data.get("edit_exp")
        if not edit:
            ctx.user_data.pop("state", None)
            return await update.message.reply_text("Что-то пошло не так.", reply_markup=kb_main())
        if text:
            edit["name"] = text.strip()
        ctx.user_data["edit_exp"] = edit
        ctx.user_data["state"] = "edit:qty"
        return await update.message.reply_text(
            "Шаг 2 из 5 — Количество.\n\n"
            f"Текущее: <b>{fmt_qty(Decimal(edit['qty']))}</b>\n\nОтправьте новое количество или нажмите «⏭ Оставить».\n\n/cancel — отмена.",
            parse_mode=ParseMode.HTML,
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("⏭ Оставить", callback_data="edit:keep:qty")],
                [InlineKeyboardButton("❌ Отмена",   callback_data=f"exp:list:{edit['project_id']}")],
            ]),
        )

    if state == "edit:qty":
        edit = ctx.user_data.get("edit_exp")
        if not edit:
            ctx.user_data.pop("state", None)
            return await update.message.reply_text("Что-то пошло не так.", reply_markup=kb_main())
        val = parse_number(text)
        if val is None:
            return await update.message.reply_text("Нужно число. Например: 1")
        edit["qty"] = str(val)
        ctx.user_data["edit_exp"] = edit
        ctx.user_data["state"] = "edit:price"
        return await update.message.reply_text(
            "Шаг 3 из 5 — Цена за единицу.\n\n"
            f"Текущая: <b>{money(edit['price'])}</b>\n\nОтправьте новую цену или нажмите «⏭ Оставить».\n\n/cancel — отмена.",
            parse_mode=ParseMode.HTML,
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("⏭ Оставить", callback_data="edit:keep:price")],
                [InlineKeyboardButton("❌ Отмена",   callback_data=f"exp:list:{edit['project_id']}")],
            ]),
        )

    if state == "edit:price":
        edit = ctx.user_data.get("edit_exp")
        if not edit:
            ctx.user_data.pop("state", None)
            return await update.message.reply_text("Что-то пошло не так.", reply_markup=kb_main())
        val = parse_money(text)
        if val is None:
            return await update.message.reply_text("Нужно число. Например: 850")
        edit["price"] = str(val)
        ctx.user_data["edit_exp"] = edit
        ctx.user_data["state"] = "edit:category"
        return await update.message.reply_text(
            "Шаг 4 из 5 — Категория.\n\n"
            f"Текущая: {esc(edit['category'])}\n\nВыберите новую категорию или нажмите «⏭ Оставить».",
            parse_mode=ParseMode.HTML,
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("🏷 Выбрать категорию", callback_data="edit:pick_category")],
                [InlineKeyboardButton("⏭ Оставить",           callback_data="edit:keep:cat")],
                [InlineKeyboardButton("❌ Отмена",             callback_data=f"exp:list:{edit['project_id']}")],
            ]),
        )

    if state == "edit:comment":
        edit = ctx.user_data.get("edit_exp")
        if not edit:
            ctx.user_data.pop("state", None)
            return await update.message.reply_text("Что-то пошло не так.", reply_markup=kb_main())
        edit["comment"] = text.strip() if text.strip() else None
        ctx.user_data["edit_exp"] = edit
        return await finish_edit(update, ctx)

    # --- приглашение ---
    if state == "invite:username":
        pid = ctx.user_data.get("invite_pid")
        if not pid:
            ctx.user_data.pop("state", None)
            return await update.message.reply_text("Что-то пошло не так.", reply_markup=kb_main())
        uname = text.lstrip("@").strip()
        if not uname:
            return await update.message.reply_text("Отправьте username, например: @vasya")
        async with pool.acquire() as c:
            target = await c.fetchrow("SELECT user_id, first_name, username FROM users WHERE LOWER(username)=LOWER($1)", uname)
        if not target:
            return await update.message.reply_text(
                f"Пользователь @{esc(uname)} не найден. Попросите его написать боту /start.",
                parse_mode=ParseMode.HTML,
            )
        async with pool.acquire() as c:
            await c.execute(
                "INSERT INTO project_members(project_id, user_id, role) VALUES($1,$2,$3) ON CONFLICT (project_id, user_id) DO NOTHING",
                pid, target["user_id"], ROLE_MEMBER,
            )
        ctx.user_data.pop("state", None)
        ctx.user_data.pop("invite_pid", None)
        name = target["first_name"] or ("@" + (target["username"] or uname))
        await update.message.reply_text(f"✅ {esc(name)} добавлен(а) в проект.", parse_mode=ParseMode.HTML)
        return await send_project_msg(update.message, pid, edit=False, user_id=u.id)

    # --- расход ---
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
                f"• Сумма: {money(qty*price)}\n\nВыберите категорию:",
                parse_mode=ParseMode.HTML,
                reply_markup=kb_categories(exp["project_id"]),
            )
        exp["name"] = text
        ctx.user_data["exp"] = exp
        ctx.user_data["state"] = "exp:qty"
        return await update.message.reply_text(f"Название: <b>{esc(text)}</b>\n\nСколько? (например 1, 3, 2.5)", parse_mode=ParseMode.HTML)

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
        return await update.message.reply_text(_exp_summary(exp), parse_mode=ParseMode.HTML, reply_markup=kb_confirm(exp["project_id"]))

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


async def _send_search_results_msg(update, pid, query):
    async with pool.acquire() as c:
        rows = await c.fetch(
            """
            SELECT id, name, qty, price, category, comment, author_username, author_id, created_at
            FROM expenses
            WHERE project_id=$1 AND (LOWER(name) LIKE $2 OR LOWER(category) LIKE $2 OR LOWER(COALESCE(comment,'')) LIKE $2)
            ORDER BY id DESC
            LIMIT 25
            """,
            pid, f"%{query.lower()}%",
        )
    if not rows:
        return await update.message.reply_text(
            f"По запросу «{esc(query)}» ничего не найдено.",
            parse_mode=ParseMode.HTML,
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("🔍 Искать ещё",  callback_data=f"exp:search:{pid}")],
                [InlineKeyboardButton("◀️ К расходам", callback_data=f"exp:list:{pid}")],
            ]),
        )
    kb_rows = []
    for r in rows:
        qv = Decimal(str(r["qty"])); pv = Decimal(str(r["price"])); s = qv * pv
        label = f"{r['category']} {r['name']} · {money(s)}"[:60]
        kb_rows.append([InlineKeyboardButton(label, callback_data=f"exp:search_pick:{r['id']}")])
    kb_rows.append([InlineKeyboardButton("🔍 Искать ещё",  callback_data=f"exp:search:{pid}")])
    kb_rows.append([InlineKeyboardButton("◀️ К расходам", callback_data=f"exp:list:{pid}")])
    return await update.message.reply_text(
        f"🔍 Найдено по «{esc(query)}»: {len(rows)}",
        parse_mode=ParseMode.HTML,
        reply_markup=InlineKeyboardMarkup(kb_rows),
    )


async def callbacks(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    data = q.data or ""
    u = update.effective_user
   
    await upsert_user(u)

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
            "➕ <b>Новый проект</b>\n\nШаг 1 из 3. Отправьте название проекта.\n\n/cancel — отмена",
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
                "INSERT INTO projects(name, budget, creator_id, is_personal) VALUES($1,$2,$3,$4) RETURNING id",
                name, budget, u.id, is_personal,
            )
            pid = row["id"]
            await c.execute(
                "INSERT INTO project_members(project_id, user_id, role) VALUES($1,$2,$3) ON CONFLICT (project_id, user_id) DO NOTHING",
                pid, u.id, ROLE_OWNER,
            )
        ctx.user_data.pop("state", None)
        ctx.user_data.pop("new_project", None)
        ctx.user_data["current_project"] = pid
        type_label = "👤 Личный" if is_personal else "👥 Командный"
        await q.edit_message_text(f"✅ Проект «{esc(name)}» создан.\nТип: {type_label}", parse_mode=ParseMode.HTML)
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
            owner = await c.fetchrow("SELECT username, first_name FROM users WHERE user_id=$1", p["creator_id"] if p else 0) if p else None
        if not p:
            return await q.edit_message_text("Проект не найден.")
        owner_name = ("@" + owner["username"]) if owner and owner["username"] else (owner["first_name"] if owner and owner["first_name"] else f"id{p['creator_id']}")
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
                [InlineKeyboardButton("🔄 Изменить тип",   callback_data=f"project:type:{pid}")],
                [InlineKeyboardButton("👥 Пригласить",     callback_data=f"project:invite:{pid}")],
                [InlineKeyboardButton("◀️ Назад к проекту", callback_data=f"project:view:{pid}")],
            ])
        else:
            kb = InlineKeyboardMarkup([
                [InlineKeyboardButton("◀️ Назад к проекту", callback_data=f"project:view:{pid}")],
            ])
        return await q.edit_message_text(text, parse_mode=ParseMode.HTML, reply_markup=kb)

    if data.startswith("project:budget:"):
        pid = int(data.split(":")[2])
        if not await user_can_edit_project(pid, u.id):
            return await q.edit_message_text("Нет прав на изменение бюджета.")
        ctx.user_data["state"] = "project:budget"
        ctx.user_data["budget_pid"] = pid
        return await q.edit_message_text("Отправьте новый бюджет числом (например 150000).\n\n/cancel — отмена.")

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
        warn = "\n\n⚠️ <b>Внимание:</b> все приглашённые потеряют доступ к проекту. Их расходы сохранятся." if not cur else ""
        text = f"Проект: <b>{esc(p['name'])}</b>\nСейчас: {cur_label}\n\nСменить на: {new_label}?{warn}"
        return await q.edit_message_text(text, parse_mode=ParseMode.HTML, reply_markup=kb_change_type_menu(pid, cur))

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
                return await q.edit_message_text("Приглашать может только создатель проекта.")
            async with pool.acquire() as c:
                p = await c.fetchrow("SELECT is_personal FROM projects WHERE id=$1", pid)
            if not p:
                return await q.edit_message_text("Проект не найден.")
            if p["is_personal"]:
                return await q.edit_message_text("Это личный проект. Сначала смените тип на «👥 Командный».")
            return await q.edit_message_text("👥 <b>Пригласить в проект</b>\n\nВыберите способ:", parse_mode=ParseMode.HTML, reply_markup=kb_invite_menu(pid))

        if len(parts) == 4 and parts[2] == "user":
            pid = int(parts[3])
            if not await user_can_edit_project(pid, u.id):
                return await q.edit_message_text("Приглашать может только создатель.")
            ctx.user_data["state"] = "invite:username"
            ctx.user_data["invite_pid"] = pid
            return await q.edit_message_text(
                "Отправьте username пользователя, например: <code>@vasya</code>\n\n⚠️ Пользователь должен хотя бы раз написать боту /start.\n\n/cancel — отмена.",
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
                return await q.edit_message_text("Это личный проект. Сначала смените тип на «👥 Командный».")
            code = await get_or_create_invite_code(pid)
            link = invite_link(code)
            return await q.edit_message_text(
                f"🔗 <b>Ссылка-приглашение</b>\n\n<code>{esc(link)}</code>\n\nОтправьте её другу.",
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
        return await q.edit_message_text("Удалить проект и все его расходы? Действие необратимо.", reply_markup=kb)

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
            "➕ <b>Новый расход</b>\n\nНапишите одной строкой, например:\n• <code>такси 1200</code>\n• <code>3 банки краски по 850</code>\n• <code>свет 2 x 500</code>\n\nИли отправьте только название.\n\n/cancel — отмена",
            parse_mode=ParseMode.HTML,
            reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("❌ Отмена", callback_data=f"exp:cancel:{pid}")]]),
        )

    if data.startswith("exp:list:"):
        pid = int(data.split(":")[2])
        if not await user_can_view_project(pid, u.id):
            return await q.edit_message_text("У вас нет доступа к этому проекту.")
        async with pool.acquire() as c:
            p = await c.fetchrow("SELECT * FROM projects WHERE id=$1", pid)
            rows = await c.fetch("SELECT * FROM expenses WHERE project_id=$1 ORDER BY id DESC LIMIT 100", pid)
        if not p:
            return await q.edit_message_text("Проект не найден.")
        if not rows:
            body = "Расходов пока нет."
        else:
            rows = _sort_rows_by_category(rows)
            blocks = []
            for r in rows:
                qv = Decimal(str(r["qty"])); pv = Decimal(str(r["price"])); s = qv * pv
                author = ("@" + r["author_username"]) if r["author_username"] else (f"id{r['author_id']}" if r["author_id"] else "—")
                dt = r["created_at"].strftime("%d.%m.%Y %H:%M") if r["created_at"] else ""
                block = [f"{esc(r['category'])}", f"<b>{esc(r['name'])}</b>", f"{fmt_qty(qv)} × {money(pv)} = {money(s)}"]
                if r["comment"]:
                    block.append(f"💬 {esc(r['comment'])}")
                block.append(f"👤 {esc(author)} · {dt}")
                blocks.append("\n".join(block))
            body = "\n\n".join(blocks)
        text = f"📋 <b>Расходы проекта «{esc(p['name'])}»</b>\n\n{body}"
        if len(text) > 4000:
            text = text[:3900] + "\n\n… (показаны не все)"
        kb = InlineKeyboardMarkup([
            [InlineKeyboardButton("➕ Добавить расход", callback_data=f"exp:add:{pid}")],
            [InlineKeyboardButton("✏️ Редактировать",   callback_data=f"exp:edit:{pid}"),
             InlineKeyboardButton("🗑 Удалить расход",  callback_data=f"exp:delete:{pid}")],
            [InlineKeyboardButton("🔍 Поиск",           callback_data=f"exp:search:{pid}")],
            [InlineKeyboardButton("◀️ Назад к проекту", callback_data=f"project:view:{pid}")],
        ])
        return await q.edit_message_text(text, parse_mode=ParseMode.HTML, reply_markup=kb)

    # --- ПОИСК ---
    if data.startswith("exp:search:"):
        pid = int(data.split(":")[2])
        if not await user_can_view_project(pid, u.id):
            return await q.edit_message_text("У вас нет доступа к этому проекту.")
        ctx.user_data["state"] = "search:query"
        ctx.user_data["search_pid"] = pid
        return await q.edit_message_text(
            "🔍 <b>Поиск по расходам</b>\n\nВведите слово или часть названия, категории или комментария.\nНапример: <code>такси</code>\n\n/cancel — отмена.",
            parse_mode=ParseMode.HTML,
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("◀️ К расходам", callback_data=f"exp:list:{pid}")],
            ]),
        )

    if data.startswith("exp:search_pick:"):
        eid = int(data.split(":")[2])
        async with pool.acquire() as c:
            e = await c.fetchrow(
                "SELECT id, project_id, name, qty, price, category, comment, author_username, author_id, created_at FROM expenses WHERE id=$1",
                eid,
            )
        if not e:
            return await q.edit_message_text("Расход уже удалён.")
        if not await user_can_view_project(e["project_id"], u.id):
            return await q.edit_message_text("У вас нет доступа к этому проекту.")
        qv = Decimal(str(e["qty"])); pv = Decimal(str(e["price"])); s = qv * pv
        author = ("@" + e["author_username"]) if e["author_username"] else (f"id{e['author_id']}" if e["author_id"] else "—")
        dt = e["created_at"].strftime("%d.%m.%Y %H:%M") if e["created_at"] else ""
        text = (
            f"🔍 <b>Расход</b>\n\n"
            f"• Категория: {esc(e['category'])}\n"
            f"• Название: <b>{esc(e['name'])}</b>\n"
            f"• Количество: {fmt_qty(qv)}\n"
            f"• Цена: {money(pv)}\n"
            f"• Сумма: <b>{money(s)}</b>\n"
        )
        if e["comment"]:
            text += f"• Комментарий: {esc(e['comment'])}\n"
        text += f"• Добавил: {esc(author)}\n• Дата: {dt}"
        kb = InlineKeyboardMarkup([
            [InlineKeyboardButton("✏️ Редактировать", callback_data=f"exp:edit_pick:{eid}"),
             InlineKeyboardButton("🗑 Удалить",       callback_data=f"exp:delete_pick:{eid}")],
            [InlineKeyboardButton("🔍 Искать ещё",    callback_data=f"exp:search:{e['project_id']}")],
            [InlineKeyboardButton("◀️ К расходам",   callback_data=f"exp:list:{e['project_id']}")],
        ])
        return await q.edit_message_text(text, parse_mode=ParseMode.HTML, reply_markup=kb)

    # --- РЕДАКТИРОВАНИЕ ---
    if data.startswith("exp:edit:"):
        pid = int(data.split(":")[2])
        if not await user_can_view_project(pid, u.id):
            return await q.edit_message_text("У вас нет доступа к этому проекту.")
        async with pool.acquire() as c:
            cnt = await c.fetchval("SELECT COUNT(*) FROM expenses WHERE project_id=$1", pid)
        if cnt == 0:
            return await q.edit_message_text("Расходов нет.", reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("◀️ К расходам", callback_data=f"exp:list:{pid}")]]))
        if cnt > 25:
            ctx.user_data["state"] = "edit:search"
            ctx.user_data["edit_pid"] = pid
            return await q.edit_message_text(f"В проекте {cnt} расходов. Введите часть названия для поиска.\n\n/cancel — отмена.")
        return await show_edit_choices(q, pid, u.id)

    if data.startswith("exp:edit_pick:"):
        eid = int(data.split(":")[2])
        async with pool.acquire() as c:
            e = await c.fetchrow(
                "SELECT id, project_id, name, qty, price, category, comment FROM expenses WHERE id=$1",
                eid,
            )
        if not e:
            return await q.edit_message_text("Расход уже удалён.")
        if not await user_can_view_project(e["project_id"], u.id):
            return await q.edit_message_text("У вас нет доступа к этому проекту.")
        ctx.user_data["edit_exp"] = {
            "id": e["id"], "project_id": e["project_id"], "name": e["name"],
            "qty": str(e["qty"]), "price": str(e["price"]), "category": e["category"], "comment": e["comment"],
        }
        ctx.user_data["state"] = "edit:name"
        return await q.edit_message_text(
            f"Редактируем расход «{esc(e['name'])}»\n\nШаг 1 из 5 — Название.\n\n"
            f"Текущее: <b>{esc(e['name'])}</b>\n\nОтправьте новое название или нажмите «⏭ Оставить».\n\n/cancel — отмена.",
            parse_mode=ParseMode.HTML,
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("⏭ Оставить", callback_data="edit:keep:name")],
                [InlineKeyboardButton("❌ Отмена",   callback_data=f"exp:list:{e['project_id']}")],
            ]),
        )

    if data.startswith("edit:keep:"):
        what = data.split(":")[2]
        edit = ctx.user_data.get("edit_exp")
        if not edit:
            return await q.edit_message_text("Сессия истекла. Начните заново.")
        if what == "name":
            ctx.user_data["state"] = "edit:qty"
            return await q.edit_message_text(
                "Шаг 2 из 5 — Количество.\n\n"
                f"Текущее: <b>{fmt_qty(Decimal(edit['qty']))}</b>\n\nОтправьте новое количество или нажмите «⏭ Оставить».",
                parse_mode=ParseMode.HTML,
                reply_markup=InlineKeyboardMarkup([
                    [InlineKeyboardButton("⏭ Оставить", callback_data="edit:keep:qty")],
                    [InlineKeyboardButton("❌ Отмена",   callback_data=f"exp:list:{edit['project_id']}")],
                ]),
            )
        if what == "qty":
            ctx.user_data["state"] = "edit:price"
            return await q.edit_message_text(
                "Шаг 3 из 5 — Цена за единицу.\n\n"
                f"Текущая: <b>{money(edit['price'])}</b>\n\nОтправьте новую цену или нажмите «⏭ Оставить».",
                parse_mode=ParseMode.HTML,
                reply_markup=InlineKeyboardMarkup([
                    [InlineKeyboardButton("⏭ Оставить", callback_data="edit:keep:price")],
                    [InlineKeyboardButton("❌ Отмена",   callback_data=f"exp:list:{edit['project_id']}")],
                ]),
            )
        if what == "price":
            ctx.user_data["state"] = "edit:category"
            return await q.edit_message_text(
                "Шаг 4 из 5 — Категория.\n\n"
                f"Текущая: {esc(edit['category'])}\n\nВыберите новую категорию или нажмите «⏭ Оставить».",
                parse_mode=ParseMode.HTML,
                reply_markup=InlineKeyboardMarkup([
                    [InlineKeyboardButton("🏷 Выбрать категорию", callback_data="edit:pick_category")],
                    [InlineKeyboardButton("⏭ Оставить",           callback_data="edit:keep:cat")],
                    [InlineKeyboardButton("❌ Отмена",             callback_data=f"exp:list:{edit['project_id']}")],
                ]),
            )
        if what == "cat":
            ctx.user_data["state"] = "edit:comment"
            return await q.edit_message_text(
                "Шаг 5 из 5 — Комментарий.\n\n"
                f"Текущий: {esc(edit.get('comment') or '—')}\n\nОтправьте новый комментарий или нажмите «⏭ Оставить».",
                parse_mode=ParseMode.HTML,
                reply_markup=InlineKeyboardMarkup([
                    [InlineKeyboardButton("⏭ Оставить", callback_data="edit:keep:comment")],
                    [InlineKeyboardButton("❌ Отмена",   callback_data=f"exp:list:{edit['project_id']}")],
                ]),
            )
        if what == "comment":
            return await finish_edit(update, ctx)

    if data == "edit:pick_category":
        edit = ctx.user_data.get("edit_exp")
        if not edit:
            return await q.edit_message_text("Сессия истекла. Начните заново.")
        rows = [[InlineKeyboardButton(c, callback_data=f"edit:cat:{i}")] for i, c in enumerate(CATEGORIES)]
        rows.append([InlineKeyboardButton("⏭ Оставить", callback_data="edit:keep:cat")])
        return await q.edit_message_text("Выберите новую категорию:", reply_markup=InlineKeyboardMarkup(rows))

    if data.startswith("edit:cat:"):
        idx = int(data.split(":")[2])
        edit = ctx.user_data.get("edit_exp")
        if not edit:
            return await q.edit_message_text("Сессия истекла. Начните заново.")
        edit["category"] = CATEGORIES[idx]
        ctx.user_data["edit_exp"] = edit
        ctx.user_data["state"] = "edit:comment"
        return await q.edit_message_text(
            "Шаг 5 из 5 — Комментарий.\n\n"
            f"Текущий: {esc(edit.get('comment') or '—')}\n\nОтправьте новый комментарий или нажмите «⏭ Оставить».",
            parse_mode=ParseMode.HTML,
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("⏭ Оставить", callback_data="edit:keep:comment")],
                [InlineKeyboardButton("❌ Отмена",   callback_data=f"exp:list:{edit['project_id']}")],
            ]),
        )
    # --- покупка доступа ---
    if data == "buy:access":
        # проверяем, есть ли уже подписка
        if await user_has_subscription(u.id):
            async with pool.acquire() as c:
                row = await c.fetchrow(
                    "SELECT expires_at FROM subscriptions WHERE user_id=$1",
                    u.id,
                )
            exp = row["expires_at"].strftime("%d.%m.%Y") if row else "—"
            return await q.edit_message_text(
                f"💎 У вас уже есть доступ до <b>{exp}</b>.\n\n"
                "Продлить можно, нажав кнопку ниже.",
                parse_mode=ParseMode.HTML,
                reply_markup=InlineKeyboardMarkup([
                    [InlineKeyboardButton("🔄 Продлить на 30 дней — 500 ₽", callback_data="buy:pay")],
                    [InlineKeyboardButton("◀️ Назад", callback_data="menu")],
                ]),
            )
        return await q.edit_message_text(
            "💎 <b>Доступ к боту</b>\n\n"
            "Подписка на 30 дней — <b>500 ₽</b>.\n\n"
            "После оплаты доступ активируется автоматически.",
            parse_mode=ParseMode.HTML,
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("✅ Оплатить 500 ₽", callback_data="buy:pay")],
                [InlineKeyboardButton("◀️ Назад", callback_data="menu")],
            ]),
        )

    if data == "buy:pay":
        await q.edit_message_text("⏳ Создаю платёж...")
        link, err = await create_yookassa_payment(u.id, 500)
        if not link:
            return await q.edit_message_text(
                f"❌ {err}",
                reply_markup=InlineKeyboardMarkup([
                    [InlineKeyboardButton("◀️ Назад", callback_data="menu")],
                ]),
            )
        return await q.edit_message_text(
            "💎 <b>Оплата доступа</b>\n\n"
            "Нажмите кнопку ниже, чтобы перейти на страницу оплаты.\n"
            "После оплаты вернитесь в бота и нажмите «Я оплатил».",
            parse_mode=ParseMode.HTML,
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("🔗 Перейти к оплате", url=link)],
                [InlineKeyboardButton("✅ Я оплатил", callback_data="buy:check")],
                [InlineKeyboardButton("◀️ Назад", callback_data="menu")],
            ]),
        )

    if data == "buy:check":
        # ищем последний pending платёж пользователя
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
    # --- УДАЛЕНИЕ ---
    if data.startswith("exp:delete:"):
        pid = int(data.split(":")[2])
        if not await user_can_view_project(pid, u.id):
            return await q.edit_message_text("У вас нет доступа к этому проекту.")
        async with pool.acquire() as c:
            cnt = await c.fetchval("SELECT COUNT(*) FROM expenses WHERE project_id=$1", pid)
        if cnt == 0:
            return await q.edit_message_text("Расходов нет.", reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("◀️ К расходам", callback_data=f"exp:list:{pid}")]]))
        if cnt > 25:
            ctx.user_data["state"] = "delete:search"
            ctx.user_data["delete_pid"] = pid
            return await q.edit_message_text(f"В проекте {cnt} расходов. Введите часть названия для поиска.\n\n/cancel — отмена.")
        return await show_delete_choices(q, ctx, pid, u.id)

    if data.startswith("exp:delete_pick:"):
        eid = int(data.split(":")[2])
        async with pool.acquire() as c:
            e = await c.fetchrow("SELECT id, project_id, name, qty, price, category FROM expenses WHERE id=$1", eid)
        if not e:
            return await q.edit_message_text("Расход уже удалён.")
        if not await user_can_view_project(e["project_id"], u.id):
            return await q.edit_message_text("У вас нет доступа к этому проекту.")
        qv = Decimal(str(e["qty"])); pv = Decimal(str(e["price"])); s = qv * pv
        text = f"Удалить этот расход?\n\n{esc(e['category'])}\n<b>{esc(e['name'])}</b>\n{fmt_qty(qv)} × {money(pv)} = <b>{money(s)}</b>"
        kb = InlineKeyboardMarkup([
            [InlineKeyboardButton("🗑 Да, удалить", callback_data=f"exp:delete_yes:{eid}")],
            [InlineKeyboardButton("❌ Отмена",      callback_data=f"exp:list:{e['project_id']}")],
        ])
        return await q.edit_message_text(text, parse_mode=ParseMode.HTML, reply_markup=kb)

    if data.startswith("exp:delete_yes:"):
        eid = int(data.split(":")[2])
        async with pool.acquire() as c:
            e = await c.fetchrow("SELECT project_id FROM expenses WHERE id=$1", eid)
        if not e:
            return await q.edit_message_text("Расход уже удалён.")
        pid = e["project_id"]
        if not await user_can_view_project(pid, u.id):
            return await q.edit_message_text("У вас нет доступа к этому проекту.")
        async with pool.acquire() as c:
            await c.execute("DELETE FROM expenses WHERE id=$1", eid)
        await q.edit_message_text("✅ Расход удалён.")
        text, kb = await project_view(pid, user_id=u.id)
        return await q.message.reply_text(text, parse_mode=ParseMode.HTML, reply_markup=kb)

    # --- категории (новая трата) ---
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
        return await q.edit_message_text(_exp_summary(exp), parse_mode=ParseMode.HTML, reply_markup=kb_confirm(pid))

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
                "INSERT INTO expenses(project_id, category, name, qty, price, comment, author_id, author_username) VALUES($1,$2,$3,$4,$5,$6,$7,$8)",
                pid, cat, name, qty, price, comment, u.id, u.username,
            )
        ctx.user_data.pop("state", None)
        ctx.user_data.pop("exp", None)
        ctx.user_data["current_project"] = pid
        text, kb = await project_view(pid, user_id=u.id)
        return await q.edit_message_text("✅ Расход добавлен.\n\n" + text, parse_mode=ParseMode.HTML, reply_markup=kb)

    if data.startswith("exp:cancel:"):
        pid = int(data.split(":")[2])
        ctx.user_data.pop("state", None)
        ctx.user_data.pop("exp", None)
        if not await user_can_view_project(pid, u.id):
            return await q.edit_message_text("Отменено.", reply_markup=kb_main())
        text, kb = await project_view(pid, user_id=u.id)
        return await q.edit_message_text("Отменено.\n\n" + text, parse_mode=ParseMode.HTML, reply_markup=kb)

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

    log.warning("Unhandled callback: %s", data)


async def post_init(app):
    global bot_username
    await init_db()
    me = await app.bot.get_me()
    bot_username = me.username
    log.info("Bot @%s started", me.username)


def main():
    app = (
        Application.builder()
        .token(BOT_TOKEN)
        .post_init(post_init)
        .build()
    )
    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("menu", cmd_menu))
    app.add_handler(CommandHandler("cancel", cmd_cancel))
    app.add_handler(CallbackQueryHandler(callbacks))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, on_text))
    app.run_polling()


if __name__ == "__main__":
    main()
