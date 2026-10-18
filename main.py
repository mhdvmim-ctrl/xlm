"""
================================================================================
XLM ANONYMOUS DATING BOT - ULTIMATE MASTER EDITION v3.0
================================================================================
Ishlab chiquvchi: Senior Python Backend Developer & Bot Architect
Tavsif: O'zbek tilida ishlaydigan, 18+ anonim tanishuv va suhbat Telegram boti.
Platforma: Render (Background Worker 24/7) yoki Mahalliy server.
Kutubxonalar: aiogram 3.x, aiosqlite, openpyxl, python-dotenv, Pillow.
================================================================================
"""

import asyncio
import csv
import html
import io
import json
import logging
import os
import re
import signal
import sys
from datetime import datetime, timedelta
from typing import Any, Dict, List, Optional, Set, Tuple

import aiosqlite
from dotenv import load_dotenv
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

from aiogram import Bot, Dispatcher, F, Router
from aiogram.enums import ChatAction, ContentType, ParseMode
from aiogram.exceptions import (
    TelegramAPIError,
    TelegramBadRequest,
    TelegramForbiddenError,
    TelegramNotFound,
    TelegramRetryAfter,
)
from aiogram.filters import Command, CommandStart, StateFilter
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import (
    BufferedInputFile,
    CallbackQuery,
    ErrorEvent,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    InputMediaPhoto,
    KeyboardButton,
    Message,
    ReplyKeyboardMarkup,
    ReplyKeyboardRemove,
)

# ==============================================================================
# 1. KONFIGURATSIYA VA TIZIM JURNALI (LOGGING)
# ==============================================================================

env_file = ".env" if os.path.exists(".env") else (".env.example" if os.path.exists(".env.example") else None)
if env_file:
    load_dotenv(env_file, encoding="utf-8")
else:
    load_dotenv()

BOT_TOKEN: str = os.getenv("BOT_TOKEN", "").strip()
ADMIN_IDS_RAW: str = os.getenv("ADMIN_IDS", "").strip()
DATABASE_PATH: str = os.getenv("DATABASE_PATH", "dating.db").strip()
TIMEZONE: str = os.getenv("TIMEZONE", "Asia/Tashkent").strip()
LOG_LEVEL_NAME: str = os.getenv("LOG_LEVEL", "INFO").strip().upper()

ADMIN_IDS: Set[int] = set()
if ADMIN_IDS_RAW:
    for item in ADMIN_IDS_RAW.split(","):
        clean_item = item.strip()
        if clean_item.isdigit():
            ADMIN_IDS.add(int(clean_item))

db_dir = os.path.dirname(DATABASE_PATH)
if db_dir and not os.path.exists(db_dir):
    try:
        os.makedirs(db_dir, exist_ok=True)
    except Exception as e:
        print(f"[XATO] Database katalogini yaratib bo'lmadi: {e}", file=sys.stderr)

log_level = getattr(logging, LOG_LEVEL_NAME, logging.INFO)
logging.basicConfig(
    level=log_level,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
logger = logging.getLogger("xlm_dating_bot")

if not BOT_TOKEN and __name__ == "__main__":
    logger.critical(
        "DIQQAT! 'BOT_TOKEN' environment variable topilmadi yoki bo'sh! "
        "Iltimos, .env faylini yoki Render Environment sozlamalarini tekshiring."
    )
    sys.exit(1)

REGIONS: List[str] = [
    "Toshkent shahri",
    "Toshkent viloyati",
    "Andijon",
    "Buxoro",
    "Farg‘ona",
    "Jizzax",
    "Xorazm",
    "Namangan",
    "Navoiy",
    "Qashqadaryo",
    "Samarqand",
    "Sirdaryo",
    "Surxondaryo",
    "Qoraqalpog‘iston Respublikasi",
]

# 22 ta Admin ruxsatlari (Permissions)
ADMIN_PERMISSIONS: Dict[str, str] = {
    "can_view_users": "👥 Foydalanuvchilarni ko'rish",
    "can_view_usernames": "🔍 Username va kontaktni ko'rish",
    "can_ban_users": "🚫 Foydalanuvchini bloklash",
    "can_unban_users": "🔓 Blokdan chiqarish",
    "can_send_messages": "💬 Bot nomidan xabar yozish",
    "can_manage_premium": "⭐ Premium berish/bekor qilish",
    "can_manage_plans": "📦 Tariflar narxini boshqarish",
    "can_approve_payments": "✅ To'lovlarni tasdiqlash",
    "can_reject_payments": "❌ To'lovlarni rad etish",
    "can_view_payments": "💳 To'lovlar ro'yxatini ko'rish",
    "can_manage_channels": "📢 Majburiy obuna kanallari",
    "can_broadcast": "📣 Ommaviy xabar tarqatish",
    "can_view_stats": "📊 Statistikani ko'rish",
    "can_export_data": "📤 Ma'lumotlarni eksport qilish",
    "can_manage_reports": "🛡 Shikoyatlar va moderatsiya",
    "can_manage_settings": "⚙️ Bot sozlamalari (karta)",
    "can_toggle_maintenance": "🔄 Xizmat holatini o'zgartirish",
    "can_view_audit_logs": "🧾 Adminlar jurnali (audit)",
    "can_manage_admins": "👑 Adminlar boshqaruvi",
    "can_hide_profiles": "🙈 Anketalarni yashirish",
    "can_delete_profiles": "🗑 Anketalarni o'chirish",
    "can_view_chats": "💬 Chatlar auditi",
}

DEFAULT_SUBADMIN_PERMS = {
    "can_view_users": True,
    "can_view_usernames": False,
    "can_ban_users": False,
    "can_unban_users": False,
    "can_send_messages": False,
    "can_manage_premium": False,
    "can_manage_plans": False,
    "can_approve_payments": False,
    "can_reject_payments": False,
    "can_view_payments": True,
    "can_manage_channels": False,
    "can_broadcast": False,
    "can_view_stats": True,
    "can_export_data": False,
    "can_manage_reports": True,
    "can_manage_settings": False,
    "can_toggle_maintenance": False,
    "can_view_audit_logs": False,
    "can_manage_admins": False,
    "can_hide_profiles": False,
    "can_delete_profiles": False,
    "can_view_chats": False,
}

# Standart CMS matnlari (Admin panelidan tahrirlash mumkin!)
DEFAULT_SYSTEM_TEXTS: Dict[str, Dict[str, str]] = {
    "start_welcome": {
        "title": "👋 /start salomlashish va qoidalar",
        "default": (
            "✨ <b>XLM Anonymous Dating Botiga xush kelibsiz!</b>\n\n"
            "🔒 <b>Maxfiylik kafolati:</b> Bu yerda siz insonlar bilan 100% anonim tarzda tanishishingiz va suhbatlashishingiz mumkin. "
            "Shaxsiy profilingiz, username va telefon raqamingiz boshqalarga aslo ko'rsatilmaydi.\n\n"
            "💌 <b>Shaxsiy chat:</b> Eslatma! Shaxsiy anonim suhbatlar faqat pullik obunadagi (Premium) foydalanuvchilarga ochiladi va ko'rinadi.\n\n"
            "🔞 <b>Foydalanish qoidalari:</b>\n"
            "1. Botdan faqat <b>18 yoshga to'lganlar</b> foydalanishi mumkin.\n"
            "2. Haqorat, behayo suratlar, spam va firibgarlik qat'iyan taqiqlanadi (profil butunlay bloklanadi).\n"
            "3. O'zaro hurmat va samimiylik asosiy mezondir.\n\n"
            "<i>Davom etish uchun bot qoidalariga rozilik bildiring:</i>"
        ),
    },
    "about_text": {
        "title": "ℹ️ 'Bot haqida' ma'lumot matni",
        "default": (
            "✨ <b>XLM Anonymous Dating Bot</b> — O'zbekistondagi eng qulay va xavfsiz 18+ anonim tanishuv xizmati.\n\n"
            "🔒 <b>Maxfiylik kafolati:</b> Sizning Telegram profilingiz, username va telefon raqamingiz boshqa suhbatdoshlarga aslo ko'rsatilmaydi.\n\n"
            "💌 <b>Shaxsiy chat:</b> Sizning shaxsiy chatingiz pullik obunadagi (Premium) foydalanuvchilarga ko'rinadi va ochiladi.\n\n"
            "⭐ <b>Premium imkoniyatlar:</b> Shaxsiy anonim chat ochish, o'g'il bola yoki qiz bola bo'yicha alohida saralash, qidiruvda ustunlik va oltin belgi.\n\n"
            "🛡 <b>Jamiyat qoidalari:</b> 18 yoshdan kichiklarga qat'iyan taqiqlanadi. Qoidabuzarlar darhol butunlay bloklanadi.\n\n"
            "💬 <b>Aloqa:</b> Savol va takliflar uchun '💬 Adminga yozish' bo'limidan foydalaning."
        ),
    },
    "payment_instructions": {
        "title": "💳 To'lov yo'riqnomasi",
        "default": (
            "Iltimos, to'lovni ko'rsatilgan kartaga o'tkazing va to'lov kvitansiyasini (skrinshot yoki PDF) botga yuboring. "
            "Administrator to'lovni bank ilovasi orqali tekshirib tasdiqlashi bilan profilingizda Premium faollashadi."
        ),
    },
    "underage_text": {
        "title": "🔞 18 yoshdan kichiklarni rad etish",
        "default": (
            "🔞 <b>Kechirasiz!</b> Ushbu bot qat'iyan <b>18 yoshga to'lgan</b> foydalanuvchilar uchun mo'ljallangan. "
            "Voyaga yetmagan shaxslarning ro'yxatdan o'tishiga ruxsat berilmaydi."
        ),
    },
    "match_text": {
        "title": "🎉 O'zaro Like (Match) tabrik xabari",
        "default": (
            "🎉 <b>O'ZARO MOSLIK (MATCH)!</b>\n\n"
            "Siz va suhbatdosh bir-biringizga like bosdingiz! Endi anonim suhbat so'rovi yuborishingiz mumkin."
        ),
    },
    "maintenance_text": {
        "title": "🛠 Texnik xizmat rejimi xabari",
        "default": (
            "🛠 <b>Bot hozirda texnik ta'mirlash rejimida.</b>\n"
            "Tez orada xizmat to'liq tiklanadi. Iltimos, kuting."
        ),
    },
    "hidden_profile_warning": {
        "title": "🙈 Yashirilgan profil ogohlantirishi",
        "default": (
            "🙈 <b>Sizning anketangiz yashirilgan holatda!</b>\n\n"
            "Boshqalarning anketalarini ko'rish uchun avval '👤 Anketam' bo'limiga o'tib, anketangizni ko'rinadigan qiling (👁️‍🗨️ Anketani ko'rsatish)."
        ),
    },
    "non_prime_filter_warning": {
        "title": "⭐ Oddiy foydalanuvchiga jins filtri taqiqi",
        "default": (
            "⭐ <b>Jins bo'yicha saralash faqat PREMIUM foydalanuvchilar uchun!</b>\n\n"
            "Oddiy foydalanuvchilar faqat <b>'🌍 Hamma'</b> bo'limi orqali anketalarni ko'rishlari mumkin. "
            "O'g'il bola yoki qiz bolani alohida filtrlash uchun Premium obunani xarid qiling!"
        ),
    },
}

# ==============================================================================
# 2. MA'LUMOTLAR BAZASI SXEMASI VA ASINXRON BOSHQARUVCHI (SQLITE)
# ==============================================================================


class ConnectionContext:
    def __init__(self, db_path: str):
        self.db_path = db_path
        self._conn = None

    async def __aenter__(self) -> aiosqlite.Connection:
        self._conn = await aiosqlite.connect(self.db_path)
        self._conn.row_factory = aiosqlite.Row
        await self._conn.execute("PRAGMA foreign_keys = ON;")
        await self._conn.execute("PRAGMA journal_mode = WAL;")
        await self._conn.execute("PRAGMA busy_timeout = 5000;")
        return self._conn

    async def __aexit__(self, exc_type, exc_val, exc_tb):
        if self._conn:
            await self._conn.close()


class Database:
    def __init__(self, db_path: str):
        self.db_path = db_path
        self._lock = asyncio.Lock()

    async def get_connection(self) -> ConnectionContext:
        return ConnectionContext(self.db_path)

    async def init_db(self):
        """Barcha zarur jadvallarni yaratish va boshlang'ich qiymatlarni o'rnatish."""
        async with self._lock:
            async with await self.get_connection() as conn:
                # 1. Foydalanuvchilar
                await conn.execute(
                    """
                    CREATE TABLE IF NOT EXISTS users (
                        user_id INTEGER PRIMARY KEY,
                        username TEXT,
                        first_name TEXT,
                        last_name TEXT,
                        phone_number TEXT,
                        created_at TEXT NOT NULL,
                        is_banned INTEGER DEFAULT 0,
                        ban_reason TEXT,
                        is_admin INTEGER DEFAULT 0,
                        last_active_at TEXT NOT NULL
                    );
                """
                )

                try:
                    await conn.execute("ALTER TABLE users ADD COLUMN phone_number TEXT;")
                except Exception:
                    pass

                # 2. Anketalar (Profiles)
                await conn.execute(
                    """
                    CREATE TABLE IF NOT EXISTS profiles (
                        user_id INTEGER PRIMARY KEY,
                        name TEXT NOT NULL,
                        age INTEGER NOT NULL,
                        birthdate TEXT,
                        gender TEXT NOT NULL,
                        region TEXT NOT NULL,
                        bio TEXT,
                        is_active INTEGER DEFAULT 1,
                        is_hidden INTEGER DEFAULT 0,
                        allow_chat_requests INTEGER DEFAULT 1,
                        is_verified INTEGER DEFAULT 0,
                        created_at TEXT NOT NULL,
                        updated_at TEXT NOT NULL,
                        FOREIGN KEY (user_id) REFERENCES users(user_id) ON DELETE CASCADE
                    );
                """
                )

                # 3. Anketa rasmlari (1 dan 3 tagacha)
                await conn.execute(
                    """
                    CREATE TABLE IF NOT EXISTS profile_photos (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        user_id INTEGER NOT NULL,
                        file_id TEXT NOT NULL,
                        file_unique_id TEXT,
                        sort_order INTEGER DEFAULT 1,
                        created_at TEXT NOT NULL,
                        FOREIGN KEY (user_id) REFERENCES users(user_id) ON DELETE CASCADE
                    );
                """
                )

                # 4. Layklar va Like emas (Swipes)
                await conn.execute(
                    """
                    CREATE TABLE IF NOT EXISTS likes (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        from_user_id INTEGER NOT NULL,
                        to_user_id INTEGER NOT NULL,
                        is_like INTEGER NOT NULL,
                        created_at TEXT NOT NULL,
                        UNIQUE(from_user_id, to_user_id),
                        FOREIGN KEY (from_user_id) REFERENCES users(user_id) ON DELETE CASCADE,
                        FOREIGN KEY (to_user_id) REFERENCES users(user_id) ON DELETE CASCADE
                    );
                """
                )

                # 5. Bloklangan foydalanuvchilar
                await conn.execute(
                    """
                    CREATE TABLE IF NOT EXISTS blocks (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        blocker_id INTEGER NOT NULL,
                        blocked_id INTEGER NOT NULL,
                        reason TEXT,
                        created_at TEXT NOT NULL,
                        UNIQUE(blocker_id, blocked_id),
                        FOREIGN KEY (blocker_id) REFERENCES users(user_id) ON DELETE CASCADE,
                        FOREIGN KEY (blocked_id) REFERENCES users(user_id) ON DELETE CASCADE
                    );
                """
                )

                # 6. Anonim chat so'rovlari
                await conn.execute(
                    """
                    CREATE TABLE IF NOT EXISTS chat_requests (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        sender_id INTEGER NOT NULL,
                        receiver_id INTEGER NOT NULL,
                        status TEXT DEFAULT 'PENDING',
                        created_at TEXT NOT NULL,
                        responded_at TEXT,
                        FOREIGN KEY (sender_id) REFERENCES users(user_id) ON DELETE CASCADE,
                        FOREIGN KEY (receiver_id) REFERENCES users(user_id) ON DELETE CASCADE
                    );
                """
                )

                # 7. Faol va yakunlangan anonim chatlar
                await conn.execute(
                    """
                    CREATE TABLE IF NOT EXISTS chats (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        user1_id INTEGER NOT NULL,
                        user2_id INTEGER NOT NULL,
                        status TEXT DEFAULT 'ACTIVE',
                        started_at TEXT NOT NULL,
                        ended_at TEXT,
                        ended_by INTEGER,
                        FOREIGN KEY (user1_id) REFERENCES users(user_id) ON DELETE CASCADE,
                        FOREIGN KEY (user2_id) REFERENCES users(user_id) ON DELETE CASCADE
                    );
                """
                )

                # 8. Chat xabarlari jurnali
                await conn.execute(
                    """
                    CREATE TABLE IF NOT EXISTS chat_messages (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        chat_id INTEGER NOT NULL,
                        sender_id INTEGER NOT NULL,
                        message_type TEXT NOT NULL,
                        content TEXT,
                        file_id TEXT,
                        created_at TEXT NOT NULL,
                        FOREIGN KEY (chat_id) REFERENCES chats(id) ON DELETE CASCADE
                    );
                """
                )

                # 9. Premium tariflari (Admin boshqaradi)
                await conn.execute(
                    """
                    CREATE TABLE IF NOT EXISTS premium_plans (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        code TEXT UNIQUE NOT NULL,
                        name TEXT NOT NULL,
                        duration_days INTEGER NOT NULL,
                        price_uzs INTEGER NOT NULL,
                        description TEXT,
                        is_active INTEGER DEFAULT 1,
                        sort_order INTEGER DEFAULT 0
                    );
                """
                )

                # 10. To'lovlar (Faqat admin qo'lda tasdiqlaydi!)
                await conn.execute(
                    """
                    CREATE TABLE IF NOT EXISTS payments (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        user_id INTEGER NOT NULL,
                        plan_id INTEGER NOT NULL,
                        amount_uzs INTEGER NOT NULL,
                        receipt_file_id TEXT NOT NULL,
                        receipt_type TEXT NOT NULL,
                        status TEXT DEFAULT 'PENDING',
                        created_at TEXT NOT NULL,
                        reviewed_by INTEGER,
                        reviewed_at TEXT,
                        rejection_reason TEXT,
                        FOREIGN KEY (user_id) REFERENCES users(user_id) ON DELETE CASCADE,
                        FOREIGN KEY (plan_id) REFERENCES premium_plans(id)
                    );
                """
                )

                # 11. Premium obunalar
                await conn.execute(
                    """
                    CREATE TABLE IF NOT EXISTS subscriptions (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        user_id INTEGER UNIQUE NOT NULL,
                        start_date TEXT NOT NULL,
                        end_date TEXT NOT NULL,
                        is_active INTEGER DEFAULT 1,
                        updated_at TEXT NOT NULL,
                        FOREIGN KEY (user_id) REFERENCES users(user_id) ON DELETE CASCADE
                    );
                """
                )

                # 12. Majburiy obuna kanallari
                await conn.execute(
                    """
                    CREATE TABLE IF NOT EXISTS required_channels (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        title TEXT NOT NULL,
                        username TEXT,
                        invite_link TEXT NOT NULL,
                        channel_id INTEGER,
                        is_active INTEGER DEFAULT 1,
                        created_at TEXT NOT NULL
                    );
                """
                )

                # 13. Shikoyatlar (Reports)
                await conn.execute(
                    """
                    CREATE TABLE IF NOT EXISTS reports (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        reporter_id INTEGER NOT NULL,
                        reported_user_id INTEGER NOT NULL,
                        reason TEXT NOT NULL,
                        details TEXT,
                        status TEXT DEFAULT 'PENDING',
                        created_at TEXT NOT NULL,
                        resolved_by INTEGER,
                        resolution_notes TEXT,
                        FOREIGN KEY (reporter_id) REFERENCES users(user_id) ON DELETE CASCADE,
                        FOREIGN KEY (reported_user_id) REFERENCES users(user_id) ON DELETE CASCADE
                    );
                """
                )

                # 14. Admin faoliyati jurnali (Audit log)
                await conn.execute(
                    """
                    CREATE TABLE IF NOT EXISTS admin_logs (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        admin_id INTEGER NOT NULL,
                        action TEXT NOT NULL,
                        target_id TEXT,
                        details TEXT,
                        created_at TEXT NOT NULL
                    );
                """
                )

                # 15. Ommaviy xabarlar jurnali (Broadcast jobs)
                await conn.execute(
                    """
                    CREATE TABLE IF NOT EXISTS broadcast_jobs (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        admin_id INTEGER NOT NULL,
                        target_filter TEXT NOT NULL,
                        total_count INTEGER DEFAULT 0,
                        success_count INTEGER DEFAULT 0,
                        failed_count INTEGER DEFAULT 0,
                        status TEXT DEFAULT 'COMPLETED',
                        created_at TEXT NOT NULL,
                        finished_at TEXT
                    );
                """
                )

                # 16. Bot umumiy sozlamalari
                await conn.execute(
                    """
                    CREATE TABLE IF NOT EXISTS settings (
                        key TEXT PRIMARY KEY,
                        value TEXT NOT NULL,
                        description TEXT,
                        updated_at TEXT NOT NULL
                    );
                """
                )

                # 17. Foydalanuvchi roziliklari (18+ va qoidalar)
                await conn.execute(
                    """
                    CREATE TABLE IF NOT EXISTS user_consents (
                        user_id INTEGER PRIMARY KEY,
                        agreed_terms INTEGER DEFAULT 1,
                        agreed_at TEXT NOT NULL,
                        age_confirmed INTEGER DEFAULT 1,
                        FOREIGN KEY (user_id) REFERENCES users(user_id) ON DELETE CASCADE
                    );
                """
                )

                # 18. Yangi adminlar va cheklangan ruxsatlar (RBAC)
                await conn.execute(
                    """
                    CREATE TABLE IF NOT EXISTS admins (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        user_id INTEGER UNIQUE NOT NULL,
                        added_by INTEGER NOT NULL,
                        admin_name TEXT NOT NULL,
                        role_title TEXT DEFAULT 'Admin',
                        permissions TEXT NOT NULL,
                        is_active INTEGER DEFAULT 1,
                        created_at TEXT NOT NULL,
                        updated_at TEXT NOT NULL,
                        FOREIGN KEY (user_id) REFERENCES users(user_id) ON DELETE CASCADE
                    );
                """
                )

                # Indekslar
                await conn.execute("CREATE INDEX IF NOT EXISTS idx_users_banned ON users(is_banned);")
                await conn.execute("CREATE INDEX IF NOT EXISTS idx_profiles_active_gender ON profiles(is_active, is_hidden, gender);")
                await conn.execute("CREATE INDEX IF NOT EXISTS idx_likes_pair ON likes(from_user_id, to_user_id);")
                await conn.execute("CREATE INDEX IF NOT EXISTS idx_blocks_pair ON blocks(blocker_id, blocked_id);")
                await conn.execute("CREATE INDEX IF NOT EXISTS idx_chats_users ON chats(user1_id, user2_id, status);")
                await conn.execute("CREATE INDEX IF NOT EXISTS idx_payments_status ON payments(status);")
                await conn.execute("CREATE INDEX IF NOT EXISTS idx_subscriptions_user ON subscriptions(user_id, is_active);")
                await conn.execute("CREATE INDEX IF NOT EXISTS idx_admins_user ON admins(user_id, is_active);")

                now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                # Standart Premium tariflarini kiritish
                plans_data = [
                    ("1_month", "1 Oylik Premium", 30, 25000, "30 kunlik anonim shaxsiy suhbat va cheksiz tanishuv.", 1, 1),
                    ("3_months", "3 Oylik Premium", 90, 65000, "90 kunlik Premium obuna. Eng ommabop!", 1, 2),
                    ("9_months", "9 Oylik Premium", 270, 175000, "270 kunlik barcha imkoniyatlar uchun tejamkor tarif.", 1, 3),
                    ("12_months", "1 Yillik Premium", 365, 220000, "365 kunlik to'liq VIP Premium maqomi va cheksiz erkinlik.", 1, 4),
                ]
                for p in plans_data:
                    await conn.execute(
                        """
                        INSERT OR IGNORE INTO premium_plans (code, name, duration_days, price_uzs, description, is_active, sort_order)
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                    """,
                        p,
                    )

                # Standart sozlamalar va CMS matnlari
                default_settings = [
                    ("card_number", "8600 0000 0000 0000", "To'lov uchun karta raqami"),
                    ("card_holder", "XLM DATING SERVICE", "Karta egasining to'liq ismi"),
                    ("maintenance_mode", "0", "Xizmat rejimi (0 - normal, 1 - texnik ishlar)"),
                ]
                for key, data in DEFAULT_SYSTEM_TEXTS.items():
                    default_settings.append((key, data["default"], data["title"]))

                for key, val, desc in default_settings:
                    await conn.execute(
                        """
                        INSERT OR IGNORE INTO settings (key, value, description, updated_at)
                        VALUES (?, ?, ?, ?)
                    """,
                        (key, val, desc, now_str),
                    )

                await conn.commit()
        logger.info("Ma'lumotlar bazasi jadvallari va sozlamalari muvaffaqiyatli tekshirildi.")

    # ------------------ FOYDALANUVCHILAR ------------------

    async def get_or_create_user(
        self,
        user_id: int,
        username: Optional[str],
        first_name: Optional[str],
        last_name: Optional[str],
        phone_number: Optional[str] = None,
    ) -> aiosqlite.Row:
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        is_admin = 1 if user_id in ADMIN_IDS else 0
        async with self._lock:
            async with await self.get_connection() as conn:
                await conn.execute(
                    """
                    INSERT INTO users (user_id, username, first_name, last_name, phone_number, created_at, is_admin, last_active_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    ON CONFLICT(user_id) DO UPDATE SET
                        username = excluded.username,
                        first_name = excluded.first_name,
                        last_name = excluded.last_name,
                        phone_number = COALESCE(excluded.phone_number, users.phone_number),
                        last_active_at = excluded.last_active_at,
                        is_admin = CASE WHEN excluded.user_id IN ({}) THEN 1 ELSE users.is_admin END
                """.format(",".join("?" * len(ADMIN_IDS)) if ADMIN_IDS else "0"),
                    (
                        user_id,
                        username,
                        first_name,
                        last_name,
                        phone_number,
                        now_str,
                        is_admin,
                        now_str,
                        *(list(ADMIN_IDS) if ADMIN_IDS else []),
                    ),
                )
                await conn.commit()
                async with conn.execute("SELECT * FROM users WHERE user_id = ?", (user_id,)) as cursor:
                    return await cursor.fetchone()

    async def update_user_phone(self, user_id: int, phone: str):
        async with self._lock:
            async with await self.get_connection() as conn:
                await conn.execute("UPDATE users SET phone_number = ? WHERE user_id = ?", (phone, user_id))
                await conn.commit()

    async def is_user_banned(self, user_id: int) -> Tuple[bool, Optional[str]]:
        async with await self.get_connection() as conn:
            async with conn.execute("SELECT is_banned, ban_reason FROM users WHERE user_id = ?", (user_id,)) as cursor:
                row = await cursor.fetchone()
                if row and row["is_banned"] == 1:
                    return True, row["ban_reason"]
                return False, None

    async def ban_user(self, user_id: int, reason: str, admin_id: int) -> bool:
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        async with self._lock:
            async with await self.get_connection() as conn:
                await conn.execute("UPDATE users SET is_banned = 1, ban_reason = ? WHERE user_id = ?", (reason, user_id))
                await conn.execute("UPDATE profiles SET is_active = 0 WHERE user_id = ?", (user_id,))
                await conn.execute(
                    "UPDATE chats SET status = 'ENDED', ended_at = ?, ended_by = ? WHERE (user1_id = ? OR user2_id = ?) AND status = 'ACTIVE'",
                    (now_str, admin_id, user_id, user_id),
                )
                await conn.execute(
                    "INSERT INTO admin_logs (admin_id, action, target_id, details, created_at) VALUES (?, 'BAN_USER', ?, ?, ?)",
                    (admin_id, str(user_id), f"Sabab: {reason}", now_str),
                )
                await conn.commit()
                return True

    async def unban_user(self, user_id: int, admin_id: int) -> bool:
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        async with self._lock:
            async with await self.get_connection() as conn:
                await conn.execute("UPDATE users SET is_banned = 0, ban_reason = NULL WHERE user_id = ?", (user_id,))
                await conn.execute("UPDATE profiles SET is_active = 1 WHERE user_id = ?", (user_id,))
                await conn.execute(
                    "INSERT INTO admin_logs (admin_id, action, target_id, details, created_at) VALUES (?, 'UNBAN_USER', ?, 'Blokdan chiqarildi', ?)",
                    (admin_id, str(user_id), now_str),
                )
                await conn.commit()
                return True

    # ------------------ ADMINLAR VA HUQUQLAR (RBAC) ------------------

    async def is_super_admin(self, user_id: int) -> bool:
        return user_id in ADMIN_IDS

    async def is_any_admin(self, user_id: int) -> bool:
        if user_id in ADMIN_IDS:
            return True
        async with await self.get_connection() as conn:
            async with conn.execute("SELECT id FROM admins WHERE user_id = ? AND is_active = 1", (user_id,)) as cur:
                return (await cur.fetchone()) is not None

    async def is_admin_allowed(self, user_id: int, permission_name: str) -> bool:
        if user_id in ADMIN_IDS:
            return True
        async with await self.get_connection() as conn:
            async with conn.execute("SELECT permissions FROM admins WHERE user_id = ? AND is_active = 1", (user_id,)) as cur:
                row = await cur.fetchone()
                if not row:
                    return False
                try:
                    perms = json.loads(row["permissions"])
                    return bool(perms.get(permission_name, False))
                except Exception:
                    return False

    async def add_subadmin(self, user_id: int, added_by: int, admin_name: str, role_title: str = "Admin") -> int:
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        perms_json = json.dumps(DEFAULT_SUBADMIN_PERMS)
        async with self._lock:
            async with await self.get_connection() as conn:
                cur = await conn.execute(
                    """
                    INSERT INTO admins (user_id, added_by, admin_name, role_title, permissions, is_active, created_at, updated_at)
                    VALUES (?, ?, ?, ?, ?, 1, ?, ?)
                    ON CONFLICT(user_id) DO UPDATE SET
                        admin_name = excluded.admin_name,
                        role_title = excluded.role_title,
                        is_active = 1,
                        updated_at = excluded.updated_at
                """,
                    (user_id, added_by, admin_name, role_title, perms_json, now_str, now_str),
                )
                await conn.commit()
                return cur.lastrowid

    async def get_all_subadmins(self) -> List[aiosqlite.Row]:
        async with await self.get_connection() as conn:
            async with conn.execute("SELECT a.*, u.username, u.first_name FROM admins a JOIN users u ON a.user_id = u.user_id ORDER BY a.id DESC") as cur:
                return await cur.fetchall()

    async def get_subadmin(self, user_id: int) -> Optional[aiosqlite.Row]:
        async with await self.get_connection() as conn:
            async with conn.execute("SELECT * FROM admins WHERE user_id = ?", (user_id,)) as cur:
                return await cur.fetchone()

    async def toggle_admin_permission(self, user_id: int, perm_key: str) -> Optional[bool]:
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        async with self._lock:
            async with await self.get_connection() as conn:
                async with conn.execute("SELECT permissions FROM admins WHERE user_id = ?", (user_id,)) as cur:
                    row = await cur.fetchone()
                    if not row:
                        return None
                    try:
                        perms = json.loads(row["permissions"])
                    except Exception:
                        perms = dict(DEFAULT_SUBADMIN_PERMS)

                new_val = not perms.get(perm_key, False)
                perms[perm_key] = new_val
                await conn.execute("UPDATE admins SET permissions = ?, updated_at = ? WHERE user_id = ?", (json.dumps(perms), now_str, user_id))
                await conn.commit()
                return new_val

    async def remove_subadmin(self, user_id: int, admin_id: int) -> bool:
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        async with self._lock:
            async with await self.get_connection() as conn:
                await conn.execute("DELETE FROM admins WHERE user_id = ?", (user_id,))
                await conn.execute(
                    "INSERT INTO admin_logs (admin_id, action, target_id, details, created_at) VALUES (?, 'REMOVE_ADMIN', ?, 'Adminlik huquqi olib tashlandi', ?)",
                    (admin_id, str(user_id), now_str),
                )
                await conn.commit()
                return True

    # ------------------ FOYDALANUVCHILAR PAGINATSIYASI ------------------

    async def get_users_paginated(self, page: int = 1, per_page: int = 10, gender_filter: str = "all") -> Tuple[List[Dict[str, Any]], int, int]:
        offset = (page - 1) * per_page
        async with await self.get_connection() as conn:
            where_clauses = []
            params: List[Any] = []

            if gender_filter in ["male", "female"]:
                where_clauses.append("p.gender = ?")
                params.append(gender_filter)
            elif gender_filter == "premium":
                now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                where_clauses.append("s.is_active = 1 AND s.end_date > ?")
                params.append(now_str)
            elif gender_filter == "banned":
                where_clauses.append("u.is_banned = 1")

            where_str = " WHERE " + " AND ".join(where_clauses) if where_clauses else ""

            count_query = f"""
                SELECT COUNT(*) as total
                FROM users u
                LEFT JOIN profiles p ON u.user_id = p.user_id
                LEFT JOIN subscriptions s ON u.user_id = s.user_id
                {where_str}
            """
            async with conn.execute(count_query, tuple(params)) as cur:
                total_count = (await cur.fetchone())["total"]

            total_pages = max(1, (total_count + per_page - 1) // per_page)

            data_query = f"""
                SELECT u.user_id, u.username, u.first_name, u.phone_number, u.is_banned, u.created_at,
                       p.name as profile_name, p.age, p.gender, p.region, p.is_active, p.is_hidden,
                       s.end_date as prem_end, s.is_active as prem_active
                FROM users u
                LEFT JOIN profiles p ON u.user_id = p.user_id
                LEFT JOIN subscriptions s ON u.user_id = s.user_id
                {where_str}
                ORDER BY u.created_at DESC
                LIMIT ? OFFSET ?
            """
            params.extend([per_page, offset])
            async with conn.execute(data_query, tuple(params)) as cur:
                rows = await cur.fetchall()
                return [dict(r) for r in rows], total_count, total_pages

    # ------------------ TO'LOVLAR PAGINATSIYASI VA TARIXI ------------------

    async def get_payments_paginated(self, page: int = 1, per_page: int = 10, status: str = "ALL") -> Tuple[List[Dict[str, Any]], int, int]:
        offset = (page - 1) * per_page
        async with await self.get_connection() as conn:
            where_clauses = []
            params: List[Any] = []

            if status in ["PENDING", "APPROVED", "REJECTED"]:
                where_clauses.append("p.status = ?")
                params.append(status)

            where_str = " WHERE " + " AND ".join(where_clauses) if where_clauses else ""

            count_query = f"SELECT COUNT(*) as total FROM payments p {where_str}"
            async with conn.execute(count_query, tuple(params)) as cur:
                total_count = (await cur.fetchone())["total"]

            total_pages = max(1, (total_count + per_page - 1) // per_page)

            data_query = f"""
                SELECT p.*, pl.name as plan_name, pl.duration_days, u.first_name, u.username
                FROM payments p
                JOIN premium_plans pl ON p.plan_id = pl.id
                JOIN users u ON p.user_id = u.user_id
                {where_str}
                ORDER BY p.id DESC
                LIMIT ? OFFSET ?
            """
            params.extend([per_page, offset])
            async with conn.execute(data_query, tuple(params)) as cur:
                rows = await cur.fetchall()
                return [dict(r) for r in rows], total_count, total_pages

    # ------------------ ANKETALAR (PROFILES) ------------------

    async def get_profile(self, user_id: int) -> Optional[aiosqlite.Row]:
        async with await self.get_connection() as conn:
            async with conn.execute("SELECT * FROM profiles WHERE user_id = ?", (user_id,)) as cursor:
                return await cursor.fetchone()

    async def get_profile_photos(self, user_id: int) -> List[aiosqlite.Row]:
        async with await self.get_connection() as conn:
            async with conn.execute(
                "SELECT * FROM profile_photos WHERE user_id = ? ORDER BY sort_order ASC, id ASC",
                (user_id,),
            ) as cursor:
                return await cursor.fetchall()

    async def save_profile(
        self,
        user_id: int,
        name: str,
        age: int,
        gender: str,
        region: str,
        bio: Optional[str],
        photos: List[str],
        birthdate: Optional[str] = None,
    ) -> bool:
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        async with self._lock:
            async with await self.get_connection() as conn:
                await conn.execute(
                    """
                    INSERT INTO profiles (user_id, name, age, birthdate, gender, region, bio, is_active, is_hidden, created_at, updated_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?, 1, 0, ?, ?)
                    ON CONFLICT(user_id) DO UPDATE SET
                        name = excluded.name,
                        age = excluded.age,
                        birthdate = excluded.birthdate,
                        gender = excluded.gender,
                        region = excluded.region,
                        bio = excluded.bio,
                        is_active = 1,
                        updated_at = excluded.updated_at
                """,
                    (user_id, name, age, birthdate, gender, region, bio, now_str, now_str),
                )

                await conn.execute("DELETE FROM profile_photos WHERE user_id = ?", (user_id,))
                for index, file_id in enumerate(photos[:3], start=1):
                    await conn.execute(
                        "INSERT INTO profile_photos (user_id, file_id, sort_order, created_at) VALUES (?, ?, ?, ?)",
                        (user_id, file_id, index, now_str),
                    )

                await conn.execute(
                    "INSERT OR REPLACE INTO user_consents (user_id, agreed_terms, agreed_at, age_confirmed) VALUES (?, 1, ?, 1)",
                    (user_id, now_str),
                )
                await conn.commit()
                return True

    async def update_profile_field(self, user_id: int, field: str, value: Any) -> bool:
        allowed = ["name", "age", "gender", "region", "bio", "is_hidden", "allow_chat_requests", "is_active"]
        if field not in allowed:
            return False
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        async with self._lock:
            async with await self.get_connection() as conn:
                await conn.execute(f"UPDATE profiles SET {field} = ?, updated_at = ? WHERE user_id = ?", (value, now_str, user_id))
                await conn.commit()
                return True

    async def delete_profile_completely(self, user_id: int) -> bool:
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        async with self._lock:
            async with await self.get_connection() as conn:
                await conn.execute(
                    "UPDATE chats SET status = 'ENDED', ended_at = ?, ended_by = ? WHERE (user1_id = ? OR user2_id = ?) AND status = 'ACTIVE'",
                    (now_str, user_id, user_id, user_id),
                )
                await conn.execute("DELETE FROM chat_requests WHERE sender_id = ? OR receiver_id = ?", (user_id, user_id))
                await conn.execute("DELETE FROM likes WHERE from_user_id = ? OR to_user_id = ?", (user_id, user_id))
                await conn.execute("DELETE FROM profile_photos WHERE user_id = ?", (user_id,))
                await conn.execute("DELETE FROM profiles WHERE user_id = ?", (user_id,))
                await conn.commit()
                return True

    # ------------------ QIDIRUV VA LAYKLAR ------------------

    async def get_next_candidate_profile(self, user_id: int, gender_filter: str = "all") -> Optional[aiosqlite.Row]:
        """
        Avval ko'rilmagan (likes jadvalida yo'q) anketalarni ko'rsatadi.
        Agar hammasi ko'rib bo'linsa — dislike qilinganlarni qayta ko'rsatadi.
        Like qilinganlar hech qachon qayta ko'rsatilmaydi.
        """
        async with await self.get_connection() as conn:
            gender_cond = ""
            gender_params: List[Any] = []
            if gender_filter in ["male", "female"]:
                gender_cond = " AND p.gender = ?"
                gender_params = [gender_filter]

            block_subq = """
                AND p.user_id NOT IN (
                    SELECT blocked_id FROM blocks WHERE blocker_id = ?
                    UNION
                    SELECT blocker_id FROM blocks WHERE blocked_id = ?
                )
            """

            # 1-urinish: Hech ko'rilmagan (likes jadvalida umuman yo'q) anketalar
            q_unseen = f"""
                SELECT p.*, u.username
                FROM profiles p
                JOIN users u ON p.user_id = u.user_id
                WHERE p.user_id != ?
                  AND p.is_active = 1
                  AND p.is_hidden = 0
                  AND u.is_banned = 0
                  AND p.user_id NOT IN (SELECT to_user_id FROM likes WHERE from_user_id = ?)
                  {block_subq}
                  {gender_cond}
                ORDER BY RANDOM() LIMIT 1;
            """
            params_unseen: List[Any] = [user_id, user_id, user_id, user_id] + gender_params
            async with conn.execute(q_unseen, tuple(params_unseen)) as cursor:
                row = await cursor.fetchone()
            if row:
                return row

            # 2-urinish: Barcha anketalar ko'rilgan — dislike qilinganlarni qayta ko'rsat
            # (Like qilinganlar is_like=1 hech qachon qayta chiqmaydi)
            q_reseen = f"""
                SELECT p.*, u.username
                FROM profiles p
                JOIN users u ON p.user_id = u.user_id
                WHERE p.user_id != ?
                  AND p.is_active = 1
                  AND p.is_hidden = 0
                  AND u.is_banned = 0
                  AND p.user_id IN (
                      SELECT to_user_id FROM likes
                      WHERE from_user_id = ? AND is_like = 0
                  )
                  AND p.user_id NOT IN (
                      SELECT to_user_id FROM likes
                      WHERE from_user_id = ? AND is_like = 1
                  )
                  {block_subq}
                  {gender_cond}
                ORDER BY RANDOM() LIMIT 1;
            """
            params_reseen: List[Any] = [user_id, user_id, user_id, user_id, user_id] + gender_params
            async with conn.execute(q_reseen, tuple(params_reseen)) as cursor:
                return await cursor.fetchone()

    async def add_like_or_dislike(self, from_user_id: int, to_user_id: int, is_like: int) -> bool:
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        async with self._lock:
            async with await self.get_connection() as conn:
                await conn.execute(
                    """
                    INSERT INTO likes (from_user_id, to_user_id, is_like, created_at)
                    VALUES (?, ?, ?, ?)
                    ON CONFLICT(from_user_id, to_user_id) DO UPDATE SET
                        is_like = excluded.is_like,
                        created_at = excluded.created_at
                """,
                    (from_user_id, to_user_id, is_like, now_str),
                )
                await conn.commit()
                return True

    async def check_mutual_like(self, user_a: int, user_b: int) -> Tuple[bool, bool]:
        async with await self.get_connection() as conn:
            async with conn.execute(
                "SELECT from_user_id, to_user_id, is_like FROM likes WHERE (from_user_id = ? AND to_user_id = ?) OR (from_user_id = ? AND to_user_id = ?)",
                (user_a, user_b, user_b, user_a),
            ) as cursor:
                rows = await cursor.fetchall()
                like_a_to_b = False
                like_b_to_a = False
                for r in rows:
                    if r["from_user_id"] == user_a and r["to_user_id"] == user_b and r["is_like"] == 1:
                        like_a_to_b = True
                    if r["from_user_id"] == user_b and r["to_user_id"] == user_a and r["is_like"] == 1:
                        like_b_to_a = True
                return like_a_to_b, like_b_to_a

    # ------------------ BLOKLAR VA SHIKOYATLAR ------------------

    async def block_user(self, blocker_id: int, blocked_id: int, reason: str = "Foydalanuvchi tanlovi") -> bool:
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        async with self._lock:
            async with await self.get_connection() as conn:
                await conn.execute(
                    "INSERT OR REPLACE INTO blocks (blocker_id, blocked_id, reason, created_at) VALUES (?, ?, ?, ?)",
                    (blocker_id, blocked_id, reason, now_str),
                )
                await conn.execute(
                    """
                    UPDATE chats SET status = 'ENDED', ended_at = ?, ended_by = ?
                    WHERE ((user1_id = ? AND user2_id = ?) OR (user1_id = ? AND user2_id = ?)) AND status = 'ACTIVE'
                """,
                    (now_str, blocker_id, blocker_id, blocked_id, blocked_id, blocker_id),
                )
                await conn.commit()
                return True

    async def is_blocked(self, user_a: int, user_b: int) -> bool:
        async with await self.get_connection() as conn:
            async with conn.execute(
                "SELECT id FROM blocks WHERE (blocker_id = ? AND blocked_id = ?) OR (blocker_id = ? AND blocked_id = ?)",
                (user_a, user_b, user_b, user_a),
            ) as cursor:
                row = await cursor.fetchone()
                return row is not None

    async def create_report(self, reporter_id: int, reported_user_id: int, reason: str, details: Optional[str] = None) -> int:
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        async with self._lock:
            async with await self.get_connection() as conn:
                cursor = await conn.execute(
                    "INSERT INTO reports (reporter_id, reported_user_id, reason, details, status, created_at) VALUES (?, ?, ?, ?, 'PENDING', ?)",
                    (reporter_id, reported_user_id, reason, details, now_str),
                )
                await conn.commit()
                return cursor.lastrowid

    # ------------------ ANONIM CHAT ------------------

    async def get_active_chat(self, user_id: int) -> Optional[aiosqlite.Row]:
        async with await self.get_connection() as conn:
            async with conn.execute("SELECT * FROM chats WHERE (user1_id = ? OR user2_id = ?) AND status = 'ACTIVE' LIMIT 1", (user_id, user_id)) as cursor:
                return await cursor.fetchone()

    async def create_chat_request(self, sender_id: int, receiver_id: int) -> int:
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        async with self._lock:
            async with await self.get_connection() as conn:
                cursor = await conn.execute(
                    "INSERT INTO chat_requests (sender_id, receiver_id, status, created_at) VALUES (?, ?, 'PENDING', ?)",
                    (sender_id, receiver_id, now_str),
                )
                await conn.commit()
                return cursor.lastrowid

    async def accept_chat_request(self, request_id: int) -> Optional[aiosqlite.Row]:
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        async with self._lock:
            async with await self.get_connection() as conn:
                async with conn.execute("SELECT * FROM chat_requests WHERE id = ? AND status = 'PENDING'", (request_id,)) as cursor:
                    req = await cursor.fetchone()
                    if not req:
                        return None

                await conn.execute("UPDATE chat_requests SET status = 'ACCEPTED', responded_at = ? WHERE id = ?", (now_str, request_id))
                await conn.execute(
                    "UPDATE chats SET status = 'ENDED', ended_at = ? WHERE (user1_id IN (?, ?) OR user2_id IN (?, ?)) AND status = 'ACTIVE'",
                    (now_str, req["sender_id"], req["receiver_id"], req["sender_id"], req["receiver_id"]),
                )

                cursor2 = await conn.execute(
                    "INSERT INTO chats (user1_id, user2_id, status, started_at) VALUES (?, ?, 'ACTIVE', ?)",
                    (req["sender_id"], req["receiver_id"], now_str),
                )
                chat_id = cursor2.lastrowid
                await conn.commit()

                async with conn.execute("SELECT * FROM chats WHERE id = ?", (chat_id,)) as cur:
                    return await cur.fetchone()

    async def reject_chat_request(self, request_id: int) -> bool:
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        async with self._lock:
            async with await self.get_connection() as conn:
                await conn.execute("UPDATE chat_requests SET status = 'REJECTED', responded_at = ? WHERE id = ? AND status = 'PENDING'", (now_str, request_id))
                await conn.commit()
                return True

    async def end_active_chat(self, chat_id: int, ended_by: int) -> Optional[Tuple[int, int]]:
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        async with self._lock:
            async with await self.get_connection() as conn:
                async with conn.execute("SELECT user1_id, user2_id FROM chats WHERE id = ? AND status = 'ACTIVE'", (chat_id,)) as cursor:
                    chat = await cursor.fetchone()
                    if not chat:
                        return None

                await conn.execute("UPDATE chats SET status = 'ENDED', ended_at = ?, ended_by = ? WHERE id = ?", (now_str, ended_by, chat_id))
                await conn.commit()
                return chat["user1_id"], chat["user2_id"]

    async def log_chat_message(self, chat_id: int, sender_id: int, msg_type: str, content: Optional[str] = None, file_id: Optional[str] = None):
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        async with self._lock:
            async with await self.get_connection() as conn:
                await conn.execute(
                    "INSERT INTO chat_messages (chat_id, sender_id, message_type, content, file_id, created_at) VALUES (?, ?, ?, ?, ?, ?)",
                    (chat_id, sender_id, msg_type, content, file_id, now_str),
                )
                await conn.commit()

    # ------------------ PREMIUM VA TO'LOVLAR ------------------

    async def is_premium_active(self, user_id: int) -> Tuple[bool, Optional[str]]:
        if user_id in ADMIN_IDS:
            return True, "Cheksiz (Super Admin)"

        async with await self.get_connection() as conn:
            async with conn.execute("SELECT end_date, is_active FROM subscriptions WHERE user_id = ? AND is_active = 1", (user_id,)) as cursor:
                row = await cursor.fetchone()
                if not row:
                    return False, None

                try:
                    end_dt = datetime.strptime(row["end_date"], "%Y-%m-%d %H:%M:%S")
                    if end_dt > datetime.now():
                        return True, row["end_date"]
                    else:
                        await conn.execute("UPDATE subscriptions SET is_active = 0 WHERE user_id = ?", (user_id,))
                        await conn.commit()
                        return False, row["end_date"]
                except Exception:
                    return False, None

    async def get_premium_plans(self, only_active: bool = True) -> List[aiosqlite.Row]:
        async with await self.get_connection() as conn:
            query = "SELECT * FROM premium_plans"
            if only_active:
                query += " WHERE is_active = 1"
            query += " ORDER BY sort_order ASC, duration_days ASC"
            async with conn.execute(query) as cursor:
                return await cursor.fetchall()

    async def get_premium_plan_by_id(self, plan_id: int) -> Optional[aiosqlite.Row]:
        async with await self.get_connection() as conn:
            async with conn.execute("SELECT * FROM premium_plans WHERE id = ?", (plan_id,)) as cursor:
                return await cursor.fetchone()

    async def update_premium_plan(self, plan_id: int, price_uzs: Optional[int] = None, name: Optional[str] = None, is_active: Optional[int] = None) -> bool:
        async with self._lock:
            async with await self.get_connection() as conn:
                if price_uzs is not None:
                    await conn.execute("UPDATE premium_plans SET price_uzs = ? WHERE id = ?", (price_uzs, plan_id))
                if name is not None:
                    await conn.execute("UPDATE premium_plans SET name = ? WHERE id = ?", (name, plan_id))
                if is_active is not None:
                    await conn.execute("UPDATE premium_plans SET is_active = ? WHERE id = ?", (is_active, plan_id))
                await conn.commit()
                return True

    async def create_payment_order(self, user_id: int, plan_id: int, amount_uzs: int, receipt_file_id: str, receipt_type: str) -> int:
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        async with self._lock:
            async with await self.get_connection() as conn:
                cursor = await conn.execute(
                    "INSERT INTO payments (user_id, plan_id, amount_uzs, receipt_file_id, receipt_type, status, created_at) VALUES (?, ?, ?, ?, ?, 'PENDING', ?)",
                    (user_id, plan_id, amount_uzs, receipt_file_id, receipt_type, now_str),
                )
                await conn.commit()
                return cursor.lastrowid

    async def get_payment(self, payment_id: int) -> Optional[aiosqlite.Row]:
        async with await self.get_connection() as conn:
            async with conn.execute(
                """
                SELECT p.*, pl.name as plan_name, pl.duration_days, u.first_name, u.username
                FROM payments p
                JOIN premium_plans pl ON p.plan_id = pl.id
                JOIN users u ON p.user_id = u.user_id
                WHERE p.id = ?
            """,
                (payment_id,),
            ) as cursor:
                return await cursor.fetchone()

    async def approve_payment(self, payment_id: int, admin_id: int) -> Tuple[bool, str, Optional[int], Optional[str]]:
        now = datetime.now()
        now_str = now.strftime("%Y-%m-%d %H:%M:%S")

        async with self._lock:
            async with await self.get_connection() as conn:
                async with conn.execute(
                    "SELECT p.*, pl.duration_days, pl.name as plan_name FROM payments p JOIN premium_plans pl ON p.plan_id = pl.id WHERE p.id = ?",
                    (payment_id,),
                ) as cursor:
                    payment = await cursor.fetchone()
                    if not payment:
                        return False, "To'lov so'rovi topilmadi.", None, None

                    if payment["status"] != "PENDING":
                        return False, f"Ushbu to'lov allaqachon '{payment['status']}' qilingan!", None, None

                user_id = payment["user_id"]
                duration_days = payment["duration_days"]

                async with conn.execute("SELECT end_date FROM subscriptions WHERE user_id = ?", (user_id,)) as sub_cur:
                    sub = await sub_cur.fetchone()
                    start_date = now

                    if sub and sub["end_date"]:
                        try:
                            curr_end = datetime.strptime(sub["end_date"], "%Y-%m-%d %H:%M:%S")
                            if curr_end > now:
                                start_date = curr_end
                        except Exception:
                            start_date = now

                    new_end_date = start_date + timedelta(days=duration_days)
                    new_end_str = new_end_date.strftime("%Y-%m-%d %H:%M:%S")

                await conn.execute("UPDATE payments SET status = 'APPROVED', reviewed_by = ?, reviewed_at = ? WHERE id = ? AND status = 'PENDING'", (admin_id, now_str, payment_id))
                await conn.execute(
                    """
                    INSERT INTO subscriptions (user_id, start_date, end_date, is_active, updated_at)
                    VALUES (?, ?, ?, 1, ?)
                    ON CONFLICT(user_id) DO UPDATE SET
                        end_date = excluded.end_date,
                        is_active = 1,
                        updated_at = excluded.updated_at
                """,
                    (user_id, now_str, new_end_str, now_str),
                )
                await conn.execute(
                    "INSERT INTO admin_logs (admin_id, action, target_id, details, created_at) VALUES (?, 'APPROVE_PAYMENT', ?, ?, ?)",
                    (admin_id, str(payment_id), f"Foydalanuvchi: {user_id}, Tarif: {payment['plan_name']}, Yangi muddat: {new_end_str}", now_str),
                )
                await conn.commit()
                return True, "To'lov muvaffaqiyatli tasdiqlandi!", user_id, new_end_str

    async def reject_payment(self, payment_id: int, admin_id: int, reason: str) -> Tuple[bool, str, Optional[int]]:
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        async with self._lock:
            async with await self.get_connection() as conn:
                async with conn.execute("SELECT user_id, status FROM payments WHERE id = ?", (payment_id,)) as cursor:
                    payment = await cursor.fetchone()
                    if not payment:
                        return False, "To'lov topilmadi.", None
                    if payment["status"] != "PENDING":
                        return False, f"To'lov allaqachon '{payment['status']}' holatida.", None

                user_id = payment["user_id"]
                await conn.execute("UPDATE payments SET status = 'REJECTED', reviewed_by = ?, reviewed_at = ?, rejection_reason = ? WHERE id = ? AND status = 'PENDING'", (admin_id, now_str, reason, payment_id))
                await conn.execute("INSERT INTO admin_logs (admin_id, action, target_id, details, created_at) VALUES (?, 'REJECT_PAYMENT', ?, ?, ?)", (admin_id, str(payment_id), f"Foydalanuvchi: {user_id}, Sabab: {reason}", now_str))
                await conn.commit()
                return True, "To'lov rad etildi.", user_id

    async def grant_manual_premium(self, user_id: int, days: int, admin_id: int) -> Tuple[bool, str]:
        now = datetime.now()
        now_str = now.strftime("%Y-%m-%d %H:%M:%S")
        async with self._lock:
            async with await self.get_connection() as conn:
                async with conn.execute("SELECT end_date FROM subscriptions WHERE user_id = ?", (user_id,)) as sub_cur:
                    sub = await sub_cur.fetchone()
                    start_date = now
                    if sub and sub["end_date"]:
                        try:
                            curr_end = datetime.strptime(sub["end_date"], "%Y-%m-%d %H:%M:%S")
                            if curr_end > now:
                                start_date = curr_end
                        except Exception:
                            start_date = now

                    new_end = start_date + timedelta(days=days)
                    new_end_str = new_end.strftime("%Y-%m-%d %H:%M:%S")

                await conn.execute(
                    """
                    INSERT INTO subscriptions (user_id, start_date, end_date, is_active, updated_at)
                    VALUES (?, ?, ?, 1, ?)
                    ON CONFLICT(user_id) DO UPDATE SET
                        end_date = excluded.end_date,
                        is_active = 1,
                        updated_at = excluded.updated_at
                """,
                    (user_id, now_str, new_end_str, now_str),
                )
                await conn.execute("INSERT INTO admin_logs (admin_id, action, target_id, details, created_at) VALUES (?, 'MANUAL_PREMIUM', ?, ?, ?)", (admin_id, str(user_id), f"{days} kun berildi. Muddat: {new_end_str}", now_str))
                await conn.commit()
                return True, new_end_str

    async def revoke_premium(self, user_id: int, admin_id: int) -> Tuple[bool, str]:
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        async with self._lock:
            async with await self.get_connection() as conn:
                await conn.execute("UPDATE subscriptions SET is_active = 0, end_date = ? WHERE user_id = ?", (now_str, user_id))
                await conn.execute("INSERT INTO admin_logs (admin_id, action, target_id, details, created_at) VALUES (?, 'REVOKE_PREMIUM', ?, 'Premium bekor qilindi', ?)", (admin_id, str(user_id), now_str))
                await conn.commit()
                return True, "Premium obuna bekor qilindi."

    # ------------------ MAJBURIY OBUNA ------------------

    async def get_required_channels(self, only_active: bool = True) -> List[aiosqlite.Row]:
        async with await self.get_connection() as conn:
            query = "SELECT * FROM required_channels"
            if only_active:
                query += " WHERE is_active = 1"
            query += " ORDER BY id DESC"
            async with conn.execute(query) as cursor:
                return await cursor.fetchall()

    async def add_required_channel(self, title: str, username: Optional[str], invite_link: str, channel_id: Optional[int]) -> int:
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        async with self._lock:
            async with await self.get_connection() as conn:
                cursor = await conn.execute(
                    "INSERT INTO required_channels (title, username, invite_link, channel_id, is_active, created_at) VALUES (?, ?, ?, ?, 1, ?)",
                    (title, username, invite_link, channel_id, now_str),
                )
                await conn.commit()
                return cursor.lastrowid

    async def delete_required_channel(self, channel_id: int) -> bool:
        async with self._lock:
            async with await self.get_connection() as conn:
                await conn.execute("DELETE FROM required_channels WHERE id = ?", (channel_id,))
                await conn.commit()
                return True

    async def toggle_required_channel(self, channel_id: int) -> bool:
        async with self._lock:
            async with await self.get_connection() as conn:
                await conn.execute("UPDATE required_channels SET is_active = CASE WHEN is_active = 1 THEN 0 ELSE 1 END WHERE id = ?", (channel_id,))
                await conn.commit()
                return True

    # ------------------ SOZLAMALAR VA STATISTIKA ------------------

    async def get_setting(self, key: str, default: Optional[str] = None) -> str:
        async with await self.get_connection() as conn:
            async with conn.execute("SELECT value FROM settings WHERE key = ?", (key,)) as cursor:
                row = await cursor.fetchone()
                return row["value"] if row else (default or "")

    async def set_setting(self, key: str, value: str) -> bool:
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        async with self._lock:
            async with await self.get_connection() as conn:
                await conn.execute(
                    "INSERT INTO settings (key, value, updated_at) VALUES (?, ?, ?) ON CONFLICT(key) DO UPDATE SET value = excluded.value, updated_at = excluded.updated_at",
                    (key, value, now_str),
                )
                await conn.commit()
                return True

    async def get_statistics(self) -> Dict[str, Any]:
        async with await self.get_connection() as conn:
            stats = {}
            today_str = datetime.now().strftime("%Y-%m-%d")

            async with conn.execute("SELECT COUNT(*) as cnt FROM users") as cur:
                stats["total_users"] = (await cur.fetchone())["cnt"]
            async with conn.execute("SELECT COUNT(*) as cnt FROM profiles WHERE is_active = 1 AND is_hidden = 0") as cur:
                stats["active_profiles"] = (await cur.fetchone())["cnt"]
            async with conn.execute("SELECT COUNT(*) as cnt FROM profiles WHERE gender = 'male' AND is_active = 1") as cur:
                stats["male_count"] = (await cur.fetchone())["cnt"]
            async with conn.execute("SELECT COUNT(*) as cnt FROM profiles WHERE gender = 'female' AND is_active = 1") as cur:
                stats["female_count"] = (await cur.fetchone())["cnt"]
            async with conn.execute("SELECT COUNT(*) as cnt FROM users WHERE created_at LIKE ?", (f"{today_str}%",)) as cur:
                stats["today_registered"] = (await cur.fetchone())["cnt"]

            now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            async with conn.execute("SELECT COUNT(*) as cnt FROM subscriptions WHERE is_active = 1 AND end_date > ?", (now_str,)) as cur:
                stats["active_premium"] = (await cur.fetchone())["cnt"]
            async with conn.execute("SELECT COUNT(*) as cnt FROM subscriptions WHERE end_date <= ?", (now_str,)) as cur:
                stats["expired_premium"] = (await cur.fetchone())["cnt"]

            async with conn.execute("SELECT COUNT(*) as cnt FROM payments WHERE status = 'PENDING'") as cur:
                stats["pending_payments"] = (await cur.fetchone())["cnt"]
            async with conn.execute("SELECT COUNT(*) as cnt FROM payments WHERE status = 'APPROVED'") as cur:
                stats["approved_payments"] = (await cur.fetchone())["cnt"]
            async with conn.execute("SELECT COUNT(*) as cnt FROM payments WHERE status = 'REJECTED'") as cur:
                stats["rejected_payments"] = (await cur.fetchone())["cnt"]

            async with conn.execute("SELECT COUNT(*) as cnt FROM chats WHERE status = 'ACTIVE'") as cur:
                stats["active_chats"] = (await cur.fetchone())["cnt"]
            async with conn.execute("SELECT COUNT(*) as cnt FROM reports WHERE status = 'PENDING'") as cur:
                stats["pending_reports"] = (await cur.fetchone())["cnt"]
            async with conn.execute("SELECT COUNT(*) as cnt FROM users WHERE is_banned = 1") as cur:
                stats["banned_users"] = (await cur.fetchone())["cnt"]
            async with conn.execute("SELECT COUNT(*) as cnt FROM admins WHERE is_active = 1") as cur:
                stats["subadmins_count"] = (await cur.fetchone())["cnt"]

            return stats

    async def get_export_data(self, table_name: str) -> List[Dict[str, Any]]:
        async with await self.get_connection() as conn:
            if table_name == "users":
                query = "SELECT u.user_id, u.username, u.first_name, u.phone_number, u.created_at, u.is_banned, p.name, p.age, p.gender, p.region FROM users u LEFT JOIN profiles p ON u.user_id = p.user_id"
            elif table_name == "payments":
                query = "SELECT p.id, p.user_id, pl.name as plan_name, p.amount_uzs, p.status, p.created_at, p.reviewed_by, p.reviewed_at, p.rejection_reason FROM payments p LEFT JOIN premium_plans pl ON p.plan_id = pl.id"
            elif table_name == "reports":
                query = "SELECT id, reporter_id, reported_user_id, reason, details, status, created_at, resolved_by FROM reports"
            elif table_name == "subscriptions":
                query = "SELECT s.id, s.user_id, u.first_name, u.username, s.start_date, s.end_date, s.is_active FROM subscriptions s JOIN users u ON s.user_id = u.user_id"
            else:
                return []

            async with conn.execute(query) as cursor:
                rows = await cursor.fetchall()
                return [dict(r) for r in rows]


db = Database(DATABASE_PATH)

# ==============================================================================
# 3. FSM (FINITE STATE MACHINE) HOLATLAR GURUHI
# ==============================================================================


class RegistrationState(StatesGroup):
    waiting_for_terms = State()  # Qoidalar va shartlarga rozilik
    waiting_for_age = State()
    waiting_for_name = State()
    waiting_for_contact = State()
    waiting_for_gender = State()
    waiting_for_photos = State()
    waiting_for_region = State()
    waiting_for_bio = State()
    confirm_profile = State()


class ProfileEditState(StatesGroup):
    menu = State()
    edit_name = State()
    edit_age = State()
    edit_gender = State()
    edit_photos = State()
    edit_region = State()
    edit_bio = State()


class SearchState(StatesGroup):
    selecting_filter = State()
    browsing = State()


class ChatState(StatesGroup):
    in_active_chat = State()


class PaymentState(StatesGroup):
    selecting_plan = State()
    uploading_receipt = State()


class SupportState(StatesGroup):
    waiting_for_message = State()


class ReportState(StatesGroup):
    waiting_for_reason = State()
    waiting_for_details = State()


class AdminState(StatesGroup):
    in_admin_mode = State()  # Admin menyusida qolish holati
    search_user_id = State()
    send_message_user_id = State()
    send_message_text = State()
    grant_premium_days = State()
    ban_user_reason = State()
    reject_payment_reason = State()
    edit_plan_select = State()
    edit_plan_price = State()
    edit_plan_name = State()
    channel_title = State()
    channel_username = State()
    channel_link = State()
    broadcast_target = State()
    broadcast_content = State()
    confirm_broadcast = State()
    settings_card = State()
    settings_cardholder = State()
    settings_about = State()
    confirm_maintenance_restart = State()
    reply_support_user_id = State()
    reply_support_text = State()
    add_admin_id = State()
    add_admin_name = State()
    cms_enter_text = State()  # Matnni tahrirlash holati


# ==============================================================================
# 4. KLAVIATURALAR (REPLY KEYBOARDS & INLINE BUTTONS)
# ==============================================================================


def get_main_menu_keyboard(is_admin: bool = False) -> ReplyKeyboardMarkup:
    """
    Asosiy foydalanuvchi menyusi:
    - '🔎 Anketalarni ko‘rish' alohida 1 qatorda!
    - '🛡 Admin paneli' faqat admin vakolati bor foydalanuvchilarga ko'rinadi!
    """
    buttons = [
        [KeyboardButton(text="🔎 Anketalarni ko‘rish")],
        [
            KeyboardButton(text="👤 Anketam"),
            KeyboardButton(text="⭐ Premium olish"),
        ],
        [
            KeyboardButton(text="💬 Adminga yozish"),
            KeyboardButton(text="ℹ️ Bot haqida"),
        ],
        [KeyboardButton(text="🗑 Anketani o‘chirish")],
    ]
    if is_admin:
        buttons.append([KeyboardButton(text="🛡 Admin paneli")])
    return ReplyKeyboardMarkup(keyboard=buttons, resize_keyboard=True, is_persistent=True)


def get_terms_agreement_keyboard() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="✅ Qoidalarga roziman")],
            [KeyboardButton(text="❌ Bekor qilish")],
        ],
        resize_keyboard=True,
    )


def get_contact_request_keyboard() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="📱 Kontaktni yuborish", request_contact=True)],
            [KeyboardButton(text="⏭ O‘tkazib yuborish")],
            [KeyboardButton(text="❌ Bekor qilish")],
        ],
        resize_keyboard=True,
    )


def get_gender_keyboard() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        keyboard=[
            [
                KeyboardButton(text="👨 O‘g‘il bola"),
                KeyboardButton(text="👩 Qiz bola"),
            ],
            [KeyboardButton(text="❌ Bekor qilish")],
        ],
        resize_keyboard=True,
    )


def get_regions_keyboard() -> ReplyKeyboardMarkup:
    rows = []
    for i in range(0, len(REGIONS), 2):
        row = [KeyboardButton(text=REGIONS[i])]
        if i + 1 < len(REGIONS):
            row.append(KeyboardButton(text=REGIONS[i + 1]))
        rows.append(row)
    rows.append([KeyboardButton(text="❌ Bekor qilish")])
    return ReplyKeyboardMarkup(keyboard=rows, resize_keyboard=True)


def get_photos_upload_keyboard(photo_count: int) -> ReplyKeyboardMarkup:
    rows = []
    if photo_count >= 1:
        rows.append([KeyboardButton(text=f"✅ Tayyor ({photo_count} ta rasm yuklandi)")])
    rows.append([KeyboardButton(text="❌ Bekor qilish")])
    return ReplyKeyboardMarkup(keyboard=rows, resize_keyboard=True)


def get_bio_keyboard() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="⏭ O‘tkazib yuborish")],
            [KeyboardButton(text="❌ Bekor qilish")],
        ],
        resize_keyboard=True,
    )


def get_confirm_profile_keyboard() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="✅ Tasdiqlash")],
            [
                KeyboardButton(text="✏️ Tahrirlash"),
                KeyboardButton(text="❌ Bekor qilish"),
            ],
        ],
        resize_keyboard=True,
    )


def get_search_filter_keyboard() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="🌍 Hamma")],
            [
                KeyboardButton(text="👨 O‘g‘il bola"),
                KeyboardButton(text="👩 Qiz bola"),
            ],
            [KeyboardButton(text="❌ Bekor qilish")],
        ],
        resize_keyboard=True,
    )


def get_card_action_keyboard() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        keyboard=[
            [
                KeyboardButton(text="❤️ Like"),
                KeyboardButton(text="👎 Like emas"),
            ],
            [
                KeyboardButton(text="💌 Shaxsiy chat"),
                KeyboardButton(text="🚩 Shikoyat"),
            ],
            [
                KeyboardButton(text="⏭ Keyingisi"),
                KeyboardButton(text="⬅️ Orqaga"),
            ],
        ],
        resize_keyboard=True,
    )


def get_active_chat_keyboard() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="❌ Suhbatni tugatish")],
            [
                KeyboardButton(text="🚫 Foydalanuvchini bloklash"),
                KeyboardButton(text="🚩 Shikoyat qilish"),
            ],
            [KeyboardButton(text="⬅️ Menyu")],
        ],
        resize_keyboard=True,
        is_persistent=True,
    )


def get_profile_edit_menu_keyboard(is_hidden: bool, allow_chat: bool) -> ReplyKeyboardMarkup:
    hide_text = "👁️‍🗨️ Anketani ko'rsatish" if is_hidden else "🙈 Anketani yashirish"
    chat_text = "🔕 Shaxsiy chatni yopish" if allow_chat else "🔔 Shaxsiy chatga ruxsat berish"

    return ReplyKeyboardMarkup(
        keyboard=[
            [
                KeyboardButton(text="✏️ Ismni o‘zgartirish"),
                KeyboardButton(text="🎂 Yoshni o‘zgartirish"),
            ],
            [
                KeyboardButton(text="👤 Jinsni o‘zgartirish"),
                KeyboardButton(text="🖼 Rasmlarni almashtirish"),
            ],
            [
                KeyboardButton(text="📍 Hududni o‘zgartirish"),
                KeyboardButton(text="📝 Bio tahrirlash"),
            ],
            [KeyboardButton(text=hide_text), KeyboardButton(text=chat_text)],
            [KeyboardButton(text="⬅️ Orqaga")],
        ],
        resize_keyboard=True,
    )


def get_cancel_keyboard() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        keyboard=[[KeyboardButton(text="❌ Bekor qilish")]],
        resize_keyboard=True,
    )


def get_admin_main_menu_keyboard() -> ReplyKeyboardMarkup:
    """
    Admin boshqaruv paneli menyusi:
    - Chiqish faqat '🚪 Admindan chiqish' tugmasi bosilganda amalga oshadi!
    - Barcha tugmalar qulay tartiblangan.
    """
    return ReplyKeyboardMarkup(
        keyboard=[
            [
                KeyboardButton(text="👥 Foydalanuvchilar"),
                KeyboardButton(text="💳 To‘lovlar"),
            ],
            [
                KeyboardButton(text="👑 Adminlar boshqaruvi"),
                KeyboardButton(text="⭐ Premium boshqaruvi"),
            ],
            [
                KeyboardButton(text="📢 Majburiy obuna"),
                KeyboardButton(text="📣 Xabar yuborish"),
            ],
            [
                KeyboardButton(text="📊 Statistika"),
                KeyboardButton(text="📤 Eksport"),
            ],
            [
                KeyboardButton(text="🛡 Moderatsiya"),
                KeyboardButton(text="📝 Bot matnlarini tahrirlash"),
            ],
            [
                KeyboardButton(text="⚙️ Bot sozlamalari"),
                KeyboardButton(text="🧾 Admin faoliyati"),
            ],
            [
                KeyboardButton(text="🔄 Xizmat holati"),
                KeyboardButton(text="🙈 Admin anketasini yashirish/ko'rsatish"),
            ],
            [KeyboardButton(text="🚪 Admindan chiqish")],
        ],
        resize_keyboard=True,
        is_persistent=True,
    )


# ==============================================================================
# 5. MIDDLEWARE VA YORDAMCHI FUNKSIYALAR
# ==============================================================================

USER_FLOOD_TRACKER: Dict[int, float] = {}


async def check_user_throttling(user_id: int, rate_limit: float = 0.3) -> bool:
    now = asyncio.get_event_loop().time()
    last_time = USER_FLOOD_TRACKER.get(user_id, 0)
    if now - last_time < rate_limit:
        return True
    USER_FLOOD_TRACKER[user_id] = now
    return False


async def check_mandatory_subscription(bot: Bot, user_id: int) -> List[Dict]:
    if await db.is_any_admin(user_id):
        return []

    channels = await db.get_required_channels(only_active=True)
    unsubscribed = []

    for ch in channels:
        target = ch["channel_id"] or ch["username"]
        if not target:
            continue
        try:
            member = await bot.get_chat_member(chat_id=target, user_id=user_id)
            if member.status in ["left", "kicked"]:
                unsubscribed.append(dict(ch))
        except Exception as e:
            logger.warning(f"Kanal obunasini tekshirishda xato (Kanal: {target}, User: {user_id}): {e}")
            continue

    return unsubscribed


def build_subscription_required_markup(channels: List[Dict]) -> InlineKeyboardMarkup:
    buttons = []
    for ch in channels:
        link = ch.get("invite_link") or f"https://t.me/{ch.get('username', '')}"
        buttons.append([InlineKeyboardButton(text=f"📢 {ch['title']}", url=link)])
    buttons.append([InlineKeyboardButton(text="🔄 Obunani tekshirish", callback_data="check_subscription_now")])
    return InlineKeyboardMarkup(inline_keyboard=buttons)


# ==============================================================================
# 6. ROUTER VA ASOSIY EVENT HANDLERLAR
# ==============================================================================

router = Router()


@router.errors()
async def error_handler(event: ErrorEvent):
    logger.error(f"Kutilmagan xatolik yuz berdi: {event.exception}", exc_info=True)
    try:
        if event.update.message:
            await event.update.message.answer(
                "⚠️ Kechirasiz, tizimda kutilmagan xatolik yuz berdi. Iltimos, /start buyrug'ini bosing."
            )
    except Exception:
        pass


# ------------------ UNIVERSAL CANCEL / EXIT HANDLER ------------------


@router.message(
    StateFilter("*"),
    F.text.in_(["❌ Bekor qilish", "⬅️ Orqaga", "⬅️ Asosiy menyu", "⬅️ Menyu", "⬅️ Admin menyusi", "🚪 Chiqish", "🚪 Admindan chiqish", "/cancel"]),
)
async def handle_universal_cancel(message: Message, state: FSMContext):
    user_id = message.from_user.id
    current_state = await state.get_state()
    is_admin = await db.is_any_admin(user_id)

    # 1. Agar admin 'Admindan chiqish' tugmasini bossagina asosiy foydalanuvchi menyusiga qaytadi!
    if message.text == "🚪 Admindan chiqish":
        await state.clear()
        return await message.answer(
            "Admin panelidan chiqdingiz. Asosiy foydalanuvchi menyusidasiz:",
            reply_markup=get_main_menu_keyboard(is_admin=True),
        )

    # 2. Agar foydalanuvchi admin rejimida bo'lsa (yoki AdminState da bo'lsa), bekor qilish uni ADMIN menyusida saqlaydi!
    if is_admin and (current_state and current_state.startswith("AdminState")):
        await state.set_state(AdminState.in_admin_mode)
        return await message.answer(
            "Amal bekor qilindi. Admin boshqaruv panelidasiz:",
            reply_markup=get_admin_main_menu_keyboard(),
        )

    # 3. Oddiy bekor qilish (Foydalanuvchi menyusi)
    await state.clear()
    await message.answer(
        "Amal bekor qilindi. Asosiy menyudasiz:",
        reply_markup=get_main_menu_keyboard(is_admin=is_admin),
    )


# ------------------ /START VA RO'YXATDAN O'TISH (QOIDALAR ROZILIGI BILAN) ------------------


@router.message(CommandStart())
async def handle_start(message: Message, state: FSMContext, bot: Bot):
    user_id = message.from_user.id
    if await check_user_throttling(user_id):
        return

    user = await db.get_or_create_user(
        user_id=user_id,
        username=message.from_user.username,
        first_name=message.from_user.first_name,
        last_name=message.from_user.last_name,
    )

    is_banned, ban_reason = await db.is_user_banned(user_id)
    if is_banned:
        await message.answer(
            f"🚫 <b>Sizning profilingiz bloklangan!</b>\n\nSabab: {html.escape(ban_reason or 'Qoidalarni buzganlik uchun')}",
            parse_mode=ParseMode.HTML,
            reply_markup=ReplyKeyboardRemove(),
        )
        return

    maintenance = await db.get_setting("maintenance_mode", "0")
    if maintenance == "1" and not await db.is_any_admin(user_id):
        maint_text = await db.get_setting("maintenance_text", DEFAULT_SYSTEM_TEXTS["maintenance_text"]["default"])
        await message.answer(maint_text, parse_mode=ParseMode.HTML)
        return

    unsubscribed = await check_mandatory_subscription(bot, user_id)
    if unsubscribed:
        markup = build_subscription_required_markup(unsubscribed)
        await message.answer(
            "👋 <b>Assalomu alaykum!</b>\n\nBotdan to'liq foydalanish uchun quyidagi kanallarga a'zo bo'ling:",
            parse_mode=ParseMode.HTML,
            reply_markup=markup,
        )
        return

    profile = await db.get_profile(user_id)
    is_admin = await db.is_any_admin(user_id)

    if profile and profile["is_active"] == 1:
        await state.clear()
        await message.answer(
            f"👋 <b>Xush kelibsiz, {html.escape(profile['name'])}!</b>\n\n"
            f"XLM Anonymous Dating Bot asosiy menyusidasiz.\n"
            f"Anketalarni ko'rish uchun pastdagi tugmani bosing:",
            parse_mode=ParseMode.HTML,
            reply_markup=get_main_menu_keyboard(is_admin=is_admin),
        )
        return

    # Ro'yxatdan o'tmagan yangi foydalanuvchi: Avval qoidalar va shaxsiy chat pullikligi aytiladi!
    await state.clear()
    await state.set_state(RegistrationState.waiting_for_terms)
    welcome_text = await db.get_setting("start_welcome", DEFAULT_SYSTEM_TEXTS["start_welcome"]["default"])

    await message.answer(
        welcome_text,
        parse_mode=ParseMode.HTML,
        reply_markup=get_terms_agreement_keyboard(),
    )


@router.message(RegistrationState.waiting_for_terms, F.text == "✅ Qoidalarga roziman")
async def handle_terms_agreed(message: Message, state: FSMContext):
    """Foydalanuvchi qoidalarga rozilik bergandan so'ng yoshi so'raladi."""
    await state.set_state(RegistrationState.waiting_for_age)
    await message.answer(
        "👍 Qoidalarni qabul qilganingiz uchun rahmat!\n\n"
        "Ro'yxatdan o'tish uchun <b>yoshingizni kiriting</b> (Masalan: <code>21</code>):",
        parse_mode=ParseMode.HTML,
        reply_markup=get_cancel_keyboard(),
    )


@router.callback_query(F.data == "check_subscription_now")
async def handle_check_subscription_callback(callback: CallbackQuery, state: FSMContext, bot: Bot):
    user_id = callback.from_user.id
    unsubscribed = await check_mandatory_subscription(bot, user_id)
    if unsubscribed:
        await callback.answer("❌ Siz hali barcha kanallarga a'zo bo'lmadingiz!", show_alert=True)
    else:
        await callback.answer("✅ Obunangiz tasdiqlandi! Rahmat.", show_alert=True)
        try:
            await callback.message.delete()
        except Exception:
            pass

        profile = await db.get_profile(user_id)
        is_admin = await db.is_any_admin(user_id)
        if profile and profile["is_active"] == 1:
            await callback.message.answer(
                "Asosiy menyu:",
                reply_markup=get_main_menu_keyboard(is_admin=is_admin),
            )
        else:
            await state.set_state(RegistrationState.waiting_for_terms)
            welcome_text = await db.get_setting("start_welcome", DEFAULT_SYSTEM_TEXTS["start_welcome"]["default"])
            await callback.message.answer(welcome_text, parse_mode=ParseMode.HTML, reply_markup=get_terms_agreement_keyboard())


@router.message(RegistrationState.waiting_for_age)
async def process_age(message: Message, state: FSMContext):
    text = (message.text or "").strip()
    age = None
    if text.isdigit():
        age = int(text)
    else:
        date_match = re.match(r"^(\d{2})[./-](\d{2})[./-](\d{4})$", text)
        if date_match:
            try:
                d, m, y = map(int, date_match.groups())
                b_date = datetime(y, m, d)
                today = datetime.today()
                age = today.year - b_date.year - ((today.month, today.day) < (b_date.month, b_date.day))
            except Exception:
                age = None

    if age is None or age < 1 or age > 100:
        await message.answer("⚠️ 18 dan 99 gacha bo'lgan son kiriting:")
        return

    if age < 18:
        await state.clear()
        underage_msg = await db.get_setting("underage_text", DEFAULT_SYSTEM_TEXTS["underage_text"]["default"])
        await message.answer(
            underage_msg,
            parse_mode=ParseMode.HTML,
            reply_markup=ReplyKeyboardRemove(),
        )
        return

    await state.update_data(age=age)
    await state.set_state(RegistrationState.waiting_for_name)
    await message.answer("Ismingizni kiriting:", reply_markup=get_cancel_keyboard())


@router.message(RegistrationState.waiting_for_name)
async def process_name(message: Message, state: FSMContext):
    name = (message.text or "").strip()
    if len(name) < 2 or len(name) > 30:
        await message.answer("Ism 2 dan 30 tagacha harfdan iborat bo'lsin:")
        return

    await state.update_data(name=name)
    await state.set_state(RegistrationState.waiting_for_contact)
    await message.answer(
        "📱 Telefon raqamingizni yuboring (ixtiyoriy):\n"
        "<i>Eslatma: Raqamingiz boshqa foydalanuvchilarga aslo ko'rsatilmaydi.</i>",
        parse_mode=ParseMode.HTML,
        reply_markup=get_contact_request_keyboard(),
    )


@router.message(RegistrationState.waiting_for_contact)
async def process_contact(message: Message, state: FSMContext):
    phone = None
    if message.contact:
        phone = message.contact.phone_number
        await db.update_user_phone(message.from_user.id, phone)
    elif message.text and message.text.startswith("+") and len(message.text) >= 9:
        phone = message.text.strip()
        await db.update_user_phone(message.from_user.id, phone)

    await state.update_data(phone=phone)
    await state.set_state(RegistrationState.waiting_for_gender)
    await message.answer("Jinsingizni tanlang:", reply_markup=get_gender_keyboard())


@router.message(RegistrationState.waiting_for_gender)
async def process_gender(message: Message, state: FSMContext):
    text = (message.text or "").strip()
    if text == "👨 O‘g‘il bola":
        gender = "male"
    elif text == "👩 Qiz bola":
        gender = "female"
    else:
        await message.answer("Tugmalardan birini tanlang:", reply_markup=get_gender_keyboard())
        return

    await state.update_data(gender=gender, photos=[])
    await state.set_state(RegistrationState.waiting_for_photos)
    await message.answer(
        "📸 <b>Suratingizni yuboring:</b>\n"
        "• Kamida 1 ta rasm yuklashingiz kifoya (ko'pi bilan 3 ta).\n"
        "• 1-rasmdan so'ng darhol '✅ Tayyor' tugmasini bosishingiz mumkin.",
        parse_mode=ParseMode.HTML,
        reply_markup=get_photos_upload_keyboard(0),
    )


@router.message(RegistrationState.waiting_for_photos)
async def process_photos(message: Message, state: FSMContext):
    text = (message.text or "").strip()
    data = await state.get_data()
    photos = data.get("photos", [])

    if text.startswith("✅ Tayyor") and len(photos) >= 1:
        await state.set_state(RegistrationState.waiting_for_region)
        await message.answer("📍 Hududingizni tanlang:", reply_markup=get_regions_keyboard())
        return

    if not message.photo:
        await message.answer("Iltimos, rasm yuboring!")
        return

    photos.append(message.photo[-1].file_id)
    await state.update_data(photos=photos)
    count = len(photos)

    if count == 1:
        await message.answer(
            "👍 Rasm qabul qilindi! Keyingi bosqichga o'tish uchun '✅ Tayyor' tugmasini bosing yoki yana rasm yuboring (ko'pi bilan 3 ta):",
            reply_markup=get_photos_upload_keyboard(count),
        )
    elif count == 2:
        await message.answer(
            "👍 2-rasm qabul qilindi. '✅ Tayyor' tugmasini bosing:",
            reply_markup=get_photos_upload_keyboard(count),
        )
    elif count >= 3:
        await state.set_state(RegistrationState.waiting_for_region)
        await message.answer("📍 Hududingizni tanlang:", reply_markup=get_regions_keyboard())


@router.message(RegistrationState.waiting_for_region)
async def process_region(message: Message, state: FSMContext):
    region = (message.text or "").strip()
    if region not in REGIONS:
        await message.answer("Iltimos, ro'yxatdagi hududlardan birini tanlang:", reply_markup=get_regions_keyboard())
        return

    await state.update_data(region=region)
    await state.set_state(RegistrationState.waiting_for_bio)
    await message.answer(
        "📝 O'zingiz haqingizda qisqacha yozing (Bio) yoki '⏭ O‘tkazib yuborish' tugmasini bosing:",
        reply_markup=get_bio_keyboard(),
    )


@router.message(RegistrationState.waiting_for_bio)
async def process_bio(message: Message, state: FSMContext):
    text = (message.text or "").strip()
    bio = None if text == "⏭ O‘tkazib yuborish" else text
    if bio and len(bio) > 300:
        await message.answer("Bio 300 belgidan oshmasligi kerak. Qisqaroq yozing:")
        return

    await state.update_data(bio=bio)
    data = await state.get_data()

    gender_name = "👨 O‘g‘il bola" if data["gender"] == "male" else "👩 Qiz bola"
    caption = (
        f"📋 <b>Anketangiz:</b>\n\n"
        f"👤 <b>Ism:</b> {html.escape(data['name'])}\n"
        f"🎂 <b>Yosh:</b> {data['age']}\n"
        f"👤 <b>Jins:</b> {gender_name}\n"
        f"📍 <b>Hudud:</b> {data['region']}\n"
        f"📝 <b>Bio:</b> {html.escape(data.get('bio') or 'Ko‘rsatilmagan')}\n\n"
        f"Anketa to'g'ri bo'lsa, '✅ Tasdiqlash' tugmasini bosing:"
    )

    photos = data.get("photos", [])
    if photos:
        await message.answer_photo(
            photo=photos[0],
            caption=caption,
            parse_mode=ParseMode.HTML,
            reply_markup=get_confirm_profile_keyboard(),
        )
    else:
        await message.answer(caption, parse_mode=ParseMode.HTML, reply_markup=get_confirm_profile_keyboard())

    await state.set_state(RegistrationState.confirm_profile)


@router.message(RegistrationState.confirm_profile)
async def process_confirm_profile(message: Message, state: FSMContext):
    text = (message.text or "").strip()
    data = await state.get_data()
    user_id = message.from_user.id

    if text == "✅ Tasdiqlash":
        await db.save_profile(
            user_id=user_id,
            name=data["name"],
            age=data["age"],
            gender=data["gender"],
            region=data["region"],
            bio=data.get("bio"),
            photos=data.get("photos", []),
        )
        await state.clear()
        is_admin = await db.is_any_admin(user_id)
        await message.answer(
            "🎉 <b>Tabriklaymiz! Anketangiz muvaffaqiyatli saqlandi!</b>\n\n"
            "Endi '🔎 Anketalarni ko‘rish' tugmasi orqali suhbatdosh izlashingiz mumkin.",
            parse_mode=ParseMode.HTML,
            reply_markup=get_main_menu_keyboard(is_admin=is_admin),
        )
    elif text == "✏️ Tahrirlash":
        await state.clear()
        await state.set_state(RegistrationState.waiting_for_name)
        await message.answer("Ismingizni qayta kiriting:", reply_markup=get_cancel_keyboard())
    else:
        await state.clear()
        is_admin = await db.is_any_admin(user_id)
        await message.answer("Bekor qilindi.", reply_markup=get_main_menu_keyboard(is_admin=is_admin))


# ==============================================================================
# 7. ANKETALARNI KO'RISH (PREMIUM FILTER CHECK BILAN)
# ==============================================================================


@router.message(F.text.in_(["🔎 Anketalarni ko‘rish", "🔎 Odam qidirish"]))
async def start_search(message: Message, state: FSMContext, bot: Bot):
    user_id = message.from_user.id
    if await check_user_throttling(user_id):
        return

    profile = await db.get_profile(user_id)
    if not profile or profile["is_active"] == 0:
        await message.answer("⚠️ Boshqalarning anketalarini ko'rish uchun avval o'z anketangizni to'ldiring. /start buyrug'ini bosing.")
        return

    # Yashirilgan anketaga qidiruv taqiqlanadi!
    if profile["is_hidden"] == 1:
        warn = await db.get_setting("hidden_profile_warning", DEFAULT_SYSTEM_TEXTS["hidden_profile_warning"]["default"])
        await message.answer(warn, parse_mode=ParseMode.HTML)
        return

    unsubscribed = await check_mandatory_subscription(bot, user_id)
    if unsubscribed:
        markup = build_subscription_required_markup(unsubscribed)
        await message.answer("Qidiruv tizimidan foydalanish uchun kanallarimizga obuna bo'ling:", reply_markup=markup)
        return

    await state.set_state(SearchState.selecting_filter)
    await message.answer("Qaysi jinsdagi insonlarni qidirmoqchisiz?", reply_markup=get_search_filter_keyboard())


@router.message(SearchState.selecting_filter)
async def process_search_filter(message: Message, state: FSMContext, bot: Bot):
    text = (message.text or "").strip()
    user_id = message.from_user.id
    is_premium, _ = await db.is_premium_active(user_id)

    # PREMIUM TEKSHIRUVI: Agar Prime bo'lmasa, faqat 'Hamma' ni ko'ra oladi!
    if text in ["👨 O‘g‘il bola", "👩 Qiz bola"] and not is_premium:
        non_prime_msg = await db.get_setting("non_prime_filter_warning", DEFAULT_SYSTEM_TEXTS["non_prime_filter_warning"]["default"])
        await message.answer(
            non_prime_msg,
            parse_mode=ParseMode.HTML,
            reply_markup=InlineKeyboardMarkup(
                inline_keyboard=[[InlineKeyboardButton(text="⭐ Premium olish", callback_data="goto_premium_plans")]]
            ),
        )
        return

    gender_filter = "all"
    if text == "👨 O‘g‘il bola":
        gender_filter = "male"
    elif text == "👩 Qiz bola":
        gender_filter = "female"
    elif text == "🌍 Hamma":
        gender_filter = "all"
    else:
        await message.answer("Tugmalardan birini tanlang:", reply_markup=get_search_filter_keyboard())
        return

    await state.update_data(gender_filter=gender_filter)
    await state.set_state(SearchState.browsing)
    await show_next_profile_card(message, state, bot)


async def show_next_profile_card(message: Message, state: FSMContext, bot: Bot):
    user_id = message.from_user.id
    data = await state.get_data()
    gender_filter = data.get("gender_filter", "all")
    was_in_rebrowse = data.get("in_rebrowse", False)

    candidate = await db.get_next_candidate_profile(user_id, gender_filter=gender_filter)
    if not candidate:
        is_admin = await db.is_any_admin(user_id)
        await state.clear()
        await message.answer(
            "😔 <b>Barcha anketalar koʻrib boʻlindi.</b>\n\n"
            "Yangi foydalanuvchilar roʻxatdan oʻtgach yana urinib koʻring!",
            parse_mode=ParseMode.HTML,
            reply_markup=get_main_menu_keyboard(is_admin=is_admin),
        )
        return

    target_user_id = candidate["user_id"]

    # Avval dislike qilingan (qayta ko'rilayotgan) anketami?
    is_previously_seen = False
    async with await db.get_connection() as _conn:
        async with _conn.execute(
            "SELECT id FROM likes WHERE from_user_id = ? AND to_user_id = ? AND is_like = 0",
            (user_id, target_user_id),
        ) as _cur:
            is_previously_seen = await _cur.fetchone() is not None

    if is_previously_seen and not was_in_rebrowse:
        await state.update_data(in_rebrowse=True)
        await message.answer(
            "🔄 <b>Yangi anketalar tugadi!</b>\n\n"
            "Ilgari oʻkazib yuborgan anketalaringizni qayta koʻrsatmoqdamiz. "
            "Bu safar yoqimli kimnidir topib olasiz! 😊",
            parse_mode=ParseMode.HTML,
        )

    await state.update_data(current_candidate_id=target_user_id)

    photos = await db.get_profile_photos(target_user_id)
    is_cand_premium, _ = await db.is_premium_active(target_user_id)

    premium_tag = "⭐ <b>PREMIUM FOYDALANUVCHI</b>\n" if is_cand_premium else ""
    rebrowse_tag = "🔄 <i>Avval koʻrilgan anketa</i>\n" if is_previously_seen else ""
    gender_name = "👨 Oʻgʻil bola" if candidate["gender"] == "male" else "👩 Qiz bola"

    caption = (
        f"{premium_tag}"
        f"{rebrowse_tag}"
        f"✨ <b>{html.escape(candidate['name'])}, {candidate['age']}</b>\n"
        f"📍 {candidate['region']}\n"
        f"👤 {gender_name}\n\n"
        f"📝 <b>Bio:</b> {html.escape(candidate['bio'] or 'Koʻrsatilmagan')}"
    )

    markup = get_card_action_keyboard()

    if photos:
        await message.answer_photo(
            photo=photos[0]["file_id"],
            caption=caption,
            parse_mode=ParseMode.HTML,
            reply_markup=markup,
        )
    else:
        await message.answer(caption, parse_mode=ParseMode.HTML, reply_markup=markup)


@router.message(SearchState.browsing, F.text == "❤️ Like")
async def handle_like_button(message: Message, state: FSMContext, bot: Bot):
    user_id = message.from_user.id
    if await check_user_throttling(user_id):
        return

    data = await state.get_data()
    target_id = data.get("current_candidate_id")
    if not target_id:
        return await show_next_profile_card(message, state, bot)

    await db.add_like_or_dislike(user_id, target_id, is_like=1)
    _, is_mutual = await db.check_mutual_like(user_id, target_id)

    if is_mutual:
        target_profile = await db.get_profile(target_id)
        match_msg = await db.get_setting("match_text", DEFAULT_SYSTEM_TEXTS["match_text"]["default"])

        await message.answer(
            f"{match_msg}\n\nSiz va <b>{html.escape(target_profile['name'])}</b> bir-biringizga like bosdingiz!",
            parse_mode=ParseMode.HTML,
        )
        try:
            my_profile = await db.get_profile(user_id)
            await bot.send_message(
                chat_id=target_id,
                text=f"{match_msg}\n\nSiz va <b>{html.escape(my_profile['name'])}</b> bir-biringizga like bosdingiz!",
                parse_mode=ParseMode.HTML,
            )
        except Exception:
            pass

    await show_next_profile_card(message, state, bot)


@router.message(SearchState.browsing, F.text.in_(["👎 Like emas", "⏭ Keyingisi"]))
async def handle_dislike_button(message: Message, state: FSMContext, bot: Bot):
    user_id = message.from_user.id
    if await check_user_throttling(user_id):
        return

    data = await state.get_data()
    target_id = data.get("current_candidate_id")
    if target_id and message.text == "👎 Like emas":
        await db.add_like_or_dislike(user_id, target_id, is_like=0)

    await show_next_profile_card(message, state, bot)


@router.message(SearchState.browsing, F.text == "💌 Shaxsiy chat")
async def handle_request_private_chat(message: Message, state: FSMContext, bot: Bot):
    user_id = message.from_user.id
    if await check_user_throttling(user_id):
        return

    data = await state.get_data()
    target_id = data.get("current_candidate_id")
    if not target_id:
        return await show_next_profile_card(message, state, bot)

    is_premium, _ = await db.is_premium_active(user_id)
    if not is_premium:
        await message.answer(
            "⭐ <b>Anonim shaxsiy chat faqat PREMIUM foydalanuvchilar uchun!</b>\n\n"
            "Oʻzingiz yoqtirgan inson bilan bevosita shaxsiy suhbat ochish uchun Premium obunani xarid qiling:",
            parse_mode=ParseMode.HTML,
            reply_markup=InlineKeyboardMarkup(
                inline_keyboard=[[InlineKeyboardButton(text="⭐ Premium tariflarini koʻring", callback_data="goto_premium_plans")]]
            ),
        )
        return

    target_profile = await db.get_profile(target_id)
    if not target_profile or await db.is_blocked(user_id, target_id):
        await message.answer("Kechirasiz, bu foydalanuvchi bilan shaxsiy chatni boshlash imkoni yoʻq.")
        return

    # Target foydalanuvchining username ini topamiz
    target_user_row = await db.get_or_create_user(target_id, None, None, None)
    target_username = None
    if target_user_row:
        target_username = target_user_row["username"]

    if target_username:
        # username mavjud — to'g'ridan-to'g'ri t.me/username ga yo'naltir
        direct_link = f"https://t.me/{target_username}"
        await message.answer(
            f"💌 <b>{html.escape(target_profile['name'])} bilan bevosita suhbat:</b>\n\n"
            f"🔗 Quyidagi tugmani bosing va toʻgʻridan-toʻgʻri uning Telegram chatiga oʻting:\n\n"
            f"⚠️ <i>Shaxsiy maʻumotlaringiz ularga koʻrinishi mumkin. Ehtiyot boʻling!</i>",
            parse_mode=ParseMode.HTML,
            reply_markup=InlineKeyboardMarkup(
                inline_keyboard=[
                    [InlineKeyboardButton(text=f"💬 {html.escape(target_profile['name'])} bilan suhbat", url=direct_link)],
                    [InlineKeyboardButton(text="⭐ Anonim suhbat (bot orqali)", callback_data=f"anon_chat_req:{target_id}")],
                ]
            ),
        )
    else:
        # Username yoʻq — target_id orqali tg://user?id= havolasi
        direct_link = f"tg://user?id={target_id}"
        await message.answer(
            f"💌 <b>{html.escape(target_profile['name'])} bilan bevosita suhbat:</b>\n\n"
            f"🔗 Quyidagi tugmani bosing va uning Telegram chatiga oʻting.\n"
            f"⚠️ <i>Agar tugma ishlamasa, bu foydalanuvchining profili yashirin boʻlishi mumkin.</i>",
            parse_mode=ParseMode.HTML,
            reply_markup=InlineKeyboardMarkup(
                inline_keyboard=[
                    [InlineKeyboardButton(text=f"💬 {html.escape(target_profile['name'])} bilan suhbat", url=direct_link)],
                    [InlineKeyboardButton(text="⭐ Anonim suhbat (bot orqali)", callback_data=f"anon_chat_req:{target_id}")],
                ]
            ),
        )

    await show_next_profile_card(message, state, bot)


@router.callback_query(F.data.startswith("anon_chat_req:"))
async def handle_anon_chat_request_cb(callback: CallbackQuery, state: FSMContext, bot: Bot):
    """Anonim suhbat so'rovi (bot orqali) — shaxsiy chat handler'dan chaqiriladi."""
    await callback.answer()
    target_id = int(callback.data.split(":")[1])
    user_id = callback.from_user.id

    target_profile = await db.get_profile(target_id)
    if not target_profile or target_profile["allow_chat_requests"] == 0 or await db.is_blocked(user_id, target_id):
        await callback.message.answer("Kechirasiz, bu foydalanuvchi bilan shaxsiy chatni boshlash imkoni yoʻq.")
        return

    existing_chat = await db.get_active_chat(target_id)
    if existing_chat:
        await callback.message.answer("Bu foydalanuvchi hozirda boshqa suhbatda. Birozdan soʻng urinib koʻring.")
        return

    request_id = await db.create_chat_request(user_id, target_id)
    my_profile = await db.get_profile(user_id)

    request_markup = InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(text="✅ Suhbatni qabul qilish", callback_data=f"chat_accept:{request_id}"),
                InlineKeyboardButton(text="❌ Rad etish", callback_data=f"chat_reject:{request_id}"),
            ]
        ]
    )

    try:
        await bot.send_message(
            chat_id=target_id,
            text=f"💌 <b>Sizga yangi anonim shaxsiy suhbat soʻrov keldi!</b>\n\n"
            f"👤 <b>Yuboruvchi:</b> {html.escape(my_profile['name'])}, {my_profile['age']} yosh\n"
            f"📍 <b>Hudud:</b> {my_profile['region']}\n\nSuhbatlashishga rozimisiz?",
            parse_mode=ParseMode.HTML,
            reply_markup=request_markup,
        )
        await callback.message.answer(f"✅ {html.escape(target_profile['name'])} ga anonim suhbat soʻrov yuborildi!")
    except Exception:
        await callback.message.answer("Kechirasiz, bu foydalanuvchi bilan shaxsiy chatni boshlash imkoni yoʻq.")


@router.callback_query(F.data.startswith("chat_accept:"))
async def handle_chat_accept(callback: CallbackQuery, state: FSMContext, bot: Bot):
    request_id = int(callback.data.split(":")[1])
    chat = await db.accept_chat_request(request_id)
    if not chat:
        await callback.answer("Bu so'rov muddati o'tgan yoki bekor qilingan.", show_alert=True)
        return

    await callback.answer("Suhbat boshlandi!")
    try:
        await callback.message.delete()
    except Exception:
        pass

    chat_text = (
        "🎉 <b>Anonim shaxsiy suhbat boshlandi!</b>\n\n"
        "🔒 Bir-biringizning shaxsiy ma'lumotlaringiz ko'rinmaydi. Suhbatni yakunlash uchun '❌ Suhbatni tugatish' tugmasini bosing."
    )
    markup = get_active_chat_keyboard()

    for uid in [chat["user1_id"], chat["user2_id"]]:
        try:
            await bot.send_message(chat_id=uid, text=chat_text, parse_mode=ParseMode.HTML, reply_markup=markup)
        except Exception:
            pass


@router.callback_query(F.data.startswith("chat_reject:"))
async def handle_chat_reject(callback: CallbackQuery):
    request_id = int(callback.data.split(":")[1])
    await db.reject_chat_request(request_id)
    await callback.answer("So'rov rad etildi.")
    try:
        await callback.message.edit_text("❌ Shaxsiy chat so'rovi rad etildi.")
    except Exception:
        pass


@router.message(F.text == "❌ Suhbatni tugatish")
async def handle_end_chat(message: Message, state: FSMContext, bot: Bot):
    user_id = message.from_user.id
    active_chat = await db.get_active_chat(user_id)
    if not active_chat:
        is_admin = await db.is_any_admin(user_id)
        await message.answer("Siz hozirda faol suhbatda emassiz.", reply_markup=get_main_menu_keyboard(is_admin=is_admin))
        return

    chat_id = active_chat["id"]
    participants = await db.end_active_chat(chat_id, ended_by=user_id)
    if participants:
        u1, u2 = participants
        partner_id = u2 if user_id == u1 else u1

        is_admin_me = await db.is_any_admin(user_id)
        is_admin_partner = await db.is_any_admin(partner_id)

        await message.answer("Suhbat yakunlandi.", reply_markup=get_main_menu_keyboard(is_admin=is_admin_me))
        try:
            await bot.send_message(chat_id=partner_id, text="❌ Suhbatdoshingiz suhbatni yakunladi.", reply_markup=get_main_menu_keyboard(is_admin=is_admin_partner))
        except Exception:
            pass


@router.message(F.text == "🚫 Foydalanuvchini bloklash")
async def handle_block_chat_partner(message: Message, state: FSMContext, bot: Bot):
    user_id = message.from_user.id
    active_chat = await db.get_active_chat(user_id)
    if not active_chat:
        return

    partner_id = active_chat["user2_id"] if user_id == active_chat["user1_id"] else active_chat["user1_id"]
    await db.block_user(user_id, partner_id, reason="Chat davomida bloklandi")
    is_admin = await db.is_any_admin(user_id)
    await message.answer("🚫 Suhbatdosh bloklandi va suhbat yakunlandi.", reply_markup=get_main_menu_keyboard(is_admin=is_admin))
    try:
        is_admin_p = await db.is_any_admin(partner_id)
        await bot.send_message(chat_id=partner_id, text="❌ Suhbat yakunlandi.", reply_markup=get_main_menu_keyboard(is_admin=is_admin_p))
    except Exception:
        pass


@router.message(SearchState.browsing, F.text == "🚩 Shikoyat")
@router.message(F.text == "🚩 Shikoyat qilish")
async def handle_report_start(message: Message, state: FSMContext):
    user_id = message.from_user.id
    data = await state.get_data()
    target_id = data.get("current_candidate_id")

    if not target_id:
        active_chat = await db.get_active_chat(user_id)
        if active_chat:
            target_id = active_chat["user2_id"] if user_id == active_chat["user1_id"] else active_chat["user1_id"]

    if not target_id:
        await message.answer("Shikoyat qilinadigan foydalanuvchi topilmadi.")
        return

    await state.set_state(ReportState.waiting_for_reason)
    await state.update_data(reported_user_id=target_id)

    reasons_markup = ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="Soxta profil"), KeyboardButton(text="Spam")],
            [KeyboardButton(text="Haqorat yoki bezovta qilish"), KeyboardButton(text="Nomaqbul rasm")],
            [KeyboardButton(text="Firibgarlik"), KeyboardButton(text="Boshqa sabab")],
            [KeyboardButton(text="❌ Bekor qilish")],
        ],
        resize_keyboard=True,
    )
    await message.answer("Shikoyat sababini tanlang:", reply_markup=reasons_markup)


@router.message(ReportState.waiting_for_reason)
async def process_report_reason(message: Message, state: FSMContext, bot: Bot):
    text = (message.text or "").strip()
    data = await state.get_data()
    reported_user_id = data.get("reported_user_id")

    report_id = await db.create_report(
        reporter_id=message.from_user.id,
        reported_user_id=reported_user_id,
        reason=text,
        details="Foydalanuvchi arizasi",
    )
    await db.block_user(message.from_user.id, reported_user_id, reason=f"Shikoyat: {text}")

    await message.answer(f"✅ Shikoyatingiz (#{report_id}) moderatsiyaga yuborildi.")
    is_admin = await db.is_any_admin(message.from_user.id)
    await state.clear()
    await message.answer("Asosiy menyu:", reply_markup=get_main_menu_keyboard(is_admin=is_admin))


# ------------------ 👤 ANKETAM VA TAHRIRLASH ------------------


@router.message(F.text == "👤 Anketam")
async def handle_my_profile(message: Message, state: FSMContext):
    user_id = message.from_user.id
    profile = await db.get_profile(user_id)
    if not profile or profile["is_active"] == 0:
        await message.answer("Siz hali anketa yaratmagansiz. /start buyrug'ini bosing.")
        return

    photos = await db.get_profile_photos(user_id)
    is_premium, end_date = await db.is_premium_active(user_id)
    premium_badge = f"⭐ <b>PREMIUM</b> ({end_date} gacha)" if is_premium else "Oddiy"
    gender_name = "👨 O‘g‘il bola" if profile["gender"] == "male" else "👩 Qiz bola"
    status_text = "🙈 Yashirilgan" if profile["is_hidden"] else "🟢 Qidiruvda faol"
    chat_perm_text = "✅ Ochiq" if profile["allow_chat_requests"] else "❌ Yopiq"

    caption = (
        f"👤 <b>SIZNING ANKETANGIZ:</b>\n\n"
        f"✨ <b>Ism:</b> {html.escape(profile['name'])}\n"
        f"🎂 <b>Yosh:</b> {profile['age']}\n"
        f"👤 <b>Jins:</b> {gender_name}\n"
        f"📍 <b>Hudud:</b> {profile['region']}\n"
        f"📝 <b>Bio:</b> {html.escape(profile['bio'] or 'Ko‘rsatilmagan')}\n\n"
        f"💎 <b>Status:</b> {premium_badge}\n"
        f"👁 <b>Holati:</b> {status_text}\n"
        f"💌 <b>Shaxsiy chat so'rovlari:</b> {chat_perm_text}\n\n"
        f"Tahrirlash uchun quyidagi tugmalardan foydalaning:"
    )

    markup = get_profile_edit_menu_keyboard(is_hidden=bool(profile["is_hidden"]), allow_chat=bool(profile["allow_chat_requests"]))

    if photos:
        await message.answer_photo(photo=photos[0]["file_id"], caption=caption, parse_mode=ParseMode.HTML, reply_markup=markup)
    else:
        await message.answer(caption, parse_mode=ParseMode.HTML, reply_markup=markup)


@router.message(F.text.in_(["👁️‍🗨️ Anketani ko'rsatish", "🙈 Anketani yashirish"]))
async def toggle_profile_hidden(message: Message):
    user_id = message.from_user.id
    profile = await db.get_profile(user_id)
    if not profile:
        return

    new_val = 0 if profile["is_hidden"] == 1 else 1
    await db.update_profile_field(user_id, "is_hidden", new_val)
    status_msg = "🙈 Anketangiz yashirildi. Sizni boshqalar ko'ra olmaydi." if new_val else "👁 Anketangiz faollashtirildi!"
    await message.answer(status_msg, reply_markup=get_profile_edit_menu_keyboard(is_hidden=bool(new_val), allow_chat=bool(profile["allow_chat_requests"])))


@router.message(F.text.in_(["🔕 Shaxsiy chatni yopish", "🔔 Shaxsiy chatga ruxsat berish"]))
async def toggle_chat_requests(message: Message):
    user_id = message.from_user.id
    profile = await db.get_profile(user_id)
    if not profile:
        return

    new_val = 0 if profile["allow_chat_requests"] == 1 else 1
    await db.update_profile_field(user_id, "allow_chat_requests", new_val)
    status_msg = "🔔 Shaxsiy chat so'rovlari yoqildi!" if new_val else "🔕 Shaxsiy chat so'rovlari o'chirildi."
    await message.answer(status_msg, reply_markup=get_profile_edit_menu_keyboard(is_hidden=bool(profile["is_hidden"]), allow_chat=bool(new_val)))


@router.message(F.text == "✏️ Ismni o‘zgartirish")
async def edit_name_start(message: Message, state: FSMContext):
    await state.set_state(ProfileEditState.edit_name)
    await message.answer("Yangi ismingizni kiriting:", reply_markup=get_cancel_keyboard())


@router.message(ProfileEditState.edit_name)
async def edit_name_finish(message: Message, state: FSMContext):
    name = (message.text or "").strip()
    if len(name) < 2 or len(name) > 30:
        await message.answer("Ism 2 dan 30 tagacha harfdan iborat bo'lsin:")
        return

    await db.update_profile_field(message.from_user.id, "name", name)
    await state.clear()
    await message.answer("✅ Ismingiz yangilandi!")
    await handle_my_profile(message, state)


@router.message(F.text == "🎂 Yoshni o‘zgartirish")
async def edit_age_start(message: Message, state: FSMContext):
    await state.set_state(ProfileEditState.edit_age)
    await message.answer("Yangi yoshingizni kiriting (18+):", reply_markup=get_cancel_keyboard())


@router.message(ProfileEditState.edit_age)
async def edit_age_finish(message: Message, state: FSMContext):
    text = (message.text or "").strip()
    if not text.isdigit() or int(text) < 18 or int(text) > 99:
        await message.answer("18 dan 99 gacha bo'lgan to'g'ri yoshingizni kiriting:")
        return

    await db.update_profile_field(message.from_user.id, "age", int(text))
    await state.clear()
    await message.answer("✅ Yoshingiz yangilandi!")
    await handle_my_profile(message, state)


@router.message(F.text == "👤 Jinsni o‘zgartirish")
async def edit_gender_start(message: Message, state: FSMContext):
    await state.set_state(ProfileEditState.edit_gender)
    await message.answer("Jinsingizni tanlang:", reply_markup=get_gender_keyboard())


@router.message(ProfileEditState.edit_gender)
async def edit_gender_finish(message: Message, state: FSMContext):
    text = (message.text or "").strip()
    g = "male" if text == "👨 O‘g‘il bola" else ("female" if text == "👩 Qiz bola" else None)
    if not g:
        await message.answer("Tugmalardan birini bosing:", reply_markup=get_gender_keyboard())
        return

    await db.update_profile_field(message.from_user.id, "gender", g)
    await state.clear()
    await message.answer("✅ Jinsingiz yangilandi!")
    await handle_my_profile(message, state)


@router.message(F.text == "📍 Hududni o‘zgartirish")
async def edit_region_start(message: Message, state: FSMContext):
    await state.set_state(ProfileEditState.edit_region)
    await message.answer("Yangi hududni tanlang:", reply_markup=get_regions_keyboard())


@router.message(ProfileEditState.edit_region)
async def edit_region_finish(message: Message, state: FSMContext):
    reg = (message.text or "").strip()
    if reg not in REGIONS:
        await message.answer("Ro'yxatdan tanlang:", reply_markup=get_regions_keyboard())
        return

    await db.update_profile_field(message.from_user.id, "region", reg)
    await state.clear()
    await message.answer("✅ Hududingiz yangilandi!")
    await handle_my_profile(message, state)


@router.message(F.text == "📝 Bio tahrirlash")
async def edit_bio_start(message: Message, state: FSMContext):
    await state.set_state(ProfileEditState.edit_bio)
    await message.answer("Yangi bio yozing (yoki '⏭ O‘tkazib yuborish'):", reply_markup=get_bio_keyboard())


@router.message(ProfileEditState.edit_bio)
async def edit_bio_finish(message: Message, state: FSMContext):
    text = (message.text or "").strip()
    bio = None if text == "⏭ O‘tkazib yuborish" else text
    await db.update_profile_field(message.from_user.id, "bio", bio)
    await state.clear()
    await message.answer("✅ Bio yangilandi!")
    await handle_my_profile(message, state)


@router.message(F.text == "🖼 Rasmlarni almashtirish")
async def edit_photos_start(message: Message, state: FSMContext):
    await state.set_state(ProfileEditState.edit_photos)
    await state.update_data(new_photos=[])
    await message.answer("Yangi suratingizni yuboring (kamida 1 ta, ko'pi bilan 3 ta):", reply_markup=get_photos_upload_keyboard(0))


@router.message(ProfileEditState.edit_photos)
async def edit_photos_finish(message: Message, state: FSMContext):
    text = (message.text or "").strip()
    data = await state.get_data()
    photos = data.get("new_photos", [])

    if text.startswith("✅ Tayyor") and len(photos) >= 1:
        user_id = message.from_user.id
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        async with db._lock:
            async with await db.get_connection() as conn:
                await conn.execute("DELETE FROM profile_photos WHERE user_id = ?", (user_id,))
                for idx, fid in enumerate(photos[:3], 1):
                    await conn.execute("INSERT INTO profile_photos (user_id, file_id, sort_order, created_at) VALUES (?, ?, ?, ?)", (user_id, fid, idx, now_str))
                await conn.commit()
        await state.clear()
        await message.answer("✅ Rasmlaringiz muvaffaqiyatli yangilandi!")
        return await handle_my_profile(message, state)

    if not message.photo:
        await message.answer("Iltimos, rasm yuboring!")
        return

    photos.append(message.photo[-1].file_id)
    await state.update_data(new_photos=photos)
    count = len(photos)

    if count < 3:
        await message.answer(f"👍 Rasm qabul qilindi. '✅ Tayyor' tugmasini bosing:", reply_markup=get_photos_upload_keyboard(count))
    else:
        user_id = message.from_user.id
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        async with db._lock:
            async with await db.get_connection() as conn:
                await conn.execute("DELETE FROM profile_photos WHERE user_id = ?", (user_id,))
                for idx, fid in enumerate(photos[:3], 1):
                    await conn.execute("INSERT INTO profile_photos (user_id, file_id, sort_order, created_at) VALUES (?, ?, ?, ?)", (user_id, fid, idx, now_str))
                await conn.commit()
        await state.clear()
        await message.answer("✅ 3 ta yangi rasm yuklandi!")
        return await handle_my_profile(message, state)


# ==============================================================================
# 8. PREMIUM TARIFLARI VA TO'LOVLAR (QO'LDA TASDIQLASH)
# ==============================================================================


@router.message(F.text == "⭐ Premium olish")
@router.callback_query(F.data == "goto_premium_plans")
async def handle_premium_plans_menu(event: Any, state: FSMContext):
    user_id = event.from_user.id
    plans = await db.get_premium_plans(only_active=True)
    if not plans:
        msg = "Hozirda faol tariflar yo'q."
        return await (event.answer(msg, show_alert=True) if isinstance(event, CallbackQuery) else event.answer(msg))

    is_premium, end_date = await db.is_premium_active(user_id)
    card_number = await db.get_setting("card_number", "8600 0000 0000 0000")
    card_holder = await db.get_setting("card_holder", "XLM DATING SERVICE")

    header = (
        "⭐ <b>XLM PREMIUM IMKONIYATLARI:</b>\n\n"
        "• 💌 Cheksiz anonim shaxsiy suhbatlar boshlash\n"
        "• 👫 O'g'il bola yoki qiz bola bo'yicha alohida saralash\n"
        "• 👑 Profilingizda oltin Premium belgisi\n"
        "• 🚀 Qidiruvda eng oldi o'rinlarda ko'rinish\n\n"
        f"💳 <b>To'lov kartasi:</b> <code>{card_number}</code>\n"
        f"👤 <b>Karta egasi:</b> <code>{card_holder}</code>\n\n"
    )
    if is_premium:
        header += f"<i>Sizda faol Premium mavjud ({end_date} gacha). Yangi to'lov mavjud muddatni uzaytiradi!</i>\n\n"

    buttons = []
    for p in plans:
        btn_text = f"📦 {p['name']} — {p['price_uzs']:,} UZS ({p['duration_days']} kun)"
        buttons.append([InlineKeyboardButton(text=btn_text, callback_data=f"buy_plan:{p['id']}")])

    # To'lov chekini bevosita yuborish tugmasi
    buttons.append([InlineKeyboardButton(text="📄 To‘lov chekini yuborish", callback_data="direct_send_receipt")])

    markup = InlineKeyboardMarkup(inline_keyboard=buttons)
    if isinstance(event, CallbackQuery):
        await event.message.answer(header, parse_mode=ParseMode.HTML, reply_markup=markup)
        await event.answer()
    else:
        await event.answer(header, parse_mode=ParseMode.HTML, reply_markup=markup)


@router.callback_query(F.data.startswith("buy_plan:"))
@router.callback_query(F.data == "direct_send_receipt")
async def handle_select_plan_or_direct_receipt(callback: CallbackQuery, state: FSMContext):
    data = callback.data
    if data == "direct_send_receipt":
        plans = await db.get_premium_plans(only_active=True)
        buttons = []
        for p in plans:
            buttons.append([InlineKeyboardButton(text=f"📦 {p['name']} ({p['price_uzs']:,} UZS)", callback_data=f"buy_plan:{p['id']}")])
        await callback.message.answer("Qaysi tarif uchun to'lov qildingiz? Tanlang:", reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
        await callback.answer()
        return

    plan_id = int(data.split(":")[1])
    plan = await db.get_premium_plan_by_id(plan_id)
    if not plan or plan["is_active"] == 0:
        await callback.answer("Tarif faol emas.", show_alert=True)
        return

    card_number = await db.get_setting("card_number", "8600 0000 0000 0000")
    card_holder = await db.get_setting("card_holder", "XLM DATING SERVICE")
    instructions = await db.get_setting("payment_instructions", DEFAULT_SYSTEM_TEXTS["payment_instructions"]["default"])

    await state.set_state(PaymentState.uploading_receipt)
    await state.update_data(selected_plan_id=plan_id, plan_price=plan["price_uzs"])

    text = (
        f"📦 <b>Tanlangan tarif:</b> {html.escape(plan['name'])}\n"
        f"⏳ <b>Muddat:</b> {plan['duration_days']} kun\n"
        f"💰 <b>Summa:</b> <code>{plan['price_uzs']:,}</code> UZS\n\n"
        f"💳 <b>Karta:</b> <code>{card_number}</code>\n"
        f"👤 <b>Karta egasi:</b> <code>{card_holder}</code>\n\n"
        f"📋 {instructions}\n\n"
        f"<i>Iltimos, to'lov kvitansiyasini (rasm yoki PDF) yuboring:</i>"
    )
    await callback.message.answer(text, parse_mode=ParseMode.HTML, reply_markup=get_cancel_keyboard())
    await callback.answer()


@router.message(PaymentState.uploading_receipt)
async def process_payment_receipt(message: Message, state: FSMContext, bot: Bot):
    data = await state.get_data()
    plan_id = data.get("selected_plan_id")
    plan = await db.get_premium_plan_by_id(plan_id)

    if not plan:
        await state.clear()
        await message.answer("Tarif topilmadi.")
        return

    file_id = None
    receipt_type = None

    if message.photo:
        file_id = message.photo[-1].file_id
        receipt_type = "photo"
    elif message.document and message.document.mime_type in ["application/pdf", "image/jpeg", "image/png"]:
        file_id = message.document.file_id
        receipt_type = "document"
    else:
        await message.answer("Iltimos, chek rasmi yoki PDF faylini yuboring!")
        return

    payment_id = await db.create_payment_order(
        user_id=message.from_user.id,
        plan_id=plan_id,
        amount_uzs=plan["price_uzs"],
        receipt_file_id=file_id,
        receipt_type=receipt_type,
    )

    await state.clear()
    is_admin = await db.is_any_admin(message.from_user.id)
    await message.answer(
        f"✅ <b>To'lov chekingiz qabul qilindi!</b>\n\n"
        f"🆔 To'lov ID: <code>#{payment_id}</code>\n"
        f"📦 Tarif: {html.escape(plan['name'])}\n"
        f"💰 Summa: {plan['price_uzs']:,} UZS\n\n"
        f"Administrator to'lovni tasdiqlashi bilan Premium avtomatik faollashadi.",
        parse_mode=ParseMode.HTML,
        reply_markup=get_main_menu_keyboard(is_admin=is_admin),
    )

    # Adminlarga bildirishnoma
    admin_markup = InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(text="✅ Tasdiqlash", callback_data=f"adm_pay_approve:{payment_id}"),
                InlineKeyboardButton(text="❌ Rad etish", callback_data=f"adm_pay_reject:{payment_id}"),
            ],
            [InlineKeyboardButton(text="👤 Foydalanuvchini ko'rish", callback_data=f"adm_view_user:{message.from_user.id}")],
        ]
    )
    admin_text = (
        f"💳 <b>YANGI TO'LOV CHEKI!</b>\n\n"
        f"🆔 ID: #{payment_id}\n"
        f"👤 Kimdan: {html.escape(message.from_user.full_name)} (ID: <code>{message.from_user.id}</code>)\n"
        f"📦 Tarif: {html.escape(plan['name'])}\n"
        f"💰 Summa: {plan['price_uzs']:,} UZS"
    )

    for a_id in ADMIN_IDS:
        try:
            if receipt_type == "photo":
                await bot.send_photo(chat_id=a_id, photo=file_id, caption=admin_text, parse_mode=ParseMode.HTML, reply_markup=admin_markup)
            else:
                await bot.send_document(chat_id=a_id, document=file_id, caption=admin_text, parse_mode=ParseMode.HTML, reply_markup=admin_markup)
        except Exception:
            pass


# ==============================================================================
# 9. ADMINGA YOZISH, BOT HAQIDA, O'CHIRISH
# ==============================================================================


@router.message(F.text == "💬 Adminga yozish")
async def handle_support_start(message: Message, state: FSMContext):
    await state.set_state(SupportState.waiting_for_message)
    await message.answer(
        "💬 Administratorga savol yoki taklifingizni yozing (matn yoki rasm):",
        reply_markup=get_cancel_keyboard(),
    )


@router.message(SupportState.waiting_for_message)
async def process_support_message(message: Message, state: FSMContext, bot: Bot):
    user_id = message.from_user.id
    admin_text = (
        f"📩 <b>FOYDALANUVCHIDAN MUROJAAT!</b>\n\n"
        f"👤 {html.escape(message.from_user.full_name)} (ID: <code>{user_id}</code>)\n"
        f"🔗 @{message.from_user.username or 'Yo‘q'}\n\n"
    )
    markup = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="✍️ Javob yozish", callback_data=f"adm_reply_support:{user_id}")]])

    for a_id in ADMIN_IDS:
        try:
            if message.text:
                await bot.send_message(chat_id=a_id, text=admin_text + html.escape(message.text), parse_mode=ParseMode.HTML, reply_markup=markup)
            elif message.photo:
                await bot.send_photo(chat_id=a_id, photo=message.photo[-1].file_id, caption=admin_text + html.escape(message.caption or ""), parse_mode=ParseMode.HTML, reply_markup=markup)
        except Exception:
            pass

    await state.clear()
    is_admin = await db.is_any_admin(user_id)
    await message.answer("✅ Xabaringiz adminga yuborildi!", reply_markup=get_main_menu_keyboard(is_admin=is_admin))


@router.message(F.text == "ℹ️ Bot haqida")
async def handle_about_bot(message: Message):
    text = await db.get_setting("about_text", DEFAULT_SYSTEM_TEXTS["about_text"]["default"])
    await message.answer(text, parse_mode=ParseMode.HTML)


@router.message(F.text == "🗑 Anketani o‘chirish")
async def handle_delete_profile_start(message: Message):
    markup = ReplyKeyboardMarkup(
        keyboard=[[KeyboardButton(text="🗑 Ha, o‘chirish")], [KeyboardButton(text="❌ Bekor qilish")]],
        resize_keyboard=True,
    )
    await message.answer(
        "⚠️ <b>Anketangizni haqiqatan ham butunlay o'chirmoqchimisiz?</b>\n"
        "Barcha rasmlaringiz, layklar va suhbatlar o'chiriladi.",
        parse_mode=ParseMode.HTML,
        reply_markup=markup,
    )


@router.message(F.text == "🗑 Ha, o‘chirish")
async def handle_delete_profile_confirm(message: Message, state: FSMContext):
    await db.delete_profile_completely(message.from_user.id)
    await state.clear()
    await message.answer("✅ Anketangiz butunlay o'chirildi.", reply_markup=ReplyKeyboardRemove())


# ==============================================================================
# 10. ADMIN BOSHQARUV PANELI (QULAY VA MUKAMMAL BOSHQARUV)
# ==============================================================================


@router.message(Command("admin"))
@router.message(F.text == "🛡 Admin paneli")
async def handle_admin_panel(message: Message, state: FSMContext):
    user_id = message.from_user.id
    if not await db.is_any_admin(user_id):
        await message.answer("❌ Sizda adminlik huquqi yo'q!")
        return

    # Admin rejimiga o'tish (FSM orqali admin rejimida qoladi!)
    await state.set_state(AdminState.in_admin_mode)
    is_super = await db.is_super_admin(user_id)
    role_str = "Bosh Admin (Super Admin)" if is_super else "Admin"
    await message.answer(
        f"🛡 <b>XLM DATING — ADMIN BOSHQARUV PANELI</b>\n"
        f"Mansab: <b>{role_str}</b>\n\n"
        f"Kerakli bo'limni tanlang:\n"
        f"<i>(Chiqish uchun '🚪 Admindan chiqish' tugmasini bosing)</i>",
        parse_mode=ParseMode.HTML,
        reply_markup=get_admin_main_menu_keyboard(),
    )


# ------------------ ADMIN ANKETASINI YASHIRISH / KO'RSATISH ------------------


@router.message(F.text == "🙈 Admin anketasini yashirish/ko'rsatish")
async def handle_admin_toggle_own_profile(message: Message):
    user_id = message.from_user.id
    if not await db.is_any_admin(user_id):
        return

    profile = await db.get_profile(user_id)
    if not profile:
        await message.answer("⚠️ Siz hali profil yaratmagansiz.")
        return

    new_val = 0 if profile["is_hidden"] == 1 else 1
    await db.update_profile_field(user_id, "is_hidden", new_val)
    st = "🙈 Sizning shaxsiy anketangiz qidiruvdan YASHIRILDI!" if new_val else "👁 Sizning anketangiz qidiruvda KO'RINADIGAN qilindi!"
    await message.answer(f"✅ {st}")


# ------------------ ADMIN: STATISTIKA ------------------


@router.message(F.text == "📊 Statistika")
async def handle_admin_statistics(message: Message):
    user_id = message.from_user.id
    if not await db.is_admin_allowed(user_id, "can_view_stats"):
        return await message.answer("❌ Sizda statistikani ko'rish huquqi yo'q.")

    stats = await db.get_statistics()
    text = (
        f"📊 <b>REAL VAQTDAGI STATISTIKA:</b>\n\n"
        f"👥 <b>Jami foydalanuvchilar:</b> {stats['total_users']}\n"
        f"📋 <b>Faol anketalar:</b> {stats['active_profiles']}\n"
        f"👨 <b>O'g'il bolalar:</b> {stats['male_count']}\n"
        f"👩 <b>Qizlar:</b> {stats['female_count']}\n"
        f"🆕 <b>Bugun ro'yxatdan o'tganlar:</b> {stats['today_registered']}\n\n"
        f"⭐ <b>Faol Premiumlar:</b> {stats['active_premium']}\n"
        f"⌛️ <b>Premium muddati tugaganlar:</b> {stats['expired_premium']}\n\n"
        f"💳 <b>Kutilayotgan to'lovlar:</b> {stats['pending_payments']}\n"
        f"✅ <b>Tasdiqlangan to'lovlar:</b> {stats['approved_payments']}\n"
        f"❌ <b>Rad etilgan to'lovlar:</b> {stats['rejected_payments']}\n\n"
        f"💬 <b>Faol suhbatlar:</b> {stats['active_chats']}\n"
        f"🚩 <b>Yangi shikoyatlar:</b> {stats['pending_reports']}\n"
        f"🚫 <b>Bloklanganlar:</b> {stats['banned_users']}\n"
        f"👑 <b>Sub-adminlar:</b> {stats['subadmins_count']}"
    )
    await message.answer(text, parse_mode=ParseMode.HTML)


# ------------------ ADMIN: TO'LOVLAR (ESKI VA YANGI CHEKLAR TARIXI BILAN) ------------------


@router.message(F.text.in_(["💳 To‘lovlar", "💳 To'lovlar"]))
async def handle_admin_payments(message: Message):
    user_id = message.from_user.id
    if not await db.is_admin_allowed(user_id, "can_view_payments"):
        return await message.answer("❌ Sizda to'lovlarni ko'rish huquqi yo'q.")
    await show_admin_payments_page(message, page=1, status="PENDING")


async def show_admin_payments_page(event: Any, page: int = 1, status: str = "PENDING"):
    payments, total_count, total_pages = await db.get_payments_paginated(page=page, per_page=8, status=status)

    status_titles = {
        "PENDING": "⏳ KUTILAYOTGAN CHEKLAR",
        "APPROVED": "✅ TASDIQLANGAN CHEKLAR",
        "REJECTED": "❌ RAD ETILGAN CHEKLAR",
        "ALL": "📋 BARCHA TO'LOVLAR TARIXI",
    }
    title = status_titles.get(status, "TO'LOVLAR")

    text = f"💳 <b>{title}</b>\n" f"Holat: <b>{status}</b> | Sahifa: <b>{page}/{total_pages}</b> (Jami: {total_count} ta)\n\n"

    if not payments:
        text += "<i>Ushbu toifada to'lovlar topilmadi.</i>"

    buttons = []
    for p in payments:
        icon = "⏳" if p["status"] == "PENDING" else ("✅" if p["status"] == "APPROVED" else "❌")
        btn_text = f"{icon} #{p['id']} - {p['first_name'] or p['user_id']} ({p['amount_uzs']:,} UZS)"
        buttons.append([InlineKeyboardButton(text=btn_text, callback_data=f"adm_pay_detail:{p['id']}")])

    # Paginatsiya
    nav = []
    if page > 1:
        nav.append(InlineKeyboardButton(text="◀️ Oldingi", callback_data=f"pay_page:{page-1}:{status}"))
    nav.append(InlineKeyboardButton(text=f"📄 {page}/{total_pages}", callback_data="ignore"))
    if page < total_pages:
        nav.append(InlineKeyboardButton(text="▶️ Keyingi", callback_data=f"pay_page:{page+1}:{status}"))
    if nav:
        buttons.append(nav)

    # Toifalar (Tablar)
    buttons.append(
        [
            InlineKeyboardButton(text="⏳ Kutilayotgan", callback_data="pay_page:1:PENDING"),
            InlineKeyboardButton(text="✅ Tasdiqlangan", callback_data="pay_page:1:APPROVED"),
            InlineKeyboardButton(text="❌ Rad etilgan", callback_data="pay_page:1:REJECTED"),
        ]
    )
    buttons.append([InlineKeyboardButton(text="📋 Barcha to'lovlar tarixi", callback_data="pay_page:1:ALL")])

    markup = InlineKeyboardMarkup(inline_keyboard=buttons)
    if isinstance(event, CallbackQuery):
        try:
            await event.message.edit_text(text, parse_mode=ParseMode.HTML, reply_markup=markup)
        except Exception:
            await event.message.answer(text, parse_mode=ParseMode.HTML, reply_markup=markup)
        await event.answer()
    else:
        await event.answer(text, parse_mode=ParseMode.HTML, reply_markup=markup)


@router.callback_query(F.data.startswith("pay_page:"))
async def handle_payments_pagination_cb(callback: CallbackQuery):
    if not await db.is_admin_allowed(callback.from_user.id, "can_view_payments"):
        return await callback.answer("Ruxsat yo'q!", show_alert=True)
    _, p_str, st = callback.data.split(":")
    await show_admin_payments_page(callback, page=int(p_str), status=st)


@router.callback_query(F.data.startswith("adm_pay_detail:"))
async def handle_admin_payment_detail(callback: CallbackQuery, bot: Bot):
    payment_id = int(callback.data.split(":")[1])
    payment = await db.get_payment(payment_id)
    if not payment:
        return await callback.answer("To'lov topilmadi.", show_alert=True)

    status_str = "⏳ Kutilmoqda (PENDING)" if payment["status"] == "PENDING" else ("✅ Tasdiqlangan (APPROVED)" if payment["status"] == "APPROVED" else f"❌ Rad etilgan (Sabab: {payment['rejection_reason']})")

    caption = (
        f"📄 <b>TO'LOV CHEKI #{payment['id']}:</b>\n\n"
        f"👤 <b>Foydalanuvchi:</b> {html.escape(payment['first_name'] or '')} (ID: <code>{payment['user_id']}</code>)\n"
        f"🔗 <b>Username:</b> @{payment['username'] or 'Yo‘q'}\n"
        f"📦 <b>Tarif:</b> {payment['plan_name']} ({payment['duration_days']} kun)\n"
        f"💰 <b>Summa:</b> {payment['amount_uzs']:,} UZS\n"
        f"🕒 <b>Yuborilgan vaqt:</b> {payment['created_at']}\n"
        f"⚡️ <b>Holat:</b> {status_str}\n"
    )

    if payment["reviewed_by"]:
        caption += f"👮‍♂️ <b>Ko'rib chiqqan admin:</b> <code>{payment['reviewed_by']}</code> ({payment['reviewed_at']})\n"

    buttons = []
    if payment["status"] == "PENDING":
        buttons.append(
            [
                InlineKeyboardButton(text="✅ Tasdiqlash", callback_data=f"adm_pay_approve:{payment['id']}"),
                InlineKeyboardButton(text="❌ Rad etish", callback_data=f"adm_pay_reject:{payment['id']}"),
            ]
        )
    buttons.append([InlineKeyboardButton(text="👤 Foydalanuvchi profilini ko'rish", callback_data=f"adm_view_user:{payment['user_id']}")])

    markup = InlineKeyboardMarkup(inline_keyboard=buttons)
    try:
        if payment["receipt_type"] == "photo":
            await callback.message.answer_photo(photo=payment["receipt_file_id"], caption=caption, parse_mode=ParseMode.HTML, reply_markup=markup)
        else:
            await callback.message.answer_document(document=payment["receipt_file_id"], caption=caption, parse_mode=ParseMode.HTML, reply_markup=markup)
    except Exception:
        await callback.message.answer(caption, parse_mode=ParseMode.HTML, reply_markup=markup)
    await callback.answer()


@router.callback_query(F.data.startswith("adm_pay_approve:"))
async def handle_admin_approve_payment(callback: CallbackQuery, bot: Bot):
    admin_id = callback.from_user.id
    if not await db.is_admin_allowed(admin_id, "can_approve_payments"):
        return await callback.answer("❌ Sizda to'lovni tasdiqlash huquqi yo'q!", show_alert=True)

    payment_id = int(callback.data.split(":")[1])
    success, msg, user_id, new_end_str = await db.approve_payment(payment_id, admin_id)
    if not success:
        return await callback.answer(msg, show_alert=True)

    await callback.answer("✅ To'lov tasdiqlandi!")
    try:
        await callback.message.edit_caption(caption=callback.message.caption + f"\n\n<b>✅ TASDIQLANDI!</b> (Admin: {admin_id})", parse_mode=ParseMode.HTML)
    except Exception:
        pass

    if user_id:
        try:
            await bot.send_message(
                chat_id=user_id,
                text=f"🎉 <b>To'lovingiz admin tomonidan tasdiqlandi!</b>\n\n⭐ <b>Premium faollashtirildi!</b>\nAmal qilish muddati: <b>{new_end_str}</b> gacha.",
                parse_mode=ParseMode.HTML,
            )
        except Exception:
            pass


@router.callback_query(F.data.startswith("adm_pay_reject:"))
async def handle_admin_reject_payment_start(callback: CallbackQuery, state: FSMContext):
    admin_id = callback.from_user.id
    if not await db.is_admin_allowed(admin_id, "can_reject_payments"):
        return await callback.answer("❌ Sizda to'lovni rad etish huquqi yo'q!", show_alert=True)

    payment_id = int(callback.data.split(":")[1])
    await state.set_state(AdminState.reject_payment_reason)
    await state.update_data(reject_payment_id=payment_id)

    markup = ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="Pul hisobga tushmadi")],
            [KeyboardButton(text="Chek soxta / Noto'g'ri chek")],
            [KeyboardButton(text="To'lov summasi to'liq emas")],
            [KeyboardButton(text="❌ Bekor qilish")],
        ],
        resize_keyboard=True,
    )
    await callback.message.answer(f"🆔 To'lov #{payment_id} ni rad etish sababini tanlang yoki yozing:", reply_markup=markup)
    await callback.answer()


@router.message(AdminState.reject_payment_reason)
async def handle_admin_reject_payment_finish(message: Message, state: FSMContext, bot: Bot):
    admin_id = message.from_user.id
    reason = (message.text or "").strip()
    data = await state.get_data()
    payment_id = data.get("reject_payment_id")

    success, msg, user_id = await db.reject_payment(payment_id, admin_id, reason)
    await state.set_state(AdminState.in_admin_mode)
    await message.answer(f"❌ To'lov #{payment_id} rad etildi.", reply_markup=get_admin_main_menu_keyboard())

    if user_id:
        try:
            await bot.send_message(
                chat_id=user_id,
                text=f"❌ <b>To'lov chekingiz rad etildi.</b>\nSabab: <i>{html.escape(reason)}</i>",
                parse_mode=ParseMode.HTML,
            )
        except Exception:
            pass


# ------------------ ADMIN: CMS BOT MATNLARINI TAHRIRLASH ------------------


@router.message(F.text == "📝 Bot matnlarini tahrirlash")
async def handle_cms_menu(message: Message):
    user_id = message.from_user.id
    if not await db.is_admin_allowed(user_id, "can_manage_settings"):
        return await message.answer("❌ Sizda sozlamalarni o'zgartirish huquqi yo'q.")

    buttons = []
    for key, data in DEFAULT_SYSTEM_TEXTS.items():
        buttons.append([InlineKeyboardButton(text=data["title"], callback_data=f"cms_view:{key}")])

    markup = InlineKeyboardMarkup(inline_keyboard=buttons)
    await message.answer(
        "📝 <b>BOT MATNLARINI TAHRIRLASH (CMS):</b>\n\n"
        "Quyidagi tizim matnlaridan birini tanlang va uni xohlagancha o'zgartiring:",
        parse_mode=ParseMode.HTML,
        reply_markup=markup,
    )


@router.callback_query(F.data.startswith("cms_view:"))
async def handle_cms_view_text(callback: CallbackQuery, state: FSMContext):
    key = callback.data.split(":")[1]
    cfg = DEFAULT_SYSTEM_TEXTS.get(key)
    if not cfg:
        return await callback.answer("Matn topilmadi.", show_alert=True)

    curr_val = await db.get_setting(key, cfg["default"])
    text = (
        f"📝 <b>MATNNI TAHRIRLASH:</b> {cfg['title']}\n\n"
        f"📌 <b>Hozirgi (eskisi) matn:</b>\n"
        f"----------------------------------------\n"
        f"{curr_val}\n"
        f"----------------------------------------\n\n"
        f"Ushbu matnni yangilash yoki asliga qaytarishni tanlang:"
    )

    markup = InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="✏️ Yangi matn kiritish", callback_data=f"cms_edit:{key}")],
            [InlineKeyboardButton(text="🔄 Asliga qaytarish (Standart)", callback_data=f"cms_reset:{key}")],
            [InlineKeyboardButton(text="⬅️ Ortga", callback_data="cms_back_list")],
        ]
    )
    await callback.message.edit_text(text, parse_mode=ParseMode.HTML, reply_markup=markup)
    await callback.answer()


@router.callback_query(F.data == "cms_back_list")
async def handle_cms_back_list(callback: CallbackQuery):
    buttons = []
    for key, data in DEFAULT_SYSTEM_TEXTS.items():
        buttons.append([InlineKeyboardButton(text=data["title"], callback_data=f"cms_view:{key}")])
    await callback.message.edit_text("Quyidagi tizim matnlaridan birini tanlang:", reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
    await callback.answer()


@router.callback_query(F.data.startswith("cms_edit:"))
async def handle_cms_edit_start(callback: CallbackQuery, state: FSMContext):
    key = callback.data.split(":")[1]
    cfg = DEFAULT_SYSTEM_TEXTS.get(key)
    await state.set_state(AdminState.cms_enter_text)
    await state.update_data(cms_edit_key=key)

    await callback.message.answer(
        f"✏️ <b>'{cfg['title']}' uchun yangi matnni yuboring:</b>\n"
        f"<i>(HTML teglari qo'llab-quvvatlanadi: &lt;b&gt;, &lt;i&gt;, &lt;code&gt;)</i>",
        parse_mode=ParseMode.HTML,
        reply_markup=get_cancel_keyboard(),
    )
    await callback.answer()


@router.message(AdminState.cms_enter_text)
async def handle_cms_save_new_text(message: Message, state: FSMContext):
    new_text = message.text or message.caption or ""
    if not new_text:
        return await message.answer("Iltimos, matn yuboring:")

    data = await state.get_data()
    key = data.get("cms_edit_key")
    await db.set_setting(key, new_text)

    await state.set_state(AdminState.in_admin_mode)
    cfg = DEFAULT_SYSTEM_TEXTS.get(key, {"title": key})
    await message.answer(
        f"✅ <b>'{cfg['title']}' muvaffaqiyatli yangilandi!</b>\nEndi bot ushbu yangi matndan foydalanadi.",
        parse_mode=ParseMode.HTML,
        reply_markup=get_admin_main_menu_keyboard(),
    )


@router.callback_query(F.data.startswith("cms_reset:"))
async def handle_cms_reset_default(callback: CallbackQuery):
    key = callback.data.split(":")[1]
    cfg = DEFAULT_SYSTEM_TEXTS.get(key)
    if cfg:
        await db.set_setting(key, cfg["default"])
        await callback.answer("✅ Matn standart holatiga qaytarildi!", show_alert=True)
        # Qayta ko'rsatish
        curr_val = cfg["default"]
        text = (
            f"📝 <b>MATNNI TAHRIRLASH:</b> {cfg['title']}\n\n"
            f"📌 <b>Hozirgi matn (Standart):</b>\n"
            f"----------------------------------------\n"
            f"{curr_val}\n"
            f"----------------------------------------"
        )
        markup = InlineKeyboardMarkup(
            inline_keyboard=[
                [InlineKeyboardButton(text="✏️ Yangi matn kiritish", callback_data=f"cms_edit:{key}")],
                [InlineKeyboardButton(text="⬅️ Ortga", callback_data="cms_back_list")],
            ]
        )
        try:
            await callback.message.edit_text(text, parse_mode=ParseMode.HTML, reply_markup=markup)
        except Exception:
            pass


# ------------------ ADMIN: FOYDALANUVCHILAR (PAGINATSIYA) ------------------


@router.message(F.text == "👥 Foydalanuvchilar")
async def handle_admin_users_list_menu(message: Message):
    if not await db.is_admin_allowed(message.from_user.id, "can_view_users"):
        return await message.answer("❌ Ruxsat yo'q.")
    await show_admin_users_page(message, page=1, gender_filter="all")


async def show_admin_users_page(event: Any, page: int = 1, gender_filter: str = "all"):
    users, total_count, total_pages = await db.get_users_paginated(page=page, per_page=10, gender_filter=gender_filter)

    text = f"👥 <b>FOYDALANUVCHILAR RO'YXATI</b>\n" f"Filtr: <b>{gender_filter.upper()}</b> | Sahifa: <b>{page}/{total_pages}</b> (Jami: {total_count} ta)\n\n"

    if not users:
        text += "<i>Foydalanuvchilar topilmadi.</i>"

    buttons = []
    for idx, u in enumerate(users, 1 + (page - 1) * 10):
        name = u["profile_name"] or u["first_name"] or "Foydalanuvchi"
        prem = "⭐" if u.get("prem_active") else ""
        ban = "🚫" if u["is_banned"] else "🟢"
        gender_icon = "👨" if u.get("gender") == "male" else ("👩" if u.get("gender") == "female" else "👤")

        text += f"{idx}. {ban} {gender_icon} <b>{html.escape(name)}</b> {prem} (ID: <code>{u['user_id']}</code>)\n"
        buttons.append([InlineKeyboardButton(text=f"👤 {name} ({u['user_id']})", callback_data=f"adm_view_user:{u['user_id']}")])

    nav_buttons = []
    if page > 1:
        nav_buttons.append(InlineKeyboardButton(text="◀️ Oldingi", callback_data=f"users_page:{page-1}:{gender_filter}"))
    nav_buttons.append(InlineKeyboardButton(text=f"📄 {page}/{total_pages}", callback_data="ignore"))
    if page < total_pages:
        nav_buttons.append(InlineKeyboardButton(text="▶️ Keyingi", callback_data=f"users_page:{page+1}:{gender_filter}"))
    if nav_buttons:
        buttons.append(nav_buttons)

    buttons.append(
        [
            InlineKeyboardButton(text="🌍 Barchasi", callback_data=f"users_page:1:all"),
            InlineKeyboardButton(text="👨 Yigitlar", callback_data=f"users_page:1:male"),
            InlineKeyboardButton(text="👩 Qizlar", callback_data=f"users_page:1:female"),
        ]
    )
    buttons.append(
        [
            InlineKeyboardButton(text="⭐ Premiumlar", callback_data=f"users_page:1:premium"),
            InlineKeyboardButton(text="🚫 Bloklanganlar", callback_data=f"users_page:1:banned"),
            InlineKeyboardButton(text="🔎 ID Qidirish", callback_data="adm_search_user_btn"),
        ]
    )

    markup = InlineKeyboardMarkup(inline_keyboard=buttons)
    if isinstance(event, CallbackQuery):
        try:
            await event.message.edit_text(text, parse_mode=ParseMode.HTML, reply_markup=markup)
        except Exception:
            await event.message.answer(text, parse_mode=ParseMode.HTML, reply_markup=markup)
        await event.answer()
    else:
        await event.answer(text, parse_mode=ParseMode.HTML, reply_markup=markup)


@router.callback_query(F.data.startswith("users_page:"))
async def handle_users_pagination_callback(callback: CallbackQuery):
    if not await db.is_admin_allowed(callback.from_user.id, "can_view_users"):
        return await callback.answer("❌ Ruxsat yo'q!", show_alert=True)
    _, p_str, g_filter = callback.data.split(":")
    await show_admin_users_page(callback, page=int(p_str), gender_filter=g_filter)


@router.callback_query(F.data == "adm_search_user_btn")
async def handle_admin_search_user_btn(callback: CallbackQuery, state: FSMContext):
    await state.set_state(AdminState.search_user_id)
    await callback.message.answer("Qidirilayotgan foydalanuvchining Telegram ID sini kiriting:", reply_markup=get_cancel_keyboard())
    await callback.answer()


@router.message(AdminState.search_user_id)
async def handle_admin_search_user_id(message: Message, state: FSMContext, bot: Bot):
    text = (message.text or "").strip()
    if not text.isdigit():
        return await message.answer("Faqat raqamli Telegram ID kiriting:")
    target_id = int(text)
    await state.set_state(AdminState.in_admin_mode)
    await show_admin_user_card(message, target_id, bot)


@router.callback_query(F.data.startswith("adm_view_user:"))
async def handle_admin_view_user_callback(callback: CallbackQuery, bot: Bot):
    if not await db.is_admin_allowed(callback.from_user.id, "can_view_users"):
        return await callback.answer("❌ Ruxsat yo'q!", show_alert=True)
    target_id = int(callback.data.split(":")[1])
    await show_admin_user_card(callback.message, target_id, bot)
    await callback.answer()


async def show_admin_user_card(message: Message, target_id: int, bot: Bot):
    viewer_id = message.chat.id
    async with await db.get_connection() as conn:
        async with conn.execute("SELECT * FROM users WHERE user_id = ?", (target_id,)) as cur:
            user = await cur.fetchone()
        async with conn.execute("SELECT * FROM profiles WHERE user_id = ?", (target_id,)) as cur2:
            profile = await cur2.fetchone()

    if not user:
        await message.answer("❌ Foydalanuvchi topilmadi.")
        return

    can_view_uname = await db.is_admin_allowed(viewer_id, "can_view_usernames")
    username_str = f"@{user['username']}" if (user["username"] and can_view_uname) else ("Yashiringan" if not can_view_uname else "Yo'q")
    phone_str = user["phone_number"] if (user["phone_number"] and can_view_uname) else ("Yashiringan" if not can_view_uname else "Kiritilmagan")

    is_premium, end_date = await db.is_premium_active(target_id)
    banned_status = f"🚫 BLOKLANGAN ({user['ban_reason']})" if user["is_banned"] else "🟢 Faol"

    text = (
        f"👤 <b>FOYDALANUVCHI PROFILI:</b>\n\n"
        f"🆔 <b>Telegram ID:</b> <code>{user['user_id']}</code>\n"
        f"🔗 <b>Username:</b> {username_str}\n"
        f"📱 <b>Telefon:</b> {phone_str}\n"
        f"📅 <b>Ro'yxatdan o'tgan:</b> {user['created_at']}\n"
        f"⚡️ <b>Holat:</b> {banned_status}\n"
        f"⭐ <b>Premium:</b> {'Ha (' + str(end_date) + ')' if is_premium else 'Yo‘q'}\n\n"
    )

    if profile:
        g = "👨 O‘g‘il bola" if profile["gender"] == "male" else "👩 Qiz bola"
        text += (
            f"📋 <b>ANKETA:</b>\n"
            f"• Ism: {html.escape(profile['name'])}\n"
            f"• Yosh: {profile['age']}\n"
            f"• Jins: {g}\n"
            f"• Hudud: {profile['region']}\n"
            f"• Bio: {html.escape(profile['bio'] or 'Ko‘rsatilmagan')}\n"
            f"• Yashirilgan: {'Ha' if profile['is_hidden'] else 'Yo‘q'}\n"
        )
    else:
        text += "📋 <i>Anketa yaratilmagan.</i>\n"

    buttons = []
    row1 = []
    if await db.is_admin_allowed(viewer_id, "can_send_messages"):
        row1.append(InlineKeyboardButton(text="💬 Xabar yozish", callback_data=f"adm_send_msg:{target_id}"))
    if await db.is_admin_allowed(viewer_id, "can_manage_premium"):
        row1.append(InlineKeyboardButton(text="⭐ Premium berish", callback_data=f"adm_give_prem:{target_id}"))
    if row1:
        buttons.append(row1)

    row2 = []
    if user["is_banned"]:
        if await db.is_admin_allowed(viewer_id, "can_unban_users"):
            row2.append(InlineKeyboardButton(text="🔓 Blokdan chiqarish", callback_data=f"adm_unban:{target_id}"))
    else:
        if await db.is_admin_allowed(viewer_id, "can_ban_users"):
            row2.append(InlineKeyboardButton(text="🚫 Bloklash", callback_data=f"adm_ban:{target_id}"))
    if is_premium and await db.is_admin_allowed(viewer_id, "can_manage_premium"):
        row2.append(InlineKeyboardButton(text="❌ Premiumni bekor qilish", callback_data=f"adm_rev_prem:{target_id}"))
    if row2:
        buttons.append(row2)

    if await db.is_admin_allowed(viewer_id, "can_manage_admins"):
        subadmin = await db.get_subadmin(target_id)
        if not subadmin:
            buttons.append([InlineKeyboardButton(text="👑 Admin tayinlash", callback_data=f"adm_make_admin:{target_id}")])
        else:
            buttons.append([InlineKeyboardButton(text="⚙️ Admin huquqlarini sozlash", callback_data=f"adm_edit_admin_perms:{target_id}")])

    markup = InlineKeyboardMarkup(inline_keyboard=buttons)
    photos = await db.get_profile_photos(target_id)
    if photos:
        await message.answer_photo(photo=photos[0]["file_id"], caption=text, parse_mode=ParseMode.HTML, reply_markup=markup)
    else:
        await message.answer(text, parse_mode=ParseMode.HTML, reply_markup=markup)


# ------------------ ADMINLARNI BOSHQARISH VA 22 TA HUQUQLAR (RBAC) ------------------


@router.message(F.text == "👑 Adminlar boshqaruvi")
async def handle_admin_management_menu(message: Message):
    user_id = message.from_user.id
    if not await db.is_admin_allowed(user_id, "can_manage_admins"):
        return await message.answer("❌ Ushbu bo'lim faqat Bosh Adminlar uchun ochiq!")

    subadmins = await db.get_all_subadmins()
    text = "👑 <b>ADMINLAR VA HUQUQLAR BOSHQARUVI:</b>\n\n"
    if not subadmins:
        text += "<i>Hozircha tayinlangan sub-adminlar yo'q.</i>\n"
    else:
        for idx, a in enumerate(subadmins, 1):
            text += f"{idx}. <b>{html.escape(a['admin_name'])}</b> ({a['role_title']}) - ID: <code>{a['user_id']}</code>\n"

    markup = InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="➕ Yangi admin tayinlash", callback_data="adm_start_add_admin")],
            [InlineKeyboardButton(text="📋 Adminlar ro'yxati va huquqlari", callback_data="adm_list_admins")],
        ]
    )
    await message.answer(text, parse_mode=ParseMode.HTML, reply_markup=markup)


@router.callback_query(F.data == "adm_start_add_admin")
async def handle_start_add_admin_cb(callback: CallbackQuery, state: FSMContext):
    if not await db.is_admin_allowed(callback.from_user.id, "can_manage_admins"):
        return await callback.answer("Ruxsat yo'q!", show_alert=True)
    await state.set_state(AdminState.add_admin_id)
    await callback.message.answer("Yangi admin qilmoqchi bo'lgan foydalanuvchining Telegram ID sini kiriting:", reply_markup=get_cancel_keyboard())
    await callback.answer()


@router.message(AdminState.add_admin_id)
async def handle_add_admin_id_msg(message: Message, state: FSMContext):
    text = (message.text or "").strip()
    if not text.isdigit():
        return await message.answer("Iltimos, faqat raqamli Telegram ID kiriting:")

    target_id = int(text)
    async with await db.get_connection() as conn:
        async with conn.execute("SELECT * FROM users WHERE user_id = ?", (target_id,)) as cur:
            user = await cur.fetchone()

    if not user:
        return await message.answer("❌ Bu foydalanuvchi botda ro'yxatdan o'tmagan.")

    await state.update_data(new_admin_user_id=target_id)
    await state.set_state(AdminState.add_admin_name)
    await message.answer("Ushbu admin uchun nom / lavozim kiriting (Masalan: Moderator Sardor):", reply_markup=get_cancel_keyboard())


@router.message(AdminState.add_admin_name)
async def handle_add_admin_name_msg(message: Message, state: FSMContext):
    name = (message.text or "").strip()
    data = await state.get_data()
    target_id = data.get("new_admin_user_id")

    await db.add_subadmin(user_id=target_id, added_by=message.from_user.id, admin_name=name, role_title="Admin")
    await state.set_state(AdminState.in_admin_mode)
    await message.answer(
        f"✅ <b>{html.escape(name)}</b> muvaffaqiyatli admin etib tayinlandi!\nEndi uning ruxsatlarini quyidagi menyudan sozlashingiz mumkin:",
        parse_mode=ParseMode.HTML,
        reply_markup=get_admin_main_menu_keyboard(),
    )
    await show_admin_permissions_editor(message, target_id)


@router.callback_query(F.data.startswith("adm_make_admin:"))
async def handle_make_admin_direct(callback: CallbackQuery, state: FSMContext):
    target_id = int(callback.data.split(":")[1])
    await state.update_data(new_admin_user_id=target_id)
    await state.set_state(AdminState.add_admin_name)
    await callback.message.answer("Ushbu admin uchun nom / lavozim kiriting (Masalan: Moderator Ali):", reply_markup=get_cancel_keyboard())
    await callback.answer()


@router.callback_query(F.data == "adm_list_admins")
async def handle_list_admins_cb(callback: CallbackQuery):
    if not await db.is_admin_allowed(callback.from_user.id, "can_manage_admins"):
        return await callback.answer("Ruxsat yo'q!", show_alert=True)
    subadmins = await db.get_all_subadmins()
    if not subadmins:
        return await callback.answer("Adminlar ro'yxati bo'sh.", show_alert=True)

    buttons = []
    for a in subadmins:
        buttons.append([InlineKeyboardButton(text=f"⚙️ {a['admin_name']} ({a['user_id']})", callback_data=f"adm_edit_admin_perms:{a['user_id']}")])

    await callback.message.answer("Boshqarish uchun adminni tanlang:", reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
    await callback.answer()


@router.callback_query(F.data.startswith("adm_edit_admin_perms:"))
async def handle_edit_admin_perms_cb(callback: CallbackQuery):
    target_id = int(callback.data.split(":")[1])
    await show_admin_permissions_editor(callback.message, target_id)
    await callback.answer()


async def show_admin_permissions_editor(message: Message, target_id: int):
    admin_row = await db.get_subadmin(target_id)
    if not admin_row:
        return await message.answer("Admin topilmadi.")

    try:
        perms = json.loads(admin_row["permissions"])
    except Exception:
        perms = dict(DEFAULT_SUBADMIN_PERMS)

    text = (
        f"👑 <b>ADMIN HUQUQLARINI SOZLASH:</b>\n"
        f"Admin: <b>{html.escape(admin_row['admin_name'])}</b>\n"
        f"Telegram ID: <code>{admin_row['user_id']}</code>\n\n"
        f"<i>Ruxsatni o'zgartirish uchun tugmani bosing:</i>"
    )

    buttons = []
    for key, label in ADMIN_PERMISSIONS.items():
        is_allowed = perms.get(key, False)
        status_icon = "✅" if is_allowed else "❌"
        buttons.append([InlineKeyboardButton(text=f"{status_icon} {label}", callback_data=f"adm_toggle_p:{target_id}:{key}")])

    buttons.append([InlineKeyboardButton(text="🗑 Adminlikni bekor qilish", callback_data=f"adm_delete_subadmin:{target_id}")])
    buttons.append([InlineKeyboardButton(text="⬅️ Adminlar ro'yxatiga qaytish", callback_data="adm_list_admins")])

    markup = InlineKeyboardMarkup(inline_keyboard=buttons)
    try:
        await message.edit_text(text, parse_mode=ParseMode.HTML, reply_markup=markup)
    except Exception:
        await message.answer(text, parse_mode=ParseMode.HTML, reply_markup=markup)


@router.callback_query(F.data.startswith("adm_toggle_p:"))
async def handle_toggle_permission_cb(callback: CallbackQuery):
    if not await db.is_admin_allowed(callback.from_user.id, "can_manage_admins"):
        return await callback.answer("Ruxsat yo'q!", show_alert=True)
    _, target_str, perm_key = callback.data.split(":")
    target_id = int(target_str)

    new_val = await db.toggle_admin_permission(target_id, perm_key)
    perm_label = ADMIN_PERMISSIONS.get(perm_key, perm_key)
    st = "berildi" if new_val else "cheqlandi"
    await callback.answer(f"{perm_label}: {st}!")
    await show_admin_permissions_editor(callback.message, target_id)


@router.callback_query(F.data.startswith("adm_delete_subadmin:"))
async def handle_delete_subadmin_cb(callback: CallbackQuery):
    if not await db.is_admin_allowed(callback.from_user.id, "can_manage_admins"):
        return await callback.answer("Ruxsat yo'q!", show_alert=True)
    target_id = int(callback.data.split(":")[1])
    await db.remove_subadmin(target_id, admin_id=callback.from_user.id)
    await callback.answer("Admin muvaffaqiyatli olib tashlandi!", show_alert=True)
    try:
        await callback.message.delete()
    except Exception:
        pass


# ------------------ ADMIN: BLOKLASH VA XABAR YOZISH ------------------


@router.callback_query(F.data.startswith("adm_ban:"))
async def handle_admin_ban_user_start(callback: CallbackQuery, state: FSMContext):
    if not await db.is_admin_allowed(callback.from_user.id, "can_ban_users"):
        return await callback.answer("Ruxsat yo'q!", show_alert=True)
    target_id = int(callback.data.split(":")[1])
    await state.set_state(AdminState.ban_user_reason)
    await state.update_data(ban_target_id=target_id)
    await callback.message.answer(f"Foydalanuvchi {target_id} ni bloklash sababini kiriting:", reply_markup=get_cancel_keyboard())
    await callback.answer()


@router.message(AdminState.ban_user_reason)
async def handle_admin_ban_user_finish(message: Message, state: FSMContext, bot: Bot):
    reason = (message.text or "").strip()
    data = await state.get_data()
    target_id = data.get("ban_target_id")
    await db.ban_user(target_id, reason, admin_id=message.from_user.id)
    await state.set_state(AdminState.in_admin_mode)
    await message.answer(f"✅ Foydalanuvchi {target_id} bloklandi!", reply_markup=get_admin_main_menu_keyboard())
    try:
        await bot.send_message(chat_id=target_id, text=f"🚫 Profilingiz bloklandi.\nSabab: {html.escape(reason)}")
    except Exception:
        pass


@router.callback_query(F.data.startswith("adm_unban:"))
async def handle_admin_unban_user(callback: CallbackQuery, bot: Bot):
    if not await db.is_admin_allowed(callback.from_user.id, "can_unban_users"):
        return await callback.answer("Ruxsat yo'q!", show_alert=True)
    target_id = int(callback.data.split(":")[1])
    await db.unban_user(target_id, admin_id=callback.from_user.id)
    await callback.answer("Foydalanuvchi blokdan chiqarildi!", show_alert=True)
    try:
        await bot.send_message(chat_id=target_id, text="🟢 Profilingiz blokdan chiqarildi.")
    except Exception:
        pass


@router.callback_query(F.data.startswith("adm_send_msg:"))
async def handle_admin_send_msg_start(callback: CallbackQuery, state: FSMContext):
    if not await db.is_admin_allowed(callback.from_user.id, "can_send_messages"):
        return await callback.answer("Ruxsat yo'q!", show_alert=True)
    target_id = int(callback.data.split(":")[1])
    await state.set_state(AdminState.send_message_text)
    await state.update_data(msg_target_id=target_id)
    await callback.message.answer(f"Foydalanuvchi {target_id} ga xabaringizni yozing:", reply_markup=get_cancel_keyboard())
    await callback.answer()


@router.message(AdminState.send_message_text)
async def handle_admin_send_msg_finish(message: Message, state: FSMContext, bot: Bot):
    text = (message.text or "").strip()
    data = await state.get_data()
    target_id = data.get("msg_target_id")

    try:
        await bot.send_message(chat_id=target_id, text=f"📩 <b>ADMINISTRATOR XABARI:</b>\n\n{html.escape(text)}", parse_mode=ParseMode.HTML)
        await message.answer(f"✅ Xabar {target_id} ga yetkazildi!", reply_markup=get_admin_main_menu_keyboard())
    except Exception as e:
        await message.answer(f"❌ Xatolik: {e}", reply_markup=get_admin_main_menu_keyboard())

    await state.set_state(AdminState.in_admin_mode)


@router.callback_query(F.data.startswith("adm_give_prem:"))
async def handle_admin_give_premium_start(callback: CallbackQuery, state: FSMContext):
    if not await db.is_admin_allowed(callback.from_user.id, "can_manage_premium"):
        return await callback.answer("Ruxsat yo'q!", show_alert=True)
    target_id = int(callback.data.split(":")[1])
    await state.set_state(AdminState.grant_premium_days)
    await state.update_data(prem_target_id=target_id)
    await callback.message.answer("Necha kunlik Premium bermoqchisiz? (Masalan: 30):", reply_markup=get_cancel_keyboard())
    await callback.answer()


@router.message(AdminState.grant_premium_days)
async def handle_admin_give_premium_finish(message: Message, state: FSMContext, bot: Bot):
    text = (message.text or "").strip()
    if not text.isdigit() or int(text) < 1:
        return await message.answer("Musbat son kiriting:")

    days = int(text)
    data = await state.get_data()
    target_id = data.get("prem_target_id")
    success, end_str = await db.grant_manual_premium(target_id, days, admin_id=message.from_user.id)
    await state.set_state(AdminState.in_admin_mode)
    await message.answer(f"✅ {target_id} ga {days} kunlik Premium berildi! ({end_str} gacha)", reply_markup=get_admin_main_menu_keyboard())
    try:
        await bot.send_message(chat_id=target_id, text=f"🎁 Sizga administrator tomonidan {days} kunlik ⭐ PREMIUM taqdim etildi!")
    except Exception:
        pass


@router.callback_query(F.data.startswith("adm_rev_prem:"))
async def handle_admin_revoke_premium(callback: CallbackQuery, bot: Bot):
    if not await db.is_admin_allowed(callback.from_user.id, "can_manage_premium"):
        return await callback.answer("Ruxsat yo'q!", show_alert=True)
    target_id = int(callback.data.split(":")[1])
    await db.revoke_premium(target_id, admin_id=callback.from_user.id)
    await callback.answer("Premium bekor qilindi!", show_alert=True)
    try:
        await bot.send_message(chat_id=target_id, text="ℹ️ Premium obunangiz bekor qilindi.")
    except Exception:
        pass


# ------------------ ADMIN: PREMIUM TARIFLARINI BOSHQARISH ------------------


@router.message(F.text == "⭐ Premium boshqaruvi")
async def handle_admin_plans_menu(message: Message):
    if not await db.is_admin_allowed(message.from_user.id, "can_manage_plans"):
        return await message.answer("❌ Ruxsat yo'q.")
    plans = await db.get_premium_plans(only_active=False)
    text = "⭐ <b>PREMIUM TARIFLAR:</b>\n\n"
    buttons = []
    for p in plans:
        st = "🟢" if p["is_active"] else "🔴"
        text += f"• {st} <b>{p['name']}</b> ({p['duration_days']} kun): {p['price_uzs']:,} UZS\n"
        buttons.append([InlineKeyboardButton(text=f"✏️ {p['name']} narxini o'zgartirish", callback_data=f"adm_edit_plan:{p['id']}")])

    await message.answer(text, parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))


@router.callback_query(F.data.startswith("adm_edit_plan:"))
async def handle_edit_plan_cb(callback: CallbackQuery, state: FSMContext):
    plan_id = int(callback.data.split(":")[1])
    await state.set_state(AdminState.edit_plan_price)
    await state.update_data(edit_plan_id=plan_id)
    await callback.message.answer("Yangi narxni UZS da kiriting (Masalan: 35000):", reply_markup=get_cancel_keyboard())
    await callback.answer()


@router.message(AdminState.edit_plan_price)
async def handle_edit_plan_price_msg(message: Message, state: FSMContext):
    text = (message.text or "").strip()
    if not text.isdigit():
        return await message.answer("Faqat raqam kiriting:")
    data = await state.get_data()
    plan_id = data.get("edit_plan_id")
    await db.update_premium_plan(plan_id, price_uzs=int(text))
    await state.set_state(AdminState.in_admin_mode)
    await message.answer(f"✅ Tarif narxi {int(text):,} UZS ga yangilandi!", reply_markup=get_admin_main_menu_keyboard())


# ------------------ ADMIN: MAJBURIY OBUNA ------------------


@router.message(F.text == "📢 Majburiy obuna")
async def handle_admin_channels_menu(message: Message):
    if not await db.is_admin_allowed(message.from_user.id, "can_manage_channels"):
        return await message.answer("❌ Ruxsat yo'q.")
    channels = await db.get_required_channels(only_active=False)
    text = "📢 <b>MAJBURIY OBUNA KANALLARI:</b>\n\n"
    if not channels:
        text += "<i>Hozircha hech qanday kanal ulanmagan.</i>\n"
    else:
        for idx, ch in enumerate(channels, 1):
            st = "🟢" if ch["is_active"] else "🔴"
            text += f"{idx}. {st} <b>{html.escape(ch['title'])}</b>\nLink: {ch['invite_link']}\n\n"

    markup = InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="➕ Kanal qo'shish", callback_data="adm_add_channel")],
            [InlineKeyboardButton(text="🗑 Kanallarni boshqarish", callback_data="adm_manage_channels")],
        ]
    )
    await message.answer(text, parse_mode=ParseMode.HTML, reply_markup=markup)


@router.callback_query(F.data == "adm_add_channel")
async def handle_admin_add_channel_start(callback: CallbackQuery, state: FSMContext):
    await state.set_state(AdminState.channel_title)
    await callback.message.answer("Kanal nomini kiriting:", reply_markup=get_cancel_keyboard())
    await callback.answer()


@router.message(AdminState.channel_title)
async def handle_admin_add_channel_title(message: Message, state: FSMContext):
    title = (message.text or "").strip()
    await state.update_data(ch_title=title)
    await state.set_state(AdminState.channel_link)
    await message.answer("Kanal havolasini (link) yoki @username kiriting:")


@router.message(AdminState.channel_link)
async def handle_admin_add_channel_link(message: Message, state: FSMContext, bot: Bot):
    link = (message.text or "").strip()
    data = await state.get_data()
    title = data["ch_title"]

    username = None
    if link.startswith("@"):
        username = link.replace("@", "")
    elif "t.me/" in link:
        parts = link.split("t.me/")[-1].replace("/", "")
        if not parts.startswith("+"):
            username = parts

    channel_id = None
    if username:
        try:
            chat = await bot.get_chat(chat_id=f"@{username}")
            channel_id = chat.id
        except Exception:
            pass

    await db.add_required_channel(title, username, invite_link=link, channel_id=channel_id)
    await state.set_state(AdminState.in_admin_mode)
    await message.answer("✅ Kanal qo'shildi!", reply_markup=get_admin_main_menu_keyboard())


@router.callback_query(F.data == "adm_manage_channels")
async def handle_admin_manage_channels(callback: CallbackQuery):
    channels = await db.get_required_channels(only_active=False)
    if not channels:
        return await callback.answer("Kanallar yo'q.", show_alert=True)
    buttons = []
    for ch in channels:
        st = "O'chirish" if ch["is_active"] else "Yoqish"
        buttons.append([InlineKeyboardButton(text=f"{ch['title']} ({st})", callback_data=f"adm_tog_ch:{ch['id']}"), InlineKeyboardButton(text="🗑 O'chirish", callback_data=f"adm_del_ch:{ch['id']}")])
    await callback.message.answer("Kanallar boshqaruvi:", reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
    await callback.answer()


@router.callback_query(F.data.startswith("adm_tog_ch:"))
async def handle_admin_toggle_channel(callback: CallbackQuery):
    ch_id = int(callback.data.split(":")[1])
    await db.toggle_required_channel(ch_id)
    await callback.answer("Kanal holati o'zgartirildi!")
    try:
        await callback.message.delete()
    except Exception:
        pass


@router.callback_query(F.data.startswith("adm_del_ch:"))
async def handle_admin_delete_channel(callback: CallbackQuery):
    ch_id = int(callback.data.split(":")[1])
    await db.delete_required_channel(ch_id)
    await callback.answer("Kanal o'chirildi!")
    try:
        await callback.message.delete()
    except Exception:
        pass


# ------------------ ADMIN: OMMAVIY XABAR (BROADCAST) ------------------


@router.message(F.text == "📣 Xabar yuborish")
async def handle_admin_broadcast_start(message: Message, state: FSMContext):
    if not await db.is_admin_allowed(message.from_user.id, "can_broadcast"):
        return await message.answer("❌ Ruxsat yo'q.")
    markup = ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="Barcha foydalanuvchilar")],
            [KeyboardButton(text="Faqat o'g'il bolalar"), KeyboardButton(text="Faqat qizlar")],
            [KeyboardButton(text="Faqat Premiumlar"), KeyboardButton(text="Premium bo'lmaganlar")],
            [KeyboardButton(text="❌ Bekor qilish")],
        ],
        resize_keyboard=True,
    )
    await state.set_state(AdminState.broadcast_target)
    await message.answer("Kimlarga xabar yubormoqchisiz?", reply_markup=markup)


@router.message(AdminState.broadcast_target)
async def handle_admin_broadcast_target(message: Message, state: FSMContext):
    target = (message.text or "").strip()
    await state.update_data(broad_target=target)
    await state.set_state(AdminState.broadcast_content)
    await message.answer(f"Auditoriya: {target}\n\nXabaringizni yuboring (matn, rasm yoki video):", reply_markup=get_cancel_keyboard())


@router.message(AdminState.broadcast_content)
async def handle_admin_broadcast_content(message: Message, state: FSMContext):
    data = await state.get_data()
    target = data.get("broad_target")
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    async with await db.get_connection() as conn:
        if target == "Faqat o'g'il bolalar":
            query = "SELECT u.user_id FROM users u JOIN profiles p ON u.user_id = p.user_id WHERE u.is_banned = 0 AND p.gender = 'male'"
        elif target == "Faqat qizlar":
            query = "SELECT u.user_id FROM users u JOIN profiles p ON u.user_id = p.user_id WHERE u.is_banned = 0 AND p.gender = 'female'"
        elif target == "Faqat Premiumlar":
            query = f"SELECT user_id FROM subscriptions WHERE is_active = 1 AND end_date > '{now_str}'"
        elif target == "Premium bo'lmaganlar":
            query = f"SELECT user_id FROM users WHERE is_banned = 0 AND user_id NOT IN (SELECT user_id FROM subscriptions WHERE is_active = 1 AND end_date > '{now_str}')"
        else:
            query = "SELECT user_id FROM users WHERE is_banned = 0"

        async with conn.execute(query) as cur:
            recipients = [r["user_id"] for r in await cur.fetchall()]

    if not recipients:
        await state.set_state(AdminState.in_admin_mode)
        return await message.answer("Auditoriyada foydalanuvchilar yo'q.", reply_markup=get_admin_main_menu_keyboard())

    content_type = "text"
    file_id = None
    caption = None
    text_content = message.text

    if message.photo:
        content_type = "photo"
        file_id = message.photo[-1].file_id
        caption = message.caption
    elif message.video:
        content_type = "video"
        file_id = message.video.file_id
        caption = message.caption

    await state.update_data(recipients=recipients, b_type=content_type, b_fid=file_id, b_cap=caption, b_txt=text_content)
    await state.set_state(AdminState.confirm_broadcast)

    markup = ReplyKeyboardMarkup(keyboard=[[KeyboardButton(text="🚀 Ha, yuborish")], [KeyboardButton(text="❌ Bekor qilish")]], resize_keyboard=True)
    await message.answer(f"Xabar {len(recipients)} ta foydalanuvchiga yuborilsinmi?", reply_markup=markup)


@router.message(AdminState.confirm_broadcast)
async def handle_admin_broadcast_execution(message: Message, state: FSMContext, bot: Bot):
    if message.text != "🚀 Ha, yuborish":
        await state.set_state(AdminState.in_admin_mode)
        return await message.answer("Bekor qilindi.", reply_markup=get_admin_main_menu_keyboard())

    data = await state.get_data()
    recipients = data.get("recipients", [])
    c_type = data.get("b_type")
    f_id = data.get("b_fid")
    cap = data.get("b_cap")
    txt = data.get("b_txt")

    await state.set_state(AdminState.in_admin_mode)
    await message.answer("🚀 Xabar yuborish boshlandi...", reply_markup=get_admin_main_menu_keyboard())

    asyncio.create_task(run_broadcast_task(bot, message.from_user.id, recipients, c_type, f_id, cap, txt))


async def run_broadcast_task(bot: Bot, admin_id: int, recipients: List[int], content_type: str, file_id: Optional[str], caption: Optional[str], text: Optional[str]):
    success = 0
    failed = 0
    for uid in recipients:
        try:
            if content_type == "text":
                await bot.send_message(chat_id=uid, text=text, parse_mode=ParseMode.HTML)
            elif content_type == "photo":
                await bot.send_photo(chat_id=uid, photo=file_id, caption=caption, parse_mode=ParseMode.HTML)
            elif content_type == "video":
                await bot.send_video(chat_id=uid, video=file_id, caption=caption, parse_mode=ParseMode.HTML)
            success += 1
        except TelegramRetryAfter as e:
            await asyncio.sleep(e.retry_after)
            failed += 1
        except Exception:
            failed += 1
        await asyncio.sleep(0.04)

    try:
        await bot.send_message(chat_id=admin_id, text=f"📢 Ommaviy xabar tugadi!\nYetkazildi: {success} ta\nYetkazilmadi: {failed} ta")
    except Exception:
        pass


# ------------------ ADMIN: EKSPORT, MODERATSIYA, SOZLAMALAR, RESTART ------------------


@router.message(F.text == "📤 Eksport")
async def handle_admin_export_menu(message: Message):
    if not await db.is_admin_allowed(message.from_user.id, "can_export_data"):
        return await message.answer("❌ Ruxsat yo'q.")
    markup = InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="📊 Foydalanuvchilar (XLSX)", callback_data="adm_exp:users:xlsx"), InlineKeyboardButton(text="📄 Foydalanuvchilar (CSV)", callback_data="adm_exp:users:csv")],
            [InlineKeyboardButton(text="💳 To'lovlar (XLSX)", callback_data="adm_exp:payments:xlsx"), InlineKeyboardButton(text="📄 To'lovlar (CSV)", callback_data="adm_exp:payments:csv")],
            [InlineKeyboardButton(text="🚩 Shikoyatlar (XLSX)", callback_data="adm_exp:reports:xlsx"), InlineKeyboardButton(text="⭐ Premiumlar (XLSX)", callback_data="adm_exp:subscriptions:xlsx")],
        ]
    )
    await message.answer("Eksport formatini tanlang:", reply_markup=markup)


@router.callback_query(F.data.startswith("adm_exp:"))
async def handle_admin_export_execution(callback: CallbackQuery, bot: Bot):
    if not await db.is_admin_allowed(callback.from_user.id, "can_export_data"):
        return await callback.answer("Ruxsat yo'q!", show_alert=True)
    _, table, fmt = callback.data.split(":")
    await callback.answer("Fayl tayyorlanmoqda...")

    rows = await db.get_export_data(table)
    if not rows:
        return await callback.message.answer(f"Jadvalda ({table}) ma'lumotlar yo'q.")

    filename = f"{table}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.{fmt}"
    if fmt == "csv":
        output = io.StringIO()
        writer = csv.DictWriter(output, fieldnames=rows[0].keys())
        writer.writeheader()
        writer.writerows(rows)
        file_bytes = output.getvalue().encode("utf-8-sig")
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = table.capitalize()
        headers = list(rows[0].keys())
        ws.append(headers)

        header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")

        for r in rows:
            ws.append([str(v) if v is not None else "" for v in r.values()])

        output_b = io.BytesIO()
        wb.save(output_b)
        file_bytes = output_b.getvalue()

    await bot.send_document(chat_id=callback.from_user.id, document=BufferedInputFile(file=file_bytes, filename=filename), caption=f"📤 {table.upper()} ({fmt.upper()})")


@router.message(F.text == "🛡 Moderatsiya")
async def handle_admin_moderation_menu(message: Message):
    if not await db.is_admin_allowed(message.from_user.id, "can_manage_reports"):
        return await message.answer("❌ Ruxsat yo'q.")
    async with await db.get_connection() as conn:
        async with conn.execute("SELECT * FROM reports WHERE status = 'PENDING' ORDER BY id ASC LIMIT 5") as cur:
            reports = await cur.fetchall()

    if not reports:
        return await message.answer("✅ Yangi shikoyatlar yo'q.")

    await message.answer(f"🛡 Ko'rib chiqilishi kerak bo'lgan {len(reports)} ta shikoyat:")
    for r in reports:
        markup = InlineKeyboardMarkup(
            inline_keyboard=[
                [InlineKeyboardButton(text="👤 Profilni ko'rish", callback_data=f"adm_view_user:{r['reported_user_id']}"), InlineKeyboardButton(text="🚫 Bloklash", callback_data=f"adm_ban:{r['reported_user_id']}")],
                [InlineKeyboardButton(text="✅ Ko'rib chiqildi (Yopish)", callback_data=f"adm_close_rep:{r['id']}")],
            ]
        )
        text = f"🚩 #{r['id']} | Sabab: {html.escape(r['reason'])}\nKimdan: <code>{r['reporter_id']}</code> ➡️ Kimga: <code>{r['reported_user_id']}</code>"
        await message.answer(text, parse_mode=ParseMode.HTML, reply_markup=markup)


@router.callback_query(F.data.startswith("adm_close_rep:"))
async def handle_admin_close_report(callback: CallbackQuery):
    rep_id = int(callback.data.split(":")[1])
    async with db._lock:
        async with await db.get_connection() as conn:
            await conn.execute("UPDATE reports SET status = 'RESOLVED', resolved_by = ? WHERE id = ?", (callback.from_user.id, rep_id))
            await conn.commit()
    await callback.answer("Shikoyat yopildi!")
    try:
        await callback.message.delete()
    except Exception:
        pass


@router.message(F.text == "⚙️ Bot sozlamalari")
async def handle_admin_settings_menu(message: Message):
    if not await db.is_admin_allowed(message.from_user.id, "can_manage_settings"):
        return await message.answer("❌ Ruxsat yo'q.")
    card_number = await db.get_setting("card_number")
    card_holder = await db.get_setting("card_holder")

    markup = InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="💳 Karta raqamini o'zgartirish", callback_data="adm_set_card")],
            [InlineKeyboardButton(text="👤 Karta egasini o'zgartirish", callback_data="adm_set_cardholder")],
            [InlineKeyboardButton(text="📝 Bot matnlarini tahrirlash (CMS)", callback_data="cms_back_list")],
        ]
    )
    await message.answer(f"⚙️ <b>BOT SOZLAMALARI:</b>\n\n💳 Karta: <code>{card_number}</code>\n👤 Egasi: <code>{card_holder}</code>", parse_mode=ParseMode.HTML, reply_markup=markup)


@router.callback_query(F.data == "adm_set_card")
async def handle_admin_set_card_start(callback: CallbackQuery, state: FSMContext):
    await state.set_state(AdminState.settings_card)
    await callback.message.answer("Yangi karta raqamini kiriting:", reply_markup=get_cancel_keyboard())
    await callback.answer()


@router.message(AdminState.settings_card)
async def handle_admin_set_card_finish(message: Message, state: FSMContext):
    text = (message.text or "").strip()
    await db.set_setting("card_number", text)
    await state.set_state(AdminState.in_admin_mode)
    await message.answer(f"✅ Karta raqami yangilandi: <code>{text}</code>", parse_mode=ParseMode.HTML, reply_markup=get_admin_main_menu_keyboard())


@router.callback_query(F.data == "adm_set_cardholder")
async def handle_admin_set_cardholder_start(callback: CallbackQuery, state: FSMContext):
    await state.set_state(AdminState.settings_cardholder)
    await callback.message.answer("Karta egasining ismini kiriting:", reply_markup=get_cancel_keyboard())
    await callback.answer()


@router.message(AdminState.settings_cardholder)
async def handle_admin_set_cardholder_finish(message: Message, state: FSMContext):
    text = (message.text or "").strip()
    await db.set_setting("card_holder", text)
    await state.set_state(AdminState.in_admin_mode)
    await message.answer(f"✅ Karta egasi yangilandi: <code>{text}</code>", parse_mode=ParseMode.HTML, reply_markup=get_admin_main_menu_keyboard())


@router.message(F.text == "🧾 Admin faoliyati")
async def handle_admin_activity_log(message: Message):
    if not await db.is_admin_allowed(message.from_user.id, "can_view_audit_logs"):
        return await message.answer("❌ Ruxsat yo'q.")
    async with await db.get_connection() as conn:
        async with conn.execute("SELECT * FROM admin_logs ORDER BY id DESC LIMIT 15") as cur:
            logs = await cur.fetchall()

    if not logs:
        return await message.answer("Jurnal bo'sh.")

    text = "🧾 <b>OXIRGI ADMIN AMALLARI:</b>\n\n"
    for l in logs:
        text += f"• {l['created_at']} | Admin: <code>{l['admin_id']}</code> | {l['action']} ({l['details']})\n"
    await message.answer(text, parse_mode=ParseMode.HTML)


@router.message(F.text == "🔄 Xizmat holati")
async def handle_admin_maintenance_menu(message: Message):
    if not await db.is_admin_allowed(message.from_user.id, "can_toggle_maintenance"):
        return await message.answer("❌ Ruxsat yo'q.")
    current_status = await db.get_setting("maintenance_mode", "0")
    status_label = "🔴 TEXNIK REJIMDA" if current_status == "1" else "🟢 NORMAL ISH REJIMI"

    markup = InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="⚙️ Rejimni almashtirish", callback_data="adm_toggle_maint")],
            [InlineKeyboardButton(text="📢 Barchaga restart xabari yuborish", callback_data="adm_broadcast_restart_prompt")],
        ]
    )
    await message.answer(f"🔄 Holat: <b>{status_label}</b>", parse_mode=ParseMode.HTML, reply_markup=markup)


@router.callback_query(F.data == "adm_toggle_maint")
async def handle_admin_toggle_maintenance(callback: CallbackQuery):
    if not await db.is_admin_allowed(callback.from_user.id, "can_toggle_maintenance"):
        return await callback.answer("Ruxsat yo'q!", show_alert=True)
    cur = await db.get_setting("maintenance_mode", "0")
    new_val = "0" if cur == "1" else "1"
    await db.set_setting("maintenance_mode", new_val)
    st = "🔴 Texnik rejim YOQILDI!" if new_val == "1" else "🟢 Normal rejim TIKLANDI!"
    await callback.answer(st, show_alert=True)
    try:
        await callback.message.delete()
    except Exception:
        pass


@router.callback_query(F.data == "adm_broadcast_restart_prompt")
async def handle_admin_restart_prompt(callback: CallbackQuery, state: FSMContext):
    if not await db.is_admin_allowed(callback.from_user.id, "can_toggle_maintenance"):
        return await callback.answer("Ruxsat yo'q!", show_alert=True)
    await state.set_state(AdminState.confirm_maintenance_restart)
    markup = ReplyKeyboardMarkup(keyboard=[[KeyboardButton(text="⚠️ Ha, restart xabari yuborilsin")], [KeyboardButton(text="❌ Bekor qilish")]], resize_keyboard=True)
    await callback.message.answer("Barcha foydalanuvchilarga texnik restart haqida xabar yuborilsinmi?", reply_markup=markup)
    await callback.answer()


@router.message(AdminState.confirm_maintenance_restart)
async def handle_admin_restart_broadcast(message: Message, state: FSMContext, bot: Bot):
    if message.text != "⚠️ Ha, restart xabari yuborilsin":
        await state.set_state(AdminState.in_admin_mode)
        return await message.answer("Bekor qilindi.", reply_markup=get_admin_main_menu_keyboard())

    await state.set_state(AdminState.in_admin_mode)
    await message.answer("🚀 Xabar yuborilmoqda...", reply_markup=get_admin_main_menu_keyboard())
    async with await db.get_connection() as conn:
        async with conn.execute("SELECT user_id FROM users WHERE is_banned = 0") as cur:
            users = [r["user_id"] for r in await cur.fetchall()]

    reconnect_markup = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="🔄 Qayta ulanish", callback_data="reconnect_after_restart")]])
    for uid in users:
        try:
            await bot.send_message(chat_id=uid, text="Bot texnik xizmat uchun qayta ishga tushiriladi. Qayta ulanish tugmasini bosing:", reply_markup=reconnect_markup)
            await asyncio.sleep(0.04)
        except Exception:
            pass
    await message.answer("✅ Barchaga xabar yetkazildi.")


@router.callback_query(F.data == "reconnect_after_restart")
async def handle_reconnect_button(callback: CallbackQuery):
    user_id = callback.from_user.id
    maint = await db.get_setting("maintenance_mode", "0")
    if maint == "1" and not await db.is_any_admin(user_id):
        return await callback.answer("🛠 Texnik ishlar davom etmoqda.", show_alert=True)

    await callback.answer("✅ Ulanish tiklandi!", show_alert=True)
    try:
        await callback.message.delete()
    except Exception:
        pass
    profile = await db.get_profile(user_id)
    is_admin = await db.is_any_admin(user_id)
    if profile and profile["is_active"] == 1:
        await callback.message.answer("Asosiy menyudasiz:", reply_markup=get_main_menu_keyboard(is_admin=is_admin))
    else:
        await callback.message.answer("Boshlash uchun /start bosing.")


@router.callback_query(F.data.startswith("adm_reply_support:"))
async def handle_admin_reply_support_start(callback: CallbackQuery, state: FSMContext):
    target_id = int(callback.data.split(":")[1])
    await state.set_state(AdminState.reply_support_text)
    await state.update_data(support_reply_user_id=target_id)
    await callback.message.answer(f"Foydalanuvchi {target_id} ga javobingizni yozing:", reply_markup=get_cancel_keyboard())
    await callback.answer()


@router.message(AdminState.reply_support_text)
async def handle_admin_reply_support_finish(message: Message, state: FSMContext, bot: Bot):
    text = (message.text or "").strip()
    data = await state.get_data()
    target_id = data.get("support_reply_user_id")

    try:
        await bot.send_message(chat_id=target_id, text=f"💬 <b>ADMINISTRATOR JAVOBI:</b>\n\n{html.escape(text)}", parse_mode=ParseMode.HTML)
        await message.answer("✅ Javob yuborildi!", reply_markup=get_admin_main_menu_keyboard())
    except Exception as e:
        await message.answer(f"❌ Xatolik: {e}", reply_markup=get_admin_main_menu_keyboard())
    await state.set_state(AdminState.in_admin_mode)


# ==============================================================================
# 11. ANONIM CHAT XABARLARINI RELAY QILISH (FAQAT FAOL CHATDA BO'LSAGINA!)
# ==============================================================================


@router.message()
async def relay_chat_message(message: Message, bot: Bot):
    user_id = message.from_user.id
    active_chat = await db.get_active_chat(user_id)
    if not active_chat:
        return

    partner_id = active_chat["user2_id"] if user_id == active_chat["user1_id"] else active_chat["user1_id"]
    chat_id = active_chat["id"]

    try:
        if message.text:
            await db.log_chat_message(chat_id, user_id, "text", content=message.text)
            await bot.send_message(chat_id=partner_id, text=f"💭 {html.escape(message.text)}", parse_mode=ParseMode.HTML)
        elif message.photo:
            file_id = message.photo[-1].file_id
            caption = f"💭 {html.escape(message.caption)}" if message.caption else None
            await db.log_chat_message(chat_id, user_id, "photo", content=caption, file_id=file_id)
            await bot.send_photo(chat_id=partner_id, photo=file_id, caption=caption, parse_mode=ParseMode.HTML)
        elif message.voice:
            file_id = message.voice.file_id
            await db.log_chat_message(chat_id, user_id, "voice", file_id=file_id)
            await bot.send_voice(chat_id=partner_id, voice=file_id)
        elif message.video:
            file_id = message.video.file_id
            caption = f"💭 {html.escape(message.caption)}" if message.caption else None
            await db.log_chat_message(chat_id, user_id, "video", content=caption, file_id=file_id)
            await bot.send_video(chat_id=partner_id, video=file_id, caption=caption, parse_mode=ParseMode.HTML)
        elif message.sticker:
            file_id = message.sticker.file_id
            await db.log_chat_message(chat_id, user_id, "sticker", file_id=file_id)
            await bot.send_sticker(chat_id=partner_id, sticker=file_id)
        elif message.video_note:
            file_id = message.video_note.file_id
            await db.log_chat_message(chat_id, user_id, "video_note", file_id=file_id)
            await bot.send_video_note(chat_id=partner_id, video_note=file_id)
    except TelegramForbiddenError:
        await db.end_active_chat(chat_id, ended_by=partner_id)
        is_admin = await db.is_any_admin(user_id)
        await message.answer("❌ Suhbatdosh botni tark etdi yoki blokladi.", reply_markup=get_main_menu_keyboard(is_admin=is_admin))
    except Exception as e:
        logger.warning(f"Xabarni yetkazishda xatolik: {e}")


# ==============================================================================
# 12. DASTURNI ISHGA TUSHIRISH (RENDER 24/7)
# ==============================================================================


async def on_startup(bot: Bot):
    logger.info("Bot ishga tushirilmoqda. Ma'lumotlar bazasi tekshirilmoqda...")
    await db.init_db()
    bot_info = await bot.get_me()
    logger.info(f"Bot muvaffaqiyatli ishga tushdi: @{bot_info.username} (ID: {bot_info.id})")
    logger.info(f"Super Adminlar: {list(ADMIN_IDS)}")


async def on_shutdown(bot: Bot):
    logger.info("Bot to'xtatilmoqda...")
    await bot.session.close()


async def main():
    bot = Bot(token=BOT_TOKEN)
    storage = MemoryStorage()
    dp = Dispatcher(storage=storage)

    dp.include_router(router)
    dp.startup.register(on_startup)
    dp.shutdown.register(on_shutdown)

    await bot.delete_webhook(drop_pending_updates=True)

    try:
        logger.info("Render serverida 24/7 asinxron polling jarayoni boshlanmoqda...")
        await dp.start_polling(
            bot,
            allowed_updates=["message", "callback_query", "chat_member", "my_chat_member"],
        )
    finally:
        await bot.session.close()


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except (KeyboardInterrupt, SystemExit):
        logger.info("Dastur yakunlandi.")
