from aiogram import Router, F
from aiogram.filters import CommandStart, Command
from aiogram.types import Message, CallbackQuery
from aiogram.fsm.context import FSMContext

from database.models import User
from keyboards.main_kb import main_menu_kb, back_to_menu_kb
from utils.helpers import has_access, format_price
from config import PRICING

router = Router()

WELCOME_TEXT = """
✨ <b>PROFESSIONAL STUDENT BOT</b> ✨
────────────────────────
Sizning akademik muvaffaqiyatingiz uchun tayyor yechimlar.

<b>Xizmatlarimiz:</b>
💎 <b>Taqdimotlar</b> — <i>Premium dizayn (PPTX)</i>
📝 <b>Referatlar</b> — <i>Sifatli akademik matn (DOCX)</i>
📄 <b>Mustaqil ishlar</b> — <i>Tayyor shablon (DOCX)</i>
📘 <b>Kurs ishlari</b> — <i>Chuqur tahliliy yondashuv (DOCX)</i>

────────────────────────
🚀 <b>Kerakli bo'limni tanlang:</b>
"""

ACCOUNT_TEXT = """
👤 <b>Mening hisobim</b>

🆔 ID: <code>{user_id}</code>
👤 Ism: {full_name}
💰 Balans: <b>{balance} so'm</b>
"""


@router.message(CommandStart())
async def cmd_start(message: Message, db_user: User, state: FSMContext):
    await state.clear()
    await message.answer(
        WELCOME_TEXT,
        reply_markup=main_menu_kb(),
        parse_mode="HTML",
    )


@router.message(F.text == "🏠 Bosh menyu")
async def msg_main_menu(message: Message, state: FSMContext):
    await state.clear()
    await message.answer(
        WELCOME_TEXT,
        reply_markup=main_menu_kb(),
        parse_mode="HTML",
    )

@router.message(F.text == "❌ Bekor qilish")
async def msg_global_cancel(message: Message, state: FSMContext):
    await state.clear()
    await message.answer(
        "❌ Amaliyot bekor qilindi.\n\n" + WELCOME_TEXT,
        reply_markup=main_menu_kb(),
        parse_mode="HTML",
    )

@router.callback_query(F.data == "main_menu")
async def cb_main_menu(callback: CallbackQuery, db_user: User, state: FSMContext):
    await state.clear()
    await callback.message.delete()
    await callback.message.answer(
        WELCOME_TEXT,
        reply_markup=main_menu_kb(),
        parse_mode="HTML",
    )
    await callback.answer()


@router.message(F.text == "📊 Mening hisobim")
async def msg_my_account(message: Message, db_user: User):
    await _show_account(message, db_user)

@router.callback_query(F.data == "my_account")
async def cb_my_account(callback: CallbackQuery, db_user: User):
    await _show_account(callback.message, db_user, is_callback=True)
    await callback.answer()

async def _show_account(msg: Message, db_user: User, is_callback=False):
    text = ACCOUNT_TEXT.format(
        user_id=db_user.id,
        full_name=db_user.full_name or "—",
        balance=format_price(db_user.balance or 0)
    )
    if is_callback:
        await msg.edit_text(text, reply_markup=back_to_menu_kb(), parse_mode="HTML")
    else:
        await msg.answer(text, reply_markup=back_to_menu_kb(), parse_mode="HTML")


@router.message(Command("stop"))
async def cmd_stop(message: Message, state: FSMContext):
    await state.clear()
    await message.answer("🛑 <b>Jarayon to'xtatildi.</b>\n\nAsosiy menyuga qaytdik.", reply_markup=main_menu_kb(), parse_mode="HTML")

@router.message(Command("buy"))
async def cmd_buy(message: Message, state: FSMContext, db_user: User):
    from handlers.payment import payment_start
    await payment_start(message, state, db_user)

@router.message(Command("help"))
async def cmd_help(message: Message):
    await _show_help(message)

async def _show_help(msg: Message, is_callback=False):
    help_text = """
ℹ️ <b>Yordam markazi</b>

<b>Qanday ishlaydi?</b>
1. Xizmatni tanlang (Referat, Taqdimot va hk)
2. Mavzu kiriting va sozlamalarni tanlang
3. Fayl tayyor bo'ladi va yuklab olasiz

💰 <b>To'lov:</b> /buy buyrug'i orqali hisobni to'ldiring.
🆘 <b>Admin bilan bog'lanish:</b> {admin}
"""
    from config import ADMIN_USERNAME
    text = help_text.format(admin=ADMIN_USERNAME)
    
    if is_callback:
        await msg.edit_text(text, reply_markup=back_to_menu_kb(), parse_mode="HTML")
    else:
        await msg.answer(text, reply_markup=back_to_menu_kb(), parse_mode="HTML")

@router.callback_query(F.data == "help")
async def cb_help(callback: CallbackQuery):
    await _show_help(callback.message, is_callback=True)
    await callback.answer()

