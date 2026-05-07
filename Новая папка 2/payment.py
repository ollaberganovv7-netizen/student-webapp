from aiogram import Router, F, Bot
from aiogram.types import CallbackQuery, Message, PhotoSize, LabeledPrice, PreCheckoutQuery, InlineKeyboardMarkup, InlineKeyboardButton
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup

from database.models import User
from database.db import create_payment, add_balance
from keyboards.main_kb import main_menu_kb, back_to_menu_kb
from config import ADMIN_IDS, TOPUP_OPTIONS, PAYME_TOKEN, CLICK_TOKEN, CARDS
from utils.helpers import format_price

router = Router()

class PaymentStates(StatesGroup):
    choosing_package = State()
    choosing_method = State()
    uploading_screenshot = State()
    entering_amount = State()  # for quick chek flow

def get_cards_text() -> str:
    lines = []
    for i, c in enumerate(CARDS, 1):
        lines.append(f"🏦 <b>{i}-Karta:</b> <code>{c['number']}</code>")
        lines.append(f"👤 <b>Ega:</b> {c['holder']}\n")
    return "\n".join(lines).strip()

# ─── Quick Chek (Receipt) Upload ────────────────────────────────────────────

@router.message(F.text == "📸 Chekni yuborish")
async def quick_chek_upload(message: Message, state: FSMContext, db_user: User):
    """Step 1: Ask how much they paid."""
    await state.set_state(PaymentStates.entering_amount)
    cards_text = get_cards_text()
    await message.answer(
        "📸 <b>Chekni yuborish</b>\n\n"
        f"{cards_text}\n"
        "━━━━━━━━━━━━━━━━━━━━\n"
        "💬 <b>Qancha to'lov qildingiz?</b>\n"
        "<i>(Masalan: 5000 yoki 10000)</i>",
        parse_mode="HTML"
    )

@router.message(PaymentStates.entering_amount, F.text)
async def chek_enter_amount(message: Message, state: FSMContext):
    """Step 2: Save the amount and ask for the receipt photo."""
    amount_text = message.text.strip().replace(" ", "").replace(",", "")
    try:
        amount = int(amount_text)
    except ValueError:
        await message.answer(
            "⚠️ Iltimos, faqat <b>raqam</b> yuboring. Masalan: <code>5000</code>",
            parse_mode="HTML"
        )
        return

    await state.update_data(package_key="chek", package_label="Chek", package_price=amount)
    await state.set_state(PaymentStates.uploading_screenshot)
    await message.answer(
        f"✅ Summa qabul qilindi: <b>{amount:,} so'm</b>\n\n"
        "📸 <b>Endi chek rasmini yuboring:</b>",
        parse_mode="HTML"
    )

# ─── Step 1: Select Amount ──────────────────────────────────────────────────

@router.message(F.text == "💳 Hisobni to'ldirish")
async def payment_start(message: Message, state: FSMContext, db_user: User):
    buttons = [
        [InlineKeyboardButton(text=info["label"], callback_data=f"pkg:{key}")]
        for key, info in TOPUP_OPTIONS.items()
    ]
    buttons.append([InlineKeyboardButton(text="🏠 Bosh menyu", callback_data="main_menu")])

    await state.set_state(PaymentStates.choosing_package)
    await message.answer(
        "💳 <b>Hisobni to'ldirish miqdorini tanlang:</b>\n\n" +
        "\n".join(f"• <b>{info['label']}</b>" for info in TOPUP_OPTIONS.values()),
        reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons),
        parse_mode="HTML",
    )

# ─── Step 2: Select Method ──────────────────────────────────────────────────

@router.callback_query(F.data.startswith("pkg:"), PaymentStates.choosing_package)
async def choose_package(callback: CallbackQuery, state: FSMContext):
    pkg_key = callback.data.split(":")[1]
    pkg = TOPUP_OPTIONS.get(pkg_key)
    if not pkg:
        await callback.answer("Noma'lum paket", show_alert=True)
        return

    await state.update_data(
        package_key=pkg_key,
        package_label=pkg["label"],
        package_price=pkg["amount"],
    )
    
    await state.set_state(PaymentStates.choosing_method)
    
    method_buttons = [
        [InlineKeyboardButton(text="📲 Payme orqali (Avtomat)", callback_data="method:payme")],
        [InlineKeyboardButton(text="📲 Click orqali (Avtomat)", callback_data="method:click")],
        [InlineKeyboardButton(text="💳 Karta orqali (Skrinshot)", callback_data="method:manual")],
        [InlineKeyboardButton(text="⬅️ Orqaga", callback_data="payment_retry")]
    ]
    
    await callback.message.edit_text(
        f"💳 <b>To'lov usulini tanlang:</b>\n\n"
        f"📦 Paket: <b>{pkg['label']}</b>\n"
        f"💰 Summa: <b>{format_price(pkg['amount'])}</b>",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=method_buttons),
        parse_mode="HTML"
    )
    await callback.answer()

@router.callback_query(F.data == "payment_retry")
async def retry_payment(callback: CallbackQuery, state: FSMContext, db_user: User):
    await payment_start(callback.message, state, db_user)
    await callback.answer()

# ─── Step 3: Handle Selection ───────────────────────────────────────────────

@router.callback_query(F.data.startswith("method:"), PaymentStates.choosing_method)
async def handle_method(callback: CallbackQuery, state: FSMContext, bot: Bot):
    method = callback.data.split(":")[1]
    data = await state.get_data()
    price = data.get("package_price", 0)
    label = data.get("package_label", "Balansni to'ldirish")

    if method == "manual":
        await state.set_state(PaymentStates.uploading_screenshot)
        cards_text = get_cards_text()
        text = (
            f"💳 <b>Karta orqali to'lov</b>\n\n"
            f"Quyidagi kartalardan biriga o'tkazing:\n\n"
            f"{cards_text}\n"
            f"💰 <b>Summa:</b> {format_price(price)}\n\n"
            f"📸 <b>To'lov chekining skrinshotini yuboring:</b>"
        )
        await callback.message.edit_text(text, reply_markup=back_to_menu_kb(), parse_mode="HTML")
    
    elif method in ["payme", "click"]:
        token = PAYME_TOKEN if method == "payme" else CLICK_TOKEN
        if not token:
            await callback.answer(f"⚠️ {method.title()} to'lov tizimi hozircha ulanmagan. Iltimos, Karta orqali to'lang.", show_alert=True)
            return

        await callback.message.delete()
        await bot.send_invoice(
            chat_id=callback.message.chat.id,
            title=label,
            description="Bot balansini to'ldirish",
            payload=f"topup_{data.get('package_key')}",
            provider_token=token,
            currency="UZS",
            prices=[LabeledPrice(label=label, amount=price * 100)], # amount is in smallest units (tiyin)
            start_parameter="topup",
            need_name=False,
            need_phone_number=False,
            need_email=False,
            is_flexible=False
        )
    await callback.answer()

# ─── Manual Payment (Screenshot) ───────────────────────────────────────────

@router.message(PaymentStates.uploading_screenshot, F.photo)
async def receive_screenshot(message: Message, state: FSMContext, db_user: User, bot: Bot):
    data = await state.get_data()
    await state.clear()

    photo: PhotoSize = message.photo[-1]
    file_id = photo.file_id

    pkg_key = data.get("package_key", "")
    pkg_label = data.get("package_label", "")
    pkg_price = data.get("package_price", 0)

    payment = await create_payment(user_id=db_user.id, amount=pkg_price, package=pkg_key, screenshot_file_id=file_id)

    await message.answer(
        "✅ <b>Chek muvaffaqiyatli yuborildi!</b>\n\n"
        "⏳ Admin tekshirib, tez orada tasdiqlaydi.",
        reply_markup=main_menu_kb(),
        parse_mode="HTML",
    )

    # Notify admins
    from keyboards.admin_kb import payment_action_kb
    admin_text = (
        f"💳 <b>Yangi to'lov so'rovi!</b>\n\n"
        f"👤 Foydalanuvchi: {db_user.full_name}\n"
        f"🆔 ID: <code>{db_user.id}</code>\n"
        f"💰 Summa: {format_price(pkg_price)}\n"
        f"🔑 ID: #{payment.id}"
    )

    for admin_id in ADMIN_IDS:
        try:
            await bot.send_photo(chat_id=admin_id, photo=file_id, caption=admin_text, reply_markup=payment_action_kb(payment.id, db_user.id), parse_mode="HTML")
        except: pass

# ─── Automatic Payment Handlers ───────────────────────────────────────────

@router.pre_checkout_query()
async def process_pre_checkout_query(pre_checkout_query: PreCheckoutQuery):
    await pre_checkout_query.answer(ok=True)

@router.message(F.successful_payment)
async def successful_payment(message: Message, bot: Bot):
    payment_info = message.successful_payment
    amount = payment_info.total_amount // 100 # Convert back from tiyin to sum
    
    await add_balance(message.from_user.id, amount)
    
    await message.answer(
        f"🎉 <b>Muvaffaqiyatli to'lov!</b>\n\n"
        f"💰 Balansingizga <b>{format_price(amount)}</b> qo'shildi.\n"
        f"🚀 Endi xizmatlardan foydalanishingiz mumkin!",
        reply_markup=main_menu_kb(),
        parse_mode="HTML"
    )
    
    # Notify admins about auto-payment
    admin_text = (
        f"💸 <b>AVTOMATIK TO'LOV!</b>\n\n"
        f"👤 Foydalanuvchi: {message.from_user.full_name}\n"
        f"🆔 ID: <code>{message.from_user.id}</code>\n"
        f"💰 Summa: {format_price(amount)}"
    )
    for admin_id in ADMIN_IDS:
        try: await bot.send_message(admin_id, admin_text, parse_mode="HTML")
        except: pass
