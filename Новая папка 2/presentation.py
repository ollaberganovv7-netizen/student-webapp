import asyncio
import io
from aiogram import Router, F, Bot
from aiogram.types import (
    CallbackQuery, Message, BufferedInputFile,
    ReplyKeyboardMarkup, KeyboardButton, ReplyKeyboardRemove,
    InlineKeyboardMarkup, InlineKeyboardButton
)
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup

from database.models import User
from database.db import mark_free_used, create_request, deduct_balance
from keyboards.presentation_kb import (
    language_kb, style_kb, back_to_menu_kb, quality_kb, 
    chapters_kb, slides_grid_kb, design_selection_kb, summary_kb
)
from keyboards.main_kb import main_menu_kb
from services.ai_service import generate_presentation_content, generate_akademik_content, client, OPENAI_MODEL
from services.pptx_service import generate_pptx, analyze_template
from services.akademik_pptx_service import generate_akademik_pptx
from utils.helpers import format_price, is_free_trial, safe_topic, get_template_path
from config import PRICING, PRES_LANGUAGES, PRES_STYLES, ADMIN_IDS
import json

router = Router()

class PresentationStates(StatesGroup):
    choosing_language    = State()
    entering_topic       = State()
    reviewing_summary    = State()
    entering_manual_plan = State()   # old (unused now)
    choosing_plan_count  = State()   # new: how many chapters?
    entering_plan_step   = State()   # new: enter each chapter title one by one
    choosing_style       = State()
    waiting_for_photos   = State()   # user uploading their own images

# ─── Step 1: Quality selection ───────────────────────────────────────────────

@router.message(F.text.in_(["🆕 Taqdimot (Slayd) Yaratish", "🚀 Slayd Pro (Premium)"]))
async def start_presentation(message: Message, db_user: User, state: FSMContext):
    quality = "premium" if "Pro" in message.text else "standard"
    is_admin = db_user.id in ADMIN_IDS
    balance = db_user.balance or 0

    # Prices
    if quality == "premium":
        price_low = PRICING["presentation_pre_low"]
        price_high = PRICING["presentation_pre_high"]
        quality_label = "💎 <b>Premium (Pro)</b>"
        quality_desc = "Chuqur ilmiy tahlil, manbalar [1][2], 1500+ so'z"
    else:
        price_low = PRICING["presentation_std_low"]
        price_high = PRICING["presentation_std_high"]
        quality_label = "✨ <b>Standart</b>"
        quality_desc = "Sifatli akademik taqdimot, 1200+ so'z"

    can_afford = is_admin or balance >= price_low

    # Build keyboard
    from aiogram.types import InlineKeyboardMarkup, InlineKeyboardButton
    if can_afford:
        kb = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="✅ Davom etish", callback_data=f"pres_confirm:{quality}")],
            [InlineKeyboardButton(text="❌ Bekor qilish", callback_data="pres_cancel")]
        ])
        status_line = f"✅ <b>Balansingiz yetarli!</b>"
    else:
        needed = price_low - balance
        kb = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="💳 Balans to'ldirish", callback_data="topup_menu")],
            [InlineKeyboardButton(text="❌ Bekor qilish", callback_data="pres_cancel")]
        ])
        status_line = f"⚠️ <b>Balansingiz yetarli emas!</b>\n🔴 Yetishmayapti: <b>{format_price(needed)}</b>"

    balance_display = "🛡️ Admin (tekin)" if is_admin else format_price(balance)

    await message.answer(
        f"{quality_label}\n"
        f"<i>{quality_desc}</i>\n\n"
        f"━━━━━━━━━━━━━━━━\n"
        f"💰 <b>Narx:</b> {format_price(price_low)} — {format_price(price_high)}\n"
        f"   <i>(slaydlar soniga qarab)</i>\n"
        f"👛 <b>Sizning balansingiz:</b> {balance_display}\n"
        f"━━━━━━━━━━━━━━━━\n"
        f"{status_line}",
        reply_markup=kb,
        parse_mode="HTML"
    )

@router.callback_query(F.data.startswith("pres_confirm:"))
async def confirm_presentation(callback: CallbackQuery, state: FSMContext):
    quality = callback.data.split(":")[1]
    await state.update_data(quality=quality)
    await state.set_state(PresentationStates.choosing_language)
    await callback.message.edit_text(
        f"📊 <b>Taqdimot yaratish ({'Premium' if quality == 'premium' else 'Standart'})</b>\n\n"
        "1️⃣ <b>Tilni tanlang:</b>",
        reply_markup=language_kb(),
        parse_mode="HTML"
    )
    await callback.answer()

@router.callback_query(F.data == "pres_cancel")
async def cancel_presentation_inline(callback: CallbackQuery, state: FSMContext):
    # Set cancelled flag so running generation can detect it
    await state.update_data(cancelled=True)
    await state.clear()
    await callback.message.edit_text("❌ Amaliyot bekor qilindi.")
    await callback.answer()

@router.callback_query(F.data == "pres_cancel_generating")
async def cancel_generating(callback: CallbackQuery, state: FSMContext):
    """Cancel during active generation — sets the shared cancel flag."""
    data = await state.get_data()
    cancel_flag = data.get("cancel_flag")
    if cancel_flag and isinstance(cancel_flag, dict):
        cancel_flag["cancelled"] = True
    await state.clear()
    await callback.message.edit_text("❌ Yaratish to'xtatildi. Hech narsa hisobdan yechilmadi.")
    await callback.answer()

@router.callback_query(F.data == "topup_menu")
async def topup_from_presentation(callback: CallbackQuery, state: FSMContext):
    await state.clear()
    from keyboards.main_kb import main_menu_kb
    await callback.message.edit_text(
        "💳 <b>Balans to'ldirish</b>\n\n"
        "Quyidagi tugmani bosing va balans to'ldiring:",
        parse_mode="HTML"
    )
    await callback.answer("💳 Balans bo'limiga o'tmoqdasiz...")
    # Trigger /buy command flow
    await callback.message.answer(
        "💳 Balansni to'ldirish uchun /buy buyrug'ini yuboring yoki pastdagi menyudan foydalaning.",
        reply_markup=main_menu_kb()
    )


# ─── Step 2: Language selected ───────────────────────────────────────────────

@router.callback_query(F.data.startswith("pres_lang:"), PresentationStates.choosing_language)
async def choose_language(callback: CallbackQuery, state: FSMContext):
    lang_code = callback.data.split(":")[1]
    await state.update_data(language=lang_code)
    await state.set_state(PresentationStates.entering_topic)
    await callback.message.edit_text(
        "2️⃣ <b>Taqdimot mavzusini kiriting:</b>\n"
        "<i>(Masalan: O'zbekiston tarixi yoki Marketing)</i>",
        reply_markup=back_to_menu_kb(),
        parse_mode="HTML"
    )
    await callback.answer()

# ─── Step 2: Topic entered ────────────────────────────────────────────────────

@router.message(PresentationStates.entering_topic)
async def enter_topic(message: Message, state: FSMContext, db_user: User):
    topic = safe_topic(message.text or "")
    if not topic:
        await message.answer("❗ Mavzu bo'sh bo'lishi mumkin emas.")
        return
    
    # Set default settings
    await state.update_data(
        topic=topic,
        language="uz",
        num_slides=10,
        quality="standard",
        num_chapters=4,
        author=db_user.full_name or "Foydalanuvchi"
    )
    
    await state.set_state(PresentationStates.reviewing_summary)
    await show_summary(message, state, db_user)

async def show_summary(message: Message, state: FSMContext, db_user: User):
    data = await state.get_data()
    topic = data.get("topic")
    lang = data.get("language", "uz")
    slides = data.get("num_slides", 10)
    quality = data.get("quality", "standard")
    author = data.get("author", "Foydalanuvchi")
    
    lang_label = {"uz": "O'zbekcha 🇺🇿", "ru": "Русский 🇷🇺", "en": "English 🇺🇸"}.get(lang, lang)
    quality_label = "💎 Premium" if quality == "premium" else "✨ Standart"
    
    text = (
        f"📊 <b>Taqdimot Haqida | {quality_label}</b>\n\n"
        f"🗺️ <b>Taqdimot tili:</b> {lang_label}\n"
        f"👤 <b>Taqdimot Muallifi:</b> {author}\n"
        f"🎞️ <b>Listlari Soni:</b> {slides} ta\n"
        f"📂 <b>Mundarija bo'limlari:</b> {data.get('num_chapters', 4)} ta\n"
        f"💎 <b>Yaratish usuli:</b> Mavzu asosida\n\n"
        f"<i>Ushbu sozlamalar asosida slaydingiz yaratiladi. \"Sozlamalar\" tugmasi orqali ularni o'zgartirishingiz mumkin.</i>"
    )
    
    await message.answer(
        text,
        reply_markup=summary_kb(db_user.balance or 0, db_user.full_name or "", quality),
        parse_mode="HTML"
    )

@router.message(F.web_app_data, PresentationStates.reviewing_summary)
async def handle_webapp_data_summary(message: Message, state: FSMContext, db_user: User):
    try:
        # Delete the service "You sent data" message to keep chat clean
        try: await message.delete() 
        except: pass

        data = json.loads(message.web_app_data.data)
        if "template_id" in data:
            # Design selected
            await state.update_data(style=data["template_id"], style_name=data["template_name"])
        elif data.get("action") == "plan_update":
            mode = data.get("mode")
            if mode == "auto":
                num = int(data.get("num_chapters"))
                await state.update_data(num_chapters=num, manual_plan=None)
                await show_summary(message, state, db_user, edit=True)
            else:
                # Manual mode: plan text already built in webapp
                plan = data.get("manual_plan", "")
                num = int(data.get("num_chapters", len(plan.split('\n'))))
                await state.update_data(manual_plan=plan, num_chapters=num)
                await show_summary(message, state, db_user, edit=True)
            return
        elif data.get("action") == "content_update":
            await state.update_data(
                extra_info=data.get("extra_info"),
                tone=data.get("tone"),
                add_images=data.get("add_images")
            )
        else:
            # Settings selected
            slides = int(data.get("slides", 10))
            update_dict = dict(
                author=data.get("author"),
                language=data.get("language"),
                num_slides=slides,
                quality="premium" if data.get("premium") else "standard",
                ai_images=bool(data.get("ai_images", False))
            )
            # Allow topic editing from WebApp
            if data.get("topic"):
                update_dict["topic"] = data["topic"].strip()
            await state.update_data(**update_dict)
        
        await show_summary(message, state, db_user, edit=True)
    except Exception as e:
        print(f"Error in webapp data: {e}")

async def show_summary(message: Message, state: FSMContext, db_user: User, edit: bool = False):
    data = await state.get_data()
    topic = data.get("topic")
    lang = data.get("language", "uz")
    slides = data.get("num_slides", 10)
    quality = data.get("quality", "standard")
    author = data.get("author", "Foydalanuvchi")
    style_name = data.get("style_name", "Standart (Oq fon)")
    ai_images_count = data.get("ai_images_count", 0)
    
    lang_label = {"uz": "O'zbekcha 🇺🇿", "ru": "Русский 🇷🇺", "en": "English 🇺🇸"}.get(lang, lang)
    quality_label = "💎 Premium" if quality == "premium" else "✨ Standart"
    
    user_photos  = data.get("user_photos", [])   # list of file_ids
    user_photos_label = f"{len(user_photos)} ta rasm" if user_photos else "Yo'q"
    ai_img_label = f"🎨 {ai_images_count} ta (+{format_price(ai_images_count * 1000)})" if ai_images_count > 0 else "Yo'q"
    
    text = (
        f"📊 <b>Taqdimot Haqida | {quality_label}</b>\n\n"
        f"📝 <b>Mavzu:</b> {topic}\n"
        f"🗺️ <b>Taqdimot tili:</b> {lang_label}\n"
        f"👤 <b>Taqdimot Muallifi:</b> {author}\n"
        f"🎞️ <b>Listlari Soni:</b> {slides} ta\n"
        f"🎨 <b>Dizayn:</b> {style_name}\n"
        f"📂 <b>Reja bo'limlari:</b> {data.get('num_chapters', 'AI')}\n"
        f"🖼 <b>AI rasmlar:</b> {ai_img_label}\n"
        f"📷 <b>Yuklangan rasmlar:</b> {user_photos_label}\n\n"
        f"<i>Barcha sozlamalar tayyor bo'lsa, \"Yaratish\" tugmasini bosing.</i>"
    )
    
    quality = data.get("quality", "standard")
    kb = summary_kb(db_user.balance or 0, db_user.full_name or "", quality, topic=topic or "")
    
    summary_id = data.get("summary_msg_id")
    if summary_id:
        try: await message.bot.delete_message(chat_id=message.chat.id, message_id=summary_id)
        except: pass

    msg = await message.answer(text, reply_markup=kb, parse_mode="HTML")
    await state.update_data(summary_msg_id=msg.message_id)

# ─── Step-by-step manual plan input ─────────────────────────────────────────

@router.message(F.text == "❌ Bekor qilish", PresentationStates.entering_plan_step)
async def cancel_plan_step(message: Message, state: FSMContext, db_user: User):
    await state.set_state(PresentationStates.reviewing_summary)
    await message.answer("❌ Reja kiritish bekor qilindi.", reply_markup=ReplyKeyboardRemove())
    await show_summary(message, state, db_user)

@router.message(F.text, PresentationStates.entering_plan_step)
async def receive_plan_step(message: Message, state: FSMContext, db_user: User):
    text = (message.text or "").strip()
    data = await state.get_data()
    chapters = data.get("manual_plan_chapters", [])
    total = data.get("manual_plan_total", 4)

    chapters.append(text)
    await state.update_data(manual_plan_chapters=chapters)

    current = len(chapters)

    if current < total:
        # Ask for next chapter
        await message.answer(
            f"✅ <b>{current}-bo'lim saqlandi:</b> {text}\n\n"
            f"📌 <b>{current + 1}-bo'lim nomini yozing:</b>\n"
            f"<i>({current}/{total} kiritildi)</i>",
            parse_mode="HTML"
        )
    else:
        # All chapters entered — build the plan
        plan_text = "\n".join(chapters)
        await state.update_data(
            manual_plan=plan_text,
            num_chapters=total,
            manual_plan_chapters=[]
        )
        await state.set_state(PresentationStates.reviewing_summary)
        await message.answer(
            f"✅ <b>Reja tayyor!</b>\n\n"
            + "\n".join([f"{i+1}. {c}" for i, c in enumerate(chapters)]),
            parse_mode="HTML",
            reply_markup=ReplyKeyboardRemove()
        )
        await show_summary(message, state, db_user)

# ─── AI Image count selection ─────────────────────────────────────────────────

@router.message(F.text == "🖼 AI Rasm", PresentationStates.reviewing_summary)
async def ask_ai_image_count(message: Message, state: FSMContext):
    from aiogram.types import InlineKeyboardMarkup, InlineKeyboardButton
    data = await state.get_data()
    current = data.get("ai_images_count", 0)
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [
            InlineKeyboardButton(text="0 (Yo'q)", callback_data="ai_img:0"),
            InlineKeyboardButton(text="1 ta", callback_data="ai_img:1"),
            InlineKeyboardButton(text="2 ta", callback_data="ai_img:2"),
        ],
        [
            InlineKeyboardButton(text="3 ta", callback_data="ai_img:3"),
            InlineKeyboardButton(text="4 ta", callback_data="ai_img:4"),
            InlineKeyboardButton(text="5 ta", callback_data="ai_img:5"),
        ],
    ])
    await message.answer(
        f"🖼 <b>AI rasmlar sonini tanlang</b>\n\n"
        f"Hozirgi: <b>{current} ta</b>\n"
        f"💰 Har bir rasm: <b>1 000 so'm</b>\n\n"
        f"<i>Rasmlar faqat kontent slaydlariga qo'yiladi\n"
        f"(Reja va Xulosa slaydlariga qo'yilmaydi)</i>",
        reply_markup=kb,
        parse_mode="HTML"
    )

@router.callback_query(F.data.startswith("ai_img:"), PresentationStates.reviewing_summary)
async def set_ai_image_count(callback: CallbackQuery, state: FSMContext, db_user: User):
    count = int(callback.data.split(":")[1])
    await state.update_data(ai_images_count=count)
    label = f"{count} ta (+{count * 1000} so'm)" if count > 0 else "Yo'q"
    await callback.message.edit_text(f"✅ AI rasmlar: <b>{label}</b>", parse_mode="HTML")
    await callback.answer()
    await show_summary(callback.message, state, db_user)

# ─── Photo upload flow ────────────────────────────────────────────────────────

@router.message(F.text == "📷 Rasm yuklash", PresentationStates.reviewing_summary)
async def ask_for_photos(message: Message, state: FSMContext):
    await state.set_state(PresentationStates.waiting_for_photos)
    data = await state.get_data()
    existing = data.get("user_photos", [])
    await message.answer(
        f"📷 <b>Rasmlaringizni yuboring</b>\n\n"
        f"Hozir saqlangan: {len(existing)} ta rasm\n\n"
        "Rasm yuboring — bot ularni slayдlarga joylashtiradi.\n"
        "<i>Tugagach ✅ Tayyor tugmasini bosing.</i>",
        reply_markup=ReplyKeyboardMarkup(
            keyboard=[[KeyboardButton(text="✅ Tayyor")], [KeyboardButton(text="🗑 Rasmlarni o'chirish")]],
            resize_keyboard=True
        ),
        parse_mode="HTML"
    )

@router.message(F.photo, PresentationStates.waiting_for_photos)
async def receive_photo(message: Message, state: FSMContext):
    data = await state.get_data()
    photos = data.get("user_photos", [])
    # Get the best quality photo (last in array)
    file_id = message.photo[-1].file_id
    photos.append(file_id)
    await state.update_data(user_photos=photos)
    await message.answer(f"✅ Rasm qabul qilindi! Jami: {len(photos)} ta\n<i>Yana yuborishingiz yoki ✅ Tayyor bosishingiz mumkin.</i>", parse_mode="HTML")

@router.message(F.text == "🗑 Rasmlarni o'chirish", PresentationStates.waiting_for_photos)
async def clear_photos(message: Message, state: FSMContext):
    await state.update_data(user_photos=[])
    await message.answer("🗑 Barcha rasmlar o'chirildi.")

@router.message(F.text == "✅ Tayyor", PresentationStates.waiting_for_photos)
async def done_photos(message: Message, state: FSMContext, db_user: User):
    await state.set_state(PresentationStates.reviewing_summary)
    data = await state.get_data()
    count = len(data.get("user_photos", []))
    await message.answer(
        f"✅ {count} ta rasm saqlandi! Endi taqdimotni yaratasiz.",
        reply_markup=ReplyKeyboardRemove()
    )
    await show_summary(message, state, db_user)

@router.message(F.text, PresentationStates.waiting_for_photos)
async def fallback_text_photos(message: Message):
    await message.answer(
        "⚠️ <b>Iltimos, rasmlarni oddiy fayl yoki rasm shaklida yuboring!</b>\n\n"
        "Agar AI o'zi rasm chizishini xohlasangiz, hozir shunchaki <b>✅ Tayyor</b> tugmasini bosing "
        "va keyingi bosqichda <b>🖼 AI Rasm</b> tugmasini tanlang.",
        parse_mode="HTML"
    )

@router.message(F.text == "✅ Yaratish", PresentationStates.reviewing_summary)
async def start_generation_text(message: Message, state: FSMContext, db_user: User):
    data = await state.get_data()
    style = data.get("style", "classic")
    await finalize_presentation_logic(message, state, db_user, style)

@router.message(F.text == "❌ Bekor qilish", PresentationStates.reviewing_summary)
async def cancel_generation(message: Message, state: FSMContext):
    await state.clear()
    await message.answer("❌ Amaliyot bekor qilindi.", reply_markup=main_menu_kb())

@router.message(F.text, PresentationStates.reviewing_summary)
async def fallback_text_summary(message: Message):
    await message.answer(
        "⚠️ <b>Iltimos, quyidagi tugmalardan birini tanlang!</b>\n\n"
        "Taqdimotni yaratish uchun <b>✅ Yaratish</b> tugmasini bosing.",
        parse_mode="HTML"
    )

@router.callback_query(F.data == "pres_start_generation", PresentationStates.reviewing_summary)
async def start_generation_callback(callback: CallbackQuery, state: FSMContext, db_user: User):
    data = await state.get_data()
    style = data.get("style", "classic")
    await finalize_presentation_logic(callback.message, state, db_user, style)
    await callback.answer()

async def finalize_presentation_logic(message: Message, state: FSMContext, db_user: User, style: str):
    data = await state.get_data()
    topic = data.get("topic")
    num_slides = data.get("num_slides") or 10
    num_slides = int(num_slides)  # ensure int
    quality = data.get("quality", "standard")
    language = data.get("language", "uz")
    manual_plan = data.get("manual_plan")
    num_chapters = data.get("num_chapters", 4)
    print(f"DEBUG: num_slides={num_slides}, num_chapters={num_chapters}, quality={quality}")

    # Calculate price
    ai_images_count = int(data.get("ai_images_count", 0))
    if quality == "premium":
        if num_slides <= 12: price = PRICING["presentation_pre_low"]
        elif num_slides <= 20: price = PRICING["presentation_pre_mid"]
        else: price = PRICING["presentation_pre_high"]
    else:
        if num_slides <= 12: price = PRICING["presentation_std_low"]
        elif num_slides <= 20: price = PRICING["presentation_std_mid"]
        else: price = PRICING["presentation_std_high"]
    
    # Dynamic image pricing: 1 image = 1000 so'm
    price += ai_images_count * 1000

    # Admin check
    is_admin = db_user.id in ADMIN_IDS

    if not is_admin and (db_user.balance or 0) < price:
        await message.answer(
            f"⚠️ <b>Mablag' yetarli emas!</b>\n"
            f"Narxi: {format_price(price)}\n"
            f"Sizda: {format_price(db_user.balance or 0)}",
            reply_markup=main_menu_kb()
        )
        return

    price_note = "🛡️ Admin uchun tekin" if is_admin else format_price(price)
    ai_img_line = f"\n🎨 <b>AI rasmlar:</b> {ai_images_count} ta" if ai_images_count > 0 else ""

    # Use a dict to track cancellation (mutable, shared across async tasks)
    cancel_flag = {"cancelled": False}
    
    # Store cancel flag reference in state so cancel handler can set it
    await state.update_data(cancel_flag=cancel_flag, generating=True)
    
    cancel_kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="❌ Bekor qilish", callback_data="pres_cancel_generating")]
    ])
    
    wait_msg = await message.answer(
        f"🤖 <b>AI taqdimot yaratmoqda...</b>\n\n"
        f"📝 <b>Mavzu:</b> {topic}\n"
        f"🎞️ <b>Slaydlar:</b> {num_slides} ta\n"
        f"💎 <b>Sifat:</b> {'Premium' if quality == 'premium' else 'Standart'}"
        f"{ai_img_line}\n"
        f"👤 <b>Muallif:</b> {db_user.full_name or 'Foydalanuvchi'}\n"
        f"────────────────────\n"
        f"💰 <b>Xarajat:</b> {price_note}\n\n"
        f"⏳ <i>AI har bir slaydni chuqur o'rganib yozadi.\n"
        f"Bu jarayon 4-5 daqiqa davom etadi — sifat uchun!</i>",
        parse_mode="HTML",
        reply_markup=cancel_kb
    )

    try:
        # Helper to check if user cancelled
        def is_cancelled():
            return cancel_flag.get("cancelled", False)

        # Pass manual plan to AI instructions via topic or extra details
        full_topic = topic
        forced_titles = None

        if manual_plan:
            # Parse chapter titles from plan ("1. Kirish" → "Kirish")
            raw_lines = [l.strip() for l in manual_plan.strip().split('\n') if l.strip()]
            chapter_titles = []
            for line in raw_lines:
                import re
                clean = re.sub(r'^\d+\.\s*', '', line).strip()
                if clean:
                    chapter_titles.append(clean)
            
            # We need to fill (num_slides - 2) content slides (excluding Reja and Xulosa)
            content_slides_needed = num_slides - 2
            if content_slides_needed < len(chapter_titles):
                # If they asked for fewer slides than chapters, truncate chapters
                chapter_titles = chapter_titles[:content_slides_needed]
            
            expanded_titles = []
            if len(chapter_titles) > 0 and content_slides_needed > 0:
                base_count = content_slides_needed // len(chapter_titles)
                remainder = content_slides_needed % len(chapter_titles)
                
                for i, title in enumerate(chapter_titles):
                    # Distribute remainder to the first few chapters
                    count_for_this = base_count + (1 if i < remainder else 0)
                    if count_for_this == 1:
                        expanded_titles.append(title)
                    else:
                        for part in range(1, count_for_this + 1):
                            expanded_titles.append(f"{title} ({part}-qism)")
            
            forced_titles = expanded_titles
            
        # Resolve template path first to check for maket mode
        template_path = PRES_STYLES.get(style, {}).get("file")
        if not template_path:
            template_path = get_template_path(style)

        # Analyze template so AI can tailor content
        tpl_info = analyze_template(template_path) if template_path else {}
        tpl_context = ""
        is_maket_mode = False

        if template_path and ("maket" in template_path.lower() or "modern" in template_path.lower()):
            is_maket_mode = True
            if tpl_info.get("total_slides", 0) > 0:
                # Force exact number of slides in Maket mode
                ai_slides_count = tpl_info["total_slides"] - 1
            else:
                ai_slides_count = num_slides - 1
        else:
            ai_slides_count = num_slides - 1

        if tpl_info.get("total_slides", 0) > 1:
            if is_maket_mode and "slides" in tpl_info:
                # Provide detailed slide-by-slide block info for Maket Mode
                limits = []
                for s in tpl_info["slides"][1:]:  # skip title slide (index 0)
                    idx = s["index"]
                    body_blocks = s.get("body_blocks", 1)
                    blocks = s.get("blocks", [])
                    total_words = s.get("estimated_words", 40)
                    
                    if body_blocks >= 2:
                        # Multi-column slide — tell AI to generate separate blocks
                        block_desc = ", ".join([f"{i+1}-blok: {b['words']} so'z" for i, b in enumerate(blocks) if b['type'] == 'body'])
                        limits.append(f"Slayd {idx}: {body_blocks} ta alohida blok (kolonka) → {block_desc}. "
                                      f"Har bir blok uchun ALOHIDA punkt yozing!")
                    else:
                        limits.append(f"Slayd {idx}: maximal {total_words} ta so'z")
                
                limits_str = "\n".join(limits)
                tpl_context = (
                    f"\n\nSHABLON STRUKTURASI:\n"
                    f"Shablon {tpl_info['total_slides']} ta tayyor slayddan iborat.\n"
                    f"MUHIM! Slaydlar turli xil dizaynga ega: ba'zilari 1 ta matn bloki, ba'zilari 2 ta kolonka.\n"
                    f"2 ta blokli (kolonnali) slaydlar uchun — points massivida 2 ta ALOHIDA element qaytaring.\n"
                    f"1 ta blokli slaydlar uchun — 1-3 ta element qaytaring.\n\n"
                    f"SLAYDLAR BO'YICHA LIMITLAR:\n"
                    f"{limits_str}\n\n"
                    f"DIQQAT: Agar belgilangan so'zdan oshib ketsa, matn slaydga sig'may qoladi!"
                )
            else:
                tpl_context = (
                    f"\n\nSHABLON MA'LUMOTI: Shablon {tpl_info['total_slides']} ta tayyor slayddan iborat. "
                    f"AI matn har bir slaydga to'g'ri sig'ishi kerak. "
                    f"Har bir punktni 40-60 so'z atrofida yozing."
                )

        # Define a progress callback — shows "Slayd N/total" style progress
        import asyncio as _asyncio
        _last_progress_text = [""]
        async def update_progress(completed, total, next_title=""):
            # Build a small progress bar
            filled = int((completed / total) * 10)
            bar = "🟩" * filled + "⬜" * (10 - filled)
            next_line = f"\n⏭ <b>Keyingi:</b> <i>{next_title}</i>" if next_title else "\n✅ <b>Barcha slaydlar tayyor, fayl yig'ilmoqda...</b>"
            text = (
                f"🤖 <b>AI slayd yozmoqda...</b>\n\n"
                f"📊 <b>{completed}/{total}</b> slayd bajarildi\n"
                f"{bar}{next_line}\n\n"
                f"📝 <b>Mavzu:</b> {topic}\n"
                f"<i>Sifatli natija uchun 4-5 daqiqa kutish tavsiya etiladi.</i>"
            )
            if text != _last_progress_text[0]:
                _last_progress_text[0] = text
                try:
                    await wait_msg.edit_text(text, parse_mode="HTML")
                    await _asyncio.sleep(1.0)
                except Exception as e:
                    print(f"DEBUG progress update error: {e}")

        # Check cancellation before expensive AI call
        if is_cancelled():
            await state.clear()
            return

        raw_content = await generate_presentation_content(
            topic=full_topic + tpl_context,
            language=language,
            num_slides=ai_slides_count,
            style=style,
            quality=quality,
            num_chapters=num_chapters,
            forced_titles=forced_titles,
            progress_callback=update_progress
        )

        # Parse JSON
        print(f"DEBUG: AI Raw Response:\n{raw_content}")
        try:
            # Clean possible markdown code blocks
            clean_json = raw_content.replace("```json", "").replace("```", "").strip()
            
            # Robust JSON repair for truncated responses
            import json
            def repair_json(json_str):
                try:
                    return json.loads(json_str)
                except json.JSONDecodeError:
                    # Try appending closing brackets/quotes
                    for fix in ['"}', '"]}', '"]}]}', ']}', '}']:
                        try:
                            return json.loads(json_str + fix)
                        except json.JSONDecodeError:
                            pass
                    raise

            try:
                raw_data = repair_json(clean_json)
            except Exception:
                # If still failing, maybe there are unescaped quotes inside. Try a basic cleanup
                import re
                clean_json_fixed = re.sub(r'(?<!\\)"(?![:,\]}])', '\\"', clean_json)
                raw_data = repair_json(clean_json_fixed)

            raw_slides = raw_data.get("slides", raw_data) if isinstance(raw_data, dict) else raw_data

            # Normalize: AI may use different key names (uz/en/ru)
            # Expected: [{"title": "...", "points": ["...", "..."]}]
            def normalize_slide(s):
                if not isinstance(s, dict):
                    return {"title": str(s), "points": []}
                
                # Detect title key
                title = (
                    s.get("title") or s.get("sarlavha") or s.get("name") or
                    s.get("заголовок") or s.get("slide_title") or
                    f"Slide {s.get('slayd', '')}"
                )
                
                # Detect body key
                body = s.get("points") or s.get("matn") or s.get("content") or s.get("text") or s.get("mazmun") or ""
                
                # Normalize body to list
                if isinstance(body, str):
                    # Split long text into meaningful chunks (every ~150 chars at sentence boundary)
                    points = [p.strip() for p in body.split(". ") if p.strip()]
                    # Merge short chunks
                    merged = []
                    current = ""
                    for p in points:
                        if len(current) < 200:
                            current = (current + ". " + p).strip(". ")
                        else:
                            merged.append(current + ".")
                            current = p
                    if current:
                        merged.append(current if current.endswith(".") else current + ".")
                    points = merged if merged else [body]
                elif not isinstance(body, list):
                    points = [str(body)]
                else:
                    points = body

                # Extract image keyword if available
                img_kw = s.get("image_keyword") or ""

                return {"title": str(title), "points": points, "image_keyword": str(img_kw)}

            slides_data = [normalize_slide(s) for s in raw_slides]
            print(f"DEBUG: Normalized {len(slides_data)} slides OK.")

        except Exception as e:
            print(f"JSON Parse error: {e}\nRaw content: {raw_content}")
            # Fallback: put everything into a single slide
            slides_data = [{"title": topic, "points": [raw_content[:500]]}]

        author = data.get("author", "Foydalanuvchi")

        # ─── AKADEMIK TEMPLATE BRANCH ───────────────────────────────────
        if template_path and "akademik" in template_path.lower():
            try:
                await wait_msg.edit_text("🎓 <b>Akademik taqdimot yaratilmoqda...</b>\n\n⏳ <i>AI ilmiy matn yozmoqda...</i>", parse_mode="HTML")
            except: pass

            # Progress callback for akademik
            async def akademik_progress(completed, total, next_title=""):
                bar_filled = int((completed / total) * 10)
                bar = "🟩" * bar_filled + "⬜" * (10 - bar_filled)
                next_line = f"\n⏭ <b>Keyingi:</b> <i>{next_title}</i>" if next_title else "\n✅ <b>Tayyor!</b>"
                text = (
                    f"🎓 <b>Akademik taqdimot yaratilmoqda...</b>\n\n"
                    f"📊 <b>{completed}/{total}</b> bosqich bajarildi\n"
                    f"{bar}{next_line}"
                )
                try: await wait_msg.edit_text(text, parse_mode="HTML")
                except: pass

            # Generate structured akademik content via AI
            tag_data = await generate_akademik_content(
                topic=topic,
                language=language,
                subject_name=topic,
                completed_by=author,
                training_session="Ma'ruza mashg'uloti",
                quality=quality,
                progress_callback=akademik_progress,
            )

            # Generate AI images for akademik template
            image_data = {}
            image_tags = ["QUESTION_1_CONTENT_IMAGE", "QUESTION_2_CONTENT_IMAGE", "QUESTION_3_CONTENT_IMAGE"]
            
            if ai_images_count > 0:
                try:
                    await wait_msg.edit_text(f"🎨 <b>AI rasmlar yaratilmoqda ({min(ai_images_count, 3)} ta)...</b>", parse_mode="HTML")
                except: pass

                gen_count = 0
                for img_tag in image_tags[:ai_images_count]:
                    if is_cancelled():
                        await state.clear()
                        return
                    # Determine which question this image is for
                    q_num = img_tag.split("_")[1]  # "1", "2", or "3"
                    q_title = tag_data.get(f"QUESTION_{q_num}", topic)
                    try:
                        img_resp = await client.images.generate(
                            model="dall-e-3",
                            prompt=(
                                f"Educational scientific illustration for an academic lecture about: '{q_title}'. "
                                f"The overall topic is: '{topic}'. "
                                "Style: clean, professional, academic infographic or diagram. "
                                "Realistic and informative visual. No text, no letters, no words, no watermarks. "
                                "High resolution, white or neutral background."
                            ),
                            size="1024x1024",
                            quality="standard",
                            n=1
                        )
                        import httpx
                        async with httpx.AsyncClient() as http:
                            img_bytes = (await http.get(img_resp.data[0].url)).content
                        image_data[img_tag] = img_bytes
                        gen_count += 1
                        try:
                            await wait_msg.edit_text(f"🎨 <b>AI rasmlar: {gen_count}/{min(ai_images_count, 3)}</b>", parse_mode="HTML")
                        except: pass
                    except Exception as e:
                        print(f"DALL-E error for {img_tag}: {e}")

            try: await wait_msg.edit_text("⏳ <b>Taqdimot fayli yig'ilmoqda...</b>", parse_mode="HTML")
            except: pass

            # Generate the PPTX using akademik service
            pptx_bytes = await asyncio.get_event_loop().run_in_executor(
                None, generate_akademik_pptx, template_path, tag_data, image_data
            )

            # Mark free trial or deduct balance
            if is_admin:
                status_note = "🛡️ <b>Admin uchun tekin</b>"
            else:
                await deduct_balance(db_user.id, price)
                status_note = f"💰 <b>Hisobdan yechildi:</b> {format_price(price)}"

            try: await wait_msg.delete()
            except: pass

            await message.answer_document(
                document=BufferedInputFile(pptx_bytes, filename=f"akademik_{topic[:30]}.pptx"),
                caption=(
                    f"✅ <b>Akademik taqdimot muvaffaqiyatli tayyorlandi!</b>\n\n"
                    f"📝 <b>Mavzu:</b> {topic}\n"
                    f"🎓 <b>Tur:</b> Akademik (24 slayd)\n"
                    f"💎 <b>Sifat darajasi:</b> {'💎 Premium' if quality == 'premium' else '✨ Standart'}\n"
                    f"👤 <b>Muallif:</b> {author}\n"
                    f"────────────────────\n"
                    f"{status_note}\n\n"
                    f"📥 <i>Faylni yuklab olishingiz mumkin.</i>"
                ),
                reply_markup=main_menu_kb(),
                parse_mode="HTML"
            )
            return
        # ─── END AKADEMIK BRANCH ────────────────────────────────────────

        slide_images = {}
        
        # 1. Download and smartly map user photos using GPT-4o Vision
        user_photos = data.get("user_photos", [])
        downloaded_photos = []
        for file_id in user_photos:
            try:
                file = await message.bot.get_file(file_id)
                buf = io.BytesIO()
                await message.bot.download_file(file.file_path, buf)
                downloaded_photos.append(buf.getvalue())
            except Exception as e:
                print(f"Failed to download user photo: {e}")

        if downloaded_photos:
            try: await wait_msg.edit_text("⏳ <b>Rasmlar tahlil qilinmoqda (AI Vision)...</b>", parse_mode="HTML")
            except: pass
            
            import base64
            
            vision_content = [
                {
                    "type": "text", 
                    "text": "You are a presentation assistant. Match each provided image to the most appropriate slide index (1 to N). "
                            "Return ONLY a JSON list of integers representing the slide index for each image in the order they were provided. "
                            "Do not assign multiple images to the same slide if possible. Example output: [3, 1, 5]"
                }
            ]
            
            slides_text = "\n".join([f"Slide {i+1}: {s.get('title')} - {str(s.get('points'))[:100]}" for i, s in enumerate(slides_data)])
            vision_content.append({"type": "text", "text": f"Slides Content:\n{slides_text}"})
            
            for img_bytes in downloaded_photos:
                b64 = base64.b64encode(img_bytes).decode('utf-8')
                vision_content.append({
                    "type": "image_url",
                    "image_url": {"url": f"data:image/jpeg;base64,{b64}", "detail": "low"}
                })
                
            try:
                vision_resp = await client.chat.completions.create(
                    model="gpt-4o",
                    messages=[{"role": "user", "content": vision_content}],
                    temperature=0.1
                )
                mapping_str = vision_resp.choices[0].message.content.replace("```json", "").replace("```", "").strip()
                mapping = json.loads(mapping_str)
                
                used_slides = set()
                for i, target_slide in enumerate(mapping):
                    if i < len(downloaded_photos):
                        if isinstance(target_slide, int) and 1 <= target_slide <= len(slides_data):
                            if target_slide not in used_slides:
                                slide_images[target_slide] = downloaded_photos[i]
                                used_slides.add(target_slide)
            except Exception as e:
                print(f"Vision mapping failed: {e}. Falling back to sequential mapping.")
                for i, img_bytes in enumerate(downloaded_photos):
                    if i + 1 <= len(slides_data):
                        slide_images[i + 1] = img_bytes

        # 2. Generate AI images (dynamic count, smart placement)
        if ai_images_count > 0 and slides_data:
            try:
                await wait_msg.edit_text(f"🎨 <b>AI rasmlar yaratilmoqda ({ai_images_count} ta)...</b>", parse_mode="HTML")
            except: pass
            
            # Filter out title(0), Reja, Xulosa, Kirish from eligible slides
            EXCLUDED_WORDS = {"reja", "xulosa", "kirish", "план", "заключение", "введение", "plan", "conclusion", "introduction"}
            eligible = []
            for idx, sd in enumerate(slides_data):
                t = sd.get("title", "").lower().strip()
                # Check if any excluded word appears in the title (substring match)
                is_excluded = any(word in t for word in EXCLUDED_WORDS)
                if not is_excluded:
                    eligible.append(idx)
            
            if not eligible:
                eligible = list(range(len(slides_data)))
            
            # Evenly space images across eligible slides
            step = max(1, len(eligible) // max(1, ai_images_count))
            img_targets = eligible[::step][:ai_images_count]
            
            generated_count = 0
            for img_idx in img_targets:
                # Check cancellation before each expensive image generation
                if is_cancelled():
                    await state.clear()
                    return
                slide_idx_for_image = img_idx + 1  # +1 because slide 0 is title
                if slide_idx_for_image in slide_images:
                    continue  # user already uploaded image for this slide
                try:
                    slide_title = slides_data[img_idx].get("title", topic)
                    img_resp = await client.images.generate(
                        model="dall-e-3",
                        prompt=(
                            f"Educational scientific illustration for an academic presentation slide about: '{slide_title}'. "
                            f"The overall presentation topic is: '{topic}'. "
                            "Style: clean, professional, academic infographic or diagram. "
                            "Realistic and informative visual. No text, no letters, no words, no watermarks. "
                            "High resolution, white or neutral background."
                        ),
                        size="1024x1024",
                        quality="standard",
                        n=1
                    )
                    import httpx
                    async with httpx.AsyncClient() as http:
                        img_bytes = (await http.get(img_resp.data[0].url)).content
                    slide_images[slide_idx_for_image] = img_bytes
                    generated_count += 1
                    try:
                        await wait_msg.edit_text(f"🎨 <b>AI rasmlar: {generated_count}/{ai_images_count}</b>", parse_mode="HTML")
                    except: pass
                except Exception as e:
                    print(f"DALL-E image gen error for slide {img_idx}: {e}")

        try: await wait_msg.edit_text("⏳ <b>Taqdimot fayli yig'ilmoqda...</b>", parse_mode="HTML")
        except: pass

        pptx_bytes = await asyncio.get_event_loop().run_in_executor(
            None, 
            lambda: generate_pptx(slides_data, template_path, topic, author, slide_images, is_maket_mode)
        )

        # Mark free trial or deduct balance
        if is_admin:
            status_note = "🛡️ <b>Admin uchun tekin</b>"
        else:
            await deduct_balance(db_user.id, price)
            status_note = f"💰 <b>Hisobdan yechildi:</b> {format_price(price)}"

        quality_label = "💎 Premium" if quality == "premium" else "✨ Standart"
        
        try:
            await wait_msg.delete()
        except:
            pass
            
        await message.answer_document(
            document=BufferedInputFile(pptx_bytes, filename=f"taqdimot_{topic[:30]}.pptx"),
            caption=(
                f"✅ <b>Taqdimot muvaffaqiyatli tayyorlandi!</b>\n\n"
                f"📝 <b>Mavzu:</b> {topic}\n"
                f"📊 <b>Slaydlar soni:</b> {num_slides} ta\n"
                f"💎 <b>Sifat darajasi:</b> {quality_label}\n"
                f"👤 <b>Muallif:</b> {db_user.full_name or 'Foydalanuvchi'}\n"
                f"────────────────────\n"
                f"{status_note}\n\n"
                f"📥 <i>Faylni yuklab olishingiz mumkin.</i>"
            ),
            reply_markup=main_menu_kb(),
            parse_mode="HTML"
        )
    except Exception as e:
        import traceback
        traceback.print_exc()
        print(f"Presentation error: {e}")
        await message.answer(f"❌ Xatolik: {str(e)}\n\nIltimos, qaytadan urinib ko'ring.", reply_markup=main_menu_kb())
