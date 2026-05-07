import asyncio, json
from aiogram import Router, F
from aiogram.types import Message, CallbackQuery, BufferedInputFile
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.types import InlineKeyboardMarkup, InlineKeyboardButton

from database.models import User
from database.db import create_request, deduct_balance
from keyboards.main_kb import main_menu_kb, back_to_menu_kb
from services.ai_service import generate_document_plan, generate_document_section
from services.docx_service import generate_docx
from services.template_loader import load_template_and_example
from utils.helpers import format_price
from config import ADMIN_IDS

router = Router()

class CourseworkStates(StatesGroup):
    entering_topic = State()
    choosing_ministry = State()
    entering_student_info = State()
    entering_department = State()
    choosing_plan_method = State()
    choosing_plan_structure = State()
    entering_manual_plan = State()
    choosing_pages = State()
    reviewing_summary = State()

# ─── Keyboards ──────────────────────────────────────────────────────────────

def ministry_kb():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="1. Oliy ta'lim, fan va innovatsiyalar vazirligi", callback_data="min:1")],
        [InlineKeyboardButton(text="2. Maktabgacha va maktab ta'limi vazirligi", callback_data="min:2")],
        [InlineKeyboardButton(text="⬅️ Bekor qilish", callback_data="main_menu")]
    ])

def plan_method_kb():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="✍️ O'zim yozaman", callback_data="plan_method:manual")],
        [InlineKeyboardButton(text="🤖 Avtomatik AI yaratishi", callback_data="plan_method:ai")]
    ])

def plan_structure_kb():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📊 2 ta bob 4 ta rejalik", callback_data="struct:2:4")],
        [InlineKeyboardButton(text="📊 2 ta bob 6 ta rejalik", callback_data="struct:2:6")],
        [InlineKeyboardButton(text="📊 3 ta bob 6 ta rejalik", callback_data="struct:3:6")],
        [InlineKeyboardButton(text="📊 3 ta bob 9 ta rejalik", callback_data="struct:3:9")],
        [InlineKeyboardButton(text="⬅️ Orqaga", callback_data="plan_method:ai")]
    ])

def coursework_pages_kb():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📄 20-25 bet - 13,000 so'm", callback_data="cpages:25:13000")],
        [InlineKeyboardButton(text="📄 25-30 bet - 16,000 so'm", callback_data="cpages:30:16000")],
        [InlineKeyboardButton(text="📄 30-35 bet - 19,000 so'm", callback_data="cpages:35:19000")],
        [InlineKeyboardButton(text="📄 35-40 bet - 22,000 so'm", callback_data="cpages:40:22000")],
        [InlineKeyboardButton(text="📄 40-50 bet - 25,000 so'm", callback_data="cpages:50:25000")],
        [InlineKeyboardButton(text="⬅️ Orqaga", callback_data="main_menu")]
    ])

def confirm_coursework_kb():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🚀 Yaratish", callback_data="coursework_start")],
        [InlineKeyboardButton(text="❌ Bekor qilish", callback_data="main_menu")]
    ])

# ─── Handlers ────────────────────────────────────────────────────────────────

@router.message(F.text == "📘 Kurs ishi yaratish")
async def start_coursework(message: Message, state: FSMContext, db_user: User):
    from config import PRICING
    price_low = PRICING["coursework_low"]
    price_high = PRICING["coursework_high"]
    is_admin = db_user.id in ADMIN_IDS
    balance = db_user.balance or 0
    can_afford = is_admin or balance >= price_low
    balance_display = "🛡️ Admin (tekin)" if is_admin else format_price(balance)

    if can_afford:
        kb = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="✅ Davom etish", callback_data="cw_confirm")],
            [InlineKeyboardButton(text="❌ Bekor qilish", callback_data="main_menu")]
        ])
        status_line = "✅ <b>Balansingiz yetarli!</b>"
    else:
        needed = price_low - balance
        kb = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="💳 Balans to'ldirish", callback_data="cw_topup")],
            [InlineKeyboardButton(text="❌ Bekor qilish", callback_data="main_menu")]
        ])
        status_line = f"⚠️ <b>Balansingiz yetarli emas!</b>\n🔴 Yetishmayapti: <b>{format_price(needed)}</b>"

    await message.answer(
        "📘 <b>Kurs Ishi</b>\n"
        "<i>To'liq akademik kurs ishi — mundarija, kirish, boblar, xulosa, adabiyotlar</i>\n\n"
        "━━━━━━━━━━━━━━━━\n"
        f"💰 <b>Narx:</b> {format_price(price_low)} — {format_price(price_high)}\n"
        f"   <i>(sahifalar soniga qarab)</i>\n"
        f"👛 <b>Sizning balansingiz:</b> {balance_display}\n"
        "━━━━━━━━━━━━━━━━\n"
        f"{status_line}",
        reply_markup=kb,
        parse_mode="HTML"
    )

@router.callback_query(F.data == "cw_confirm")
async def confirm_coursework_start(callback: CallbackQuery, state: FSMContext):
    await state.set_state(CourseworkStates.entering_topic)
    await callback.message.edit_text(
        "📘 <b>Yangi kurs ishi yaratamiz!</b>\n\n"
        "Boshlash uchun <b>mavzuni</b> yuboring.\n"
        "<i>(Masalan: Raqamli iqtisodiyotda blokcheyn texnologiyalari)</i>",
        parse_mode="HTML"
    )
    await callback.answer()

@router.callback_query(F.data == "cw_topup")
async def topup_from_coursework(callback: CallbackQuery, state: FSMContext):
    await state.clear()
    await callback.message.edit_text(
        "💳 Balansni to'ldirish uchun /buy buyrug'ini yuboring.",
        parse_mode="HTML"
    )
    await callback.answer()


@router.message(CourseworkStates.entering_topic)
async def enter_topic(message: Message, state: FSMContext):
    await state.update_data(topic=message.text.strip())
    await state.set_state(CourseworkStates.choosing_ministry)
    await message.answer(
        "🏛 <b>Vazirlikni tanlang:</b>",
        reply_markup=ministry_kb(),
        parse_mode="HTML"
    )

@router.callback_query(F.data.startswith("min:"), CourseworkStates.choosing_ministry)
async def choose_ministry(callback: CallbackQuery, state: FSMContext):
    min_idx = callback.data.split(":")[1]
    min_text = "O'zbekiston Respublikasi Oliy ta'lim, fan va innovatsiyalar vazirligi" if min_idx == "1" else "O'zbekiston Respublikasi Maktabgacha va maktab ta'limi vazirligi"
    await state.update_data(ministry=min_text)
    await state.set_state(CourseworkStates.entering_student_info)
    await callback.message.edit_text(
        "👤 <b>Talaba ma'lumotlarini kiriting:</b>\n\n"
        "To'liq ismingiz, kursingiz va guruhingizni bitta xabarda yozing.\n"
        "<i>Misol: Alisherov Anvar 3-kurs 201-guruh</i>",
        parse_mode="HTML"
    )

@router.message(CourseworkStates.entering_student_info)
async def enter_student_info(message: Message, state: FSMContext):
    await state.update_data(student_info=message.text.strip())
    await state.set_state(CourseworkStates.entering_department)
    await message.answer(
        "🏢 <b>Kafedra nomini kiriting:</b>\n"
        "<i>Namuna: IQTISOD VA BIZNES KAFEDRASI</i>",
        parse_mode="HTML"
    )

@router.message(CourseworkStates.entering_department)
async def enter_department(message: Message, state: FSMContext):
    await state.update_data(department=message.text.strip())
    await state.set_state(CourseworkStates.choosing_plan_method)
    await message.answer(
        "📝 <b>Mundarijani o'zingiz yozasizmi yoki AI avtomatik yaratishini xohlaysizmi?</b>",
        reply_markup=plan_method_kb(),
        parse_mode="HTML"
    )

@router.callback_query(F.data.startswith("plan_method:"), CourseworkStates.choosing_plan_method)
async def choose_plan_method(callback: CallbackQuery, state: FSMContext):
    method = callback.data.split(":")[1]
    if method == "ai":
        await state.set_state(CourseworkStates.choosing_plan_structure)
        await callback.message.edit_text(
            "📊 <b>Mundarija tuzilmasini tanlang:</b>",
            reply_markup=plan_structure_kb(),
            parse_mode="HTML"
        )
    else:
        await state.set_state(CourseworkStates.entering_manual_plan)
        await callback.message.edit_text(
            "✍️ <b>Mundarijani yuboring:</b>\n"
            "Har bir bo'limni yangi qatordan yozing.",
            parse_mode="HTML"
        )

@router.callback_query(F.data.startswith("struct:"), CourseworkStates.choosing_plan_structure)
async def choose_plan_structure(callback: CallbackQuery, state: FSMContext):
    _, bobs, rejalar = callback.data.split(":")
    await state.update_data(plan_structure=f"{bobs} ta bob {rejalar} ta reja")
    await state.set_state(CourseworkStates.choosing_pages)
    await callback.message.edit_text(
        "📑 <b>Kurs ishi hajmini tanlang:</b>",
        reply_markup=coursework_pages_kb(),
        parse_mode="HTML"
    )

@router.callback_query(F.data.startswith("cpages:"), CourseworkStates.choosing_pages)
async def choose_pages(callback: CallbackQuery, state: FSMContext, db_user: User):
    _, pages, price = callback.data.split(":")
    await state.update_data(num_pages=int(pages), price=int(price))
    
    await state.set_state(CourseworkStates.reviewing_summary)
    data = await state.get_data()
    
    summary = (
        f"📘 <b>Kurs Ishi Yakuni</b>\n\n"
        f"📌 <b>Mavzu:</b> {data.get('topic')}\n"
        f"👤 <b>Muallif:</b> {data.get('student_info')}\n"
        f"🏢 <b>Kafedra:</b> {data.get('department')}\n"
        f"📑 <b>Hajmi:</b> {pages} bet\n"
        f"💰 <b>Narxi:</b> {format_price(price)}\n"
        f"────────────────────\n"
        f"Tasdiqlaysizmi?"
    )
    await callback.message.edit_text(summary, reply_markup=confirm_coursework_kb(), parse_mode="HTML")

@router.callback_query(F.data == "coursework_start", CourseworkStates.reviewing_summary)
async def generate_coursework(callback: CallbackQuery, state: FSMContext, db_user: User):
    data = await state.get_data()
    topic = data.get("topic")
    price = data.get("price")
    pages = data.get("num_pages")
    
    if (db_user.balance or 0) < price and db_user.id not in ADMIN_IDS:
        await callback.answer("⚠️ Balansingiz yetarli emas!", show_alert=True)
        return

    # Cancellation support
    cancel_flag = {"cancelled": False}
    await state.update_data(cancel_flag=cancel_flag)
    
    cancel_kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="❌ Bekor qilish", callback_data="cw_cancel_gen")]
    ])

    wait_msg = await callback.message.edit_text(
        f"⏳ <b>Kurs ishi tayyorlanmoqda...</b>\n\n"
        f"🤖 AI mavzuni chuqur tahlil qilmoqda.\n"
        f"Bu jarayon 2-4 daqiqa vaqt olishi mumkin.",
        parse_mode="HTML",
        reply_markup=cancel_kb
    )

    def is_cancelled():
        return cancel_flag.get("cancelled", False)

    try:
        plan = data.get("manual_plan")
        if not plan:
            struct = data.get("plan_structure", "2 ta bob 6 ta reja")
            plan_data = await generate_document_plan("coursework", f"{topic} ({struct} asosida)", data.get("language", "uz"), "pro")
            plan = plan_data.get("plan", "")

        sections = [s.strip() for s in plan.split('\n') if s.strip()]
        total_target_words = pages * 270 
        words_per_section = int(total_target_words / len(sections))
        
        # Load structure template + quality example
        template_context = load_template_and_example("coursework")
        
        # Image placeholder instruction for AI
        image_placeholder_rule = (
            "\n\nRASM QO'YISH QOIDASI (MUHIM!):\n"
            "- Matn ichida rasm, jadval, grafik yoki diagramma kerak bo'lgan joyda "
            "hech qachon ularni matn bilan tasvirlama!\n"
            "- O'rniga faqat quyidagi pleysholderni yoz:\n"
            "  [ 🖼️ SHU YERGA RASM JOYLANG: <rasm tavsifi> ]\n"
            "- Masalan:\n"
            "  [ 🖼️ SHU YERGA RASM JOYLANG: YaIM o'sish grafigi 2015-2024 ]\n"
            "  [ 🖼️ SHU YERGA RASM JOYLANG: Tashkiliy tuzilma sxemasi ]\n"
            "- Har bir bobda kamida 1-2 ta shu turdagi pleyssholder bo'lishi KERAK!\n"
            "- Pleysholderdan keyin matn odatdagidek davom etsin.\n"
        )
        
        full_content = []
        for i, sec in enumerate(sections, 1):
            # Check cancellation before each section
            if is_cancelled():
                await state.clear()
                return
            
            filled = int(((i - 1) / len(sections)) * 10)
            bar = "🟩" * filled + "⬜" * (10 - filled)
            try:
                await wait_msg.edit_text(
                    f"🤖 <b>AI kurs ishi yozmoqda...</b>\n\n"
                    f"📊 <b>{i-1}/{len(sections)}</b> bo'lim tayyor\n"
                    f"{bar}\n\n"
                    f"⏭ <b>Hozir yozilmoqda:</b> <i>{sec}</i>\n\n"
                    f"<i>Sifatli natija uchun biroz kuting...</i>",
                    parse_mode="HTML",
                    reply_markup=cancel_kb
                )
            except:
                pass
            
            text = await generate_document_section(
                topic=topic, 
                section_title=sec, 
                extra_details=(
                    f"{template_context}\n" if template_context else ""
                ) + image_placeholder_rule
                  + f"\nKurs ishi uchun juda batafsil, ilmiy matn. Kamida {words_per_section} ta so'z.", 
                language=data.get("language", "uz"),
                quality="pro",
                service_type="coursework"
            )
            full_content.append(f"# {sec}\n\n{text}")

        if is_cancelled():
            await state.clear()
            return

        bib_title = "FOYDALANILGAN ADABIYOTLAR RO'YXATI"
        bib_text = await generate_document_section(topic, bib_title, "Kamida 15-20 ta ilmiy manba.", data.get("language", "uz"), "standard", "coursework")
        full_content.append(f"# {bib_title}\n\n{bib_text}")

        content = "\n\n".join(full_content)
        
        docx_bytes = await asyncio.get_event_loop().run_in_executor(
            None, 
            generate_docx, 
            "coursework", 
            topic, 
            content, 
            data.get("student_info", "Talaba"), 
            plan,
            {"ministry": data.get("ministry"), "department": data.get("department")}
        )

        if db_user.id not in ADMIN_IDS:
            await deduct_balance(db_user.id, price)
        
        await create_request(db_user.id, "coursework", topic, {"pages": pages})
        
        await wait_msg.delete()
        file_name = f"Kurs_ishi_{topic[:20]}.docx"
        await callback.message.answer_document(
            document=BufferedInputFile(docx_bytes, filename=file_name),
            caption=f"✅ <b>Kurs ishi tayyor!</b>\n\n📌 Mavzu: {topic}\n📑 Hajmi: {pages} bet\n💰 Narxi: {format_price(price)}",
            parse_mode="HTML"
        )
        await state.clear()
    except Exception as e:
        await callback.message.answer(f"❌ Xatolik: {e}")

@router.callback_query(F.data == "cw_cancel_gen")
async def cancel_coursework_gen(callback: CallbackQuery, state: FSMContext):
    data = await state.get_data()
    cancel_flag = data.get("cancel_flag")
    if cancel_flag and isinstance(cancel_flag, dict):
        cancel_flag["cancelled"] = True
    await state.clear()
    await callback.message.edit_text("❌ Kurs ishi yaratish to'xtatildi. Hech narsa hisobdan yechilmadi.")
    await callback.answer()

@router.message(CourseworkStates.entering_manual_plan)
async def enter_manual_plan(message: Message, state: FSMContext):
    await state.update_data(manual_plan=message.text.strip())
    await state.set_state(CourseworkStates.choosing_pages)
    await message.answer("📑 <b>Kurs ishi hajmini tanlang:</b>", reply_markup=coursework_pages_kb())
