import asyncio, json, os, io
from aiogram import Router, F
from aiogram.types import Message, CallbackQuery, BufferedInputFile
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup

from database.models import User
from database.db import create_request, deduct_balance
from keyboards.main_kb import main_menu_kb, back_to_menu_kb
from keyboards.documents_kb import page_count_kb, referat_summary_kb
from services.ai_service import generate_document_plan, generate_document_section
from services.docx_service import generate_docx, generate_docx_from_template
from services.template_loader import load_template_and_example
from utils.helpers import format_price
from config import ADMIN_IDS

router = Router()

class DocumentStates(StatesGroup):
    entering_topic    = State()
    entering_subject  = State()
    choosing_quality  = State()
    choosing_pages    = State()
    reviewing_summary = State()
    waiting_for_file      = State()   # Mustaqil ish: fayl kutilmoqda
    waiting_file_language = State()   # Mustaqil ish: fayl o'qildi, til kutilmoqda

# ─── Step 1: Start ──────────────────────────────────────────────────────────

from aiogram.types import InlineKeyboardMarkup, InlineKeyboardButton
from config import PRICING, ADMIN_IDS

@router.message(F.text.in_(["📚 Referat Yaratish", "📄 Mustaqil Ish Yaratish"]))
async def start_referat_creation(message: Message, state: FSMContext, db_user: User):
    is_referat = "Referat" in message.text
    service_label = "📚 <b>Referat</b>" if is_referat else "📄 <b>Mustaqil Ish</b>"
    price_std = PRICING["essay_std"] if is_referat else PRICING["mustaqil_std"]
    price_pro = PRICING["essay_pre"] if is_referat else PRICING["mustaqil_pre"]
    service_type = "referat" if is_referat else "mustaqil"

    is_admin = db_user.id in ADMIN_IDS
    balance = db_user.balance or 0
    can_afford = is_admin or balance >= price_std
    balance_display = "🛡️ Admin (tekin)" if is_admin else format_price(balance)

    is_mustaqil = not is_referat

    if can_afford:
        if is_mustaqil:
            kb = InlineKeyboardMarkup(inline_keyboard=[
                [InlineKeyboardButton(text="✅ Mavzu yozib yaratish", callback_data=f"doc_confirm:{service_type}")],
                [InlineKeyboardButton(text="📎 Fayldan yaratish", callback_data=f"doc_file:{service_type}")],
                [InlineKeyboardButton(text="❌ Bekor qilish", callback_data="doc_cancel")]
            ])
        else:
            kb = InlineKeyboardMarkup(inline_keyboard=[
                [InlineKeyboardButton(text="✅ Davom etish", callback_data=f"doc_confirm:{service_type}")],
                [InlineKeyboardButton(text="❌ Bekor qilish", callback_data="doc_cancel")]
            ])
        status_line = "✅ <b>Balansingiz yetarli!</b>"
    else:
        needed = price_std - balance
        kb = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="💳 Balans to'ldirish", callback_data="doc_topup")],
            [InlineKeyboardButton(text="❌ Bekor qilish", callback_data="doc_cancel")]
        ])
        status_line = f"⚠️ <b>Balansingiz yetarli emas!</b>\n🔴 Yetishmayapti: <b>{format_price(price_std - balance)}</b>"

    await message.answer(
        f"{service_label}\n\n"
        f"━━━━━━━━━━━━━━━━\n"
        f"💰 <b>Narx:</b> {format_price(price_std)} — {format_price(price_pro)}\n"
        f"   <i>(Standart va Pro sifatiga qarab)</i>\n"
        f"👛 <b>Sizning balansingiz:</b> {balance_display}\n"
        f"━━━━━━━━━━━━━━━━\n"
        f"{status_line}",
        reply_markup=kb,
        parse_mode="HTML"
    )

@router.callback_query(F.data.startswith("doc_confirm:"))
async def confirm_document(callback: CallbackQuery, state: FSMContext):
    await state.set_state(DocumentStates.entering_topic)
    await callback.message.edit_text(
        "📝 <b>Mavzuni kiriting:</b>\n"
        "<i>(Masalan: O'zbekiston iqtisodiyoti rivoji)</i>",
        parse_mode="HTML"
    )
    await callback.answer()

@router.callback_query(F.data == "doc_cancel")
async def cancel_document(callback: CallbackQuery, state: FSMContext):
    await state.clear()
    await callback.message.edit_text("❌ Amaliyot bekor qilindi.")
    await callback.answer()

@router.callback_query(F.data == "doc_topup")
async def topup_from_document(callback: CallbackQuery, state: FSMContext):
    await state.clear()
    await callback.message.edit_text(
        "💳 Balansni to'ldirish uchun /buy buyrug'ini yuboring.",
        parse_mode="HTML"
    )
    await callback.answer()


# ─── File upload flow (Mustaqil ish only) ───────────────────────────────────

async def _read_file(message: Message) -> str:
    """Extract text from .txt / .docx / .pdf document."""
    doc  = message.document
    file = await message.bot.get_file(doc.file_id)
    buf  = io.BytesIO()
    await message.bot.download_file(file.file_path, buf)
    buf.seek(0)
    name = (doc.file_name or "").lower()
    if name.endswith(".txt"):
        return buf.read().decode("utf-8", errors="ignore")
    elif name.endswith(".docx"):
        from docx import Document as DocxDoc
        return "\n".join(p.text for p in DocxDoc(buf).paragraphs if p.text.strip())
    elif name.endswith(".pdf"):
        import pdfplumber
        parts = []
        with pdfplumber.open(buf) as pdf:
            for page in pdf.pages:
                t = page.extract_text()
                if t:
                    parts.append(t)
        return "\n".join(parts)
    else:
        raise ValueError("Qo'llab-quvvatlanmaydigan format. .txt, .docx yoki .pdf yuboring.")


@router.callback_query(F.data.startswith("doc_file:"))
async def ask_for_file(callback: CallbackQuery, state: FSMContext):
    service_type = callback.data.split(":")[1]   # 'mustaqil'
    await state.update_data(service_type=service_type, is_referat=False)
    await state.set_state(DocumentStates.waiting_for_file)
    await callback.message.edit_text(
        "📎 <b>Faylingizni yuboring</b>\n\n"
        "Qo'llab-quvvatlanadigan formatlar:\n"
        "• <b>.txt</b> — oddiy matn\n"
        "• <b>.docx</b> — Word hujjati\n"
        "• <b>.pdf</b> — PDF hujjati\n\n"
        "<i>AI faylni o'qib, unda yozilgan topshiriqni bajaradi va "
        "mustaqil ish sifatida formatlaydi.</i>",
        parse_mode="HTML"
    )
    await callback.answer()


@router.message(F.document, DocumentStates.waiting_for_file)
async def receive_file_for_mustaqil(message: Message, state: FSMContext, db_user: User):
    proc = await message.answer("⏳ <b>Fayl o'qilmoqda...</b>", parse_mode="HTML")
    try:
        source_text = await _read_file(message)
        if not source_text.strip():
            await proc.edit_text("❌ Fayl bo'sh yoki o'qib bo'lmadi.")
            return
        source_text = source_text[:15000]

        # Detect complexity marker
        lower = source_text.lower()
        is_murakkab = any(kw in lower for kw in [
            "murakkab", "мураккаб", "сложный", "complex", "qiyin topshiriq"
        ])
        price    = 7000 if is_murakkab else 5000
        mode     = "murakkab" if is_murakkab else "normal"
        balance  = db_user.balance or 0
        is_admin = db_user.id in ADMIN_IDS

        if not is_admin and balance < price:
            await proc.edit_text(
                f"⚠️ <b>{'Murakkab topshiriq' if is_murakkab else 'Mustaqil ish'} aniqlandi!</b>\n"
                f"💰 Narxi: <b>{format_price(price)}</b>\n"
                f"💳 Balansingiz: <b>{format_price(balance)}</b>\n\n"
                "❌ Balans yetarli emas.",
                parse_mode="HTML"
            )
            await state.clear()
            return

        await state.update_data(
            source_text=source_text,
            quality="pro",
            mode=mode,
            price=price,
            num_pages=10,        # default, AI decides actual length from file
            topic="avtomatik",   # AI extracts from file
            author=db_user.full_name or "Foydalanuvchi",
        )
        await state.set_state(DocumentStates.waiting_file_language)
        await proc.delete()

        badge = "🔴 <b>Murakkab topshiriq!</b>" if is_murakkab else "✅ <b>Fayl o'qildi!</b>"
        await message.answer(
            f"{badge}\n"
            f"📝 Hajm: ~{len(source_text.split())} so'z\n"
            f"💰 Narxi: <b>{format_price(price)}</b>\n\n"
            "Hujjat tilini tanlang:",
            reply_markup=InlineKeyboardMarkup(inline_keyboard=[
                [InlineKeyboardButton(text="🇺🇿 O'zbekcha",    callback_data="file_lang:uz")],
                [InlineKeyboardButton(text="🇷🇺 Русский",      callback_data="file_lang:ru")],
                [InlineKeyboardButton(text="🇬🇧 English",       callback_data="file_lang:en")],
                [InlineKeyboardButton(text="❌ Bekor qilish", callback_data="file_lang:cancel")],
            ]),
            parse_mode="HTML"
        )
    except ValueError as e:
        await proc.edit_text(f"❌ {e}")
    except Exception as e:
        await proc.edit_text(f"❌ Xatolik: {e}")


@router.callback_query(F.data.startswith("file_lang:"), DocumentStates.waiting_file_language)
async def file_language_chosen(callback: CallbackQuery, state: FSMContext, db_user: User):
    lang = callback.data.split(":")[1]
    if lang == "cancel":
        await state.clear()
        await callback.message.edit_text("❌ Bekor qilindi.")
        await callback.answer()
        return

    await state.update_data(language=lang)
    data       = await state.get_data()
    source_text = data["source_text"]
    mode        = data["mode"]
    price       = data["price"]
    is_admin    = db_user.id in ADMIN_IDS

    lang_label = {"uz": "O'zbekcha 🇺🇿", "ru": "Rus tili 🇷🇺", "en": "English 🇬🇧"}.get(lang, lang)
    lang_map   = {"uz": "O'zbek tilida (Lotin alifbosi)", "ru": "На русском языке", "en": "In English"}
    lang_instruction = lang_map.get(lang, "O'zbek tilida")

    await callback.message.edit_text(
        f"{'🔴 Murakkab' if mode == 'murakkab' else '🧠 AI'} rejim | {lang_label}\n"
        f"🔄 Tahlil boshlandi...",
        parse_mode="HTML"
    )
    await callback.answer()

    wait_msg = await callback.message.bot.send_message(
        chat_id=callback.message.chat.id,
        text="🧠 <b>AI faylni tahlil qilmoqda...</b>\n"
             "📋 Bo'limlar ro'yxati aniqlanmoqda...",
        parse_mode="HTML"
    )

    try:
        from services.ai_service import client, OPENAI_MODEL

        # Retry wrapper for OpenAI calls
        async def call_openai_with_retry(messages, temp, max_tok, retries=4):
            import openai
            for attempt in range(retries):
                try:
                    resp = await client.chat.completions.create(
                        model=OPENAI_MODEL,
                        messages=messages,
                        temperature=temp,
                        max_tokens=max_tok,
                        timeout=180.0  # 3 minutes per call
                    )
                    return resp.choices[0].message.content.strip()
                except Exception as e:
                    if attempt == retries - 1:
                        raise e
                    await asyncio.sleep(2 ** attempt)

        # ── STEP 1: Extract sections from file ────────────────────────────────
        sections_raw = await call_openai_with_retry(
            messages=[
                {"role": "system", "content":
                    "You are an academic document analyzer. "
                    "Read the file content and extract the exact list of sections that need to be written based ONLY on the user's instructions. "
                    "Output ONLY section names, one per line. No numbers, no extra text. If no explicit structure is given, derive a logical detailed structure for a comprehensive assignment."
                },
                {"role": "user", "content":
                    f"FILE CONTENT:\n{source_text[:8000]}\n\n"
                    f"Task: Extract or design the section plan. Language for section names: {lang_instruction}."}
            ],
            temp=0.3,
            max_tok=800
        )
        sections = [s.strip().lstrip("0123456789.-) ") for s in sections_raw.split("\n") if s.strip()]
        if not sections:
            sections = ["Kirish", "Asosiy qism", "Xulosa"]

        total_sec = len(sections)

        await wait_msg.edit_text(
            f"🧠 <b>AI tahlil qildi!</b>\n\n"
            f"📋 <b>{total_sec} ta bo'lim aniqlandi</b>\n"
            + "\n".join(f"  {i+1}. {s}" for i, s in enumerate(sections))
            + "\n\n<i>Yozish boshlandi...</i>",
            parse_mode="HTML"
        )
        await asyncio.sleep(1.5)

        # ── STEP 2: Section-by-section generation ────────────────────────────
        murakkab_rule = (
            "CRITICAL MODE: Follow file instructions EXACTLY and completely. "
            "No templates, no extra rules, no simplifications. "
            "Do exactly and fully what the file says.\n\n"
        ) if mode == "murakkab" else ""

        system_prompt = (
            "You are a world-class academic expert executing a complex assignment. "
            "You MUST fulfill the user's file instructions with absolute precision and completely unrestricted depth. "
            "If the file asks for 4000, 6000 words, or extensive details, YOU MUST DO IT. There are no length limits. "
            "Every section must be comprehensive, analytical, and highly detailed.\n\n"
            f"{murakkab_rule}"
            f"DOCUMENT LANGUAGE: {lang_instruction}\n\n"
            f"MANDATORY RULES:\n"
            "- Never shorten or summarize unless requested.\n"
            "- Do NOT use markdown symbols (**, #, ```).\n"
            "- Write ONLY in full academic paragraphs.\n"
            "- Do NOT write section titles in the body — they are headings.\n\n"
            f"FULL FILE CONTENT (your primary reference for ALL sections):\n"
            f"{'='*60}\n{source_text[:10000]}\n{'='*60}"
        )

        full_parts = []

        for i, section_title in enumerate(sections, 1):
            done = "🟩" * (i - 1)
            cur  = "🟨"
            left = "⬜" * (total_sec - i)
            bar  = done + cur + left

            # Phase 1: Planning
            try:
                await wait_msg.edit_text(
                    f"🧠 <b>AI chuqur tahlil qilmoqda...</b>\n\n"
                    f"📄 <b>Bo'lim {i}/{total_sec}</b>\n"
                    f"{bar}\n\n"
                    f"🔍 <b>1-bosqich — Reja:</b> <i>{section_title}</i>\n\n"
                    f"<i>AI avval fikrlaydi, keyin yozadi — sifat uchun sabr qiling...</i>",
                    parse_mode="HTML"
                )
            except:
                pass

            planning_prompt = (
                f"You are preparing to write the '{section_title}' section.\n\n"
                "Perform a DEEP pre-writing analysis based on the file content:\n"
                "1. What are the 5 most critical points for this section?\n"
                "2. What specific facts, statistics, or examples from the file apply here?\n"
                "3. What additional academic knowledge supports this section?\n"
                "4. What is the most logical paragraph order for maximum impact?\n"
                "5. What counterarguments exist and how to address them?\n\n"
                "Write your analysis plan (250–400 words)."
            )
            try:
                section_plan = await call_openai_with_retry(
                    messages=[
                        {"role": "system", "content": system_prompt},
                        {"role": "user",   "content": planning_prompt}
                    ],
                    temp=0.5,
                    max_tok=2000
                )
            except:
                section_plan = ""

            # Phase 2: Writing
            try:
                await wait_msg.edit_text(
                    f"✍️ <b>AI yozmoqda...</b>\n\n"
                    f"📄 <b>Bo'lim {i}/{total_sec}</b>\n"
                    f"{bar}\n\n"
                    f"📝 <b>2-bosqich — Yozish:</b> <i>{section_title}</i>\n\n"
                    f"<i>Reja tayyor — to'liq matn yozilmoqda...</i>",
                    parse_mode="HTML"
                )
            except:
                pass
            await asyncio.sleep(0.5)

            write_prompt = (
                f"Write the '{section_title}' section of the academic document.\n\n"
                f"MANDATORY:\n"
                f"- Obey ALL instructions from the file content.\n"
                f"- Write in {lang_instruction}.\n"
                f"- DO NOT artificially limit the length. If the assignment requires huge detail, provide it.\n"
                f"- Do NOT write the section title — only the content.\n"
                f"- No markdown. Only full academic paragraphs.\n"
                + (f"\nYOUR PRE-WRITING ANALYSIS (use as blueprint):\n{section_plan}" if section_plan else "")
            )

            section_text = await call_openai_with_retry(
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user",   "content": write_prompt}
                ],
                temp=0.8,
                max_tok=15000
            )
            full_parts.append(f"# {section_title}\n\n{section_text}")

        # Done
        try:
            await wait_msg.edit_text(
                f"✅ <b>Barcha {total_sec} ta bo'lim tayyor!</b>\n\n"
                f"{'🟩' * total_sec}\n\n"
                f"⏳ <b>DOCX fayl yaratilmoqda...</b>",
                parse_mode="HTML"
            )
        except:
            pass

        content = "\n\n".join(full_parts)
        topic_label = "Fayldan yaratilgan mustaqil ish"

        docx_bytes = await asyncio.get_event_loop().run_in_executor(
            None, generate_docx, "report", topic_label,
            content, db_user.full_name or "Foydalanuvchi", ""
        )

        if not is_admin:
            await deduct_balance(db_user.id, price)

        await create_request(db_user.id, "essay", topic=topic_label)
        await wait_msg.delete()
        await callback.message.bot.send_document(
            chat_id=callback.message.chat.id,
            document=BufferedInputFile(docx_bytes, filename="mustaqil_ish.docx"),
            caption=(
                f"✅ <b>Mustaqil ish tayyor!</b>\n"
                f"📎 Fayl asosida yaratildi ({total_sec} bo'lim)\n"
                f"🌐 Til: {lang_label}\n"
                f"💰 {'Hisobdan yechildi: ' + format_price(price) if not is_admin else '🛡️ Admin — tekin'}"
            ),
            reply_markup=main_menu_kb(),
            parse_mode="HTML"
        )
        await state.clear()

    except Exception as e:
        import logging
        logging.getLogger(__name__).error(f"file_language_chosen error: {e}")
        try:
            await wait_msg.edit_text(f"❌ Xatolik yuz berdi: {e}")
        except:
            await callback.message.bot.send_message(callback.message.chat.id, f"❌ Xatolik: {e}")
        await state.clear()


@router.message(DocumentStates.entering_topic)
async def enter_topic(message: Message, state: FSMContext):
    await state.update_data(topic=message.text.strip())
    await state.set_state(DocumentStates.entering_subject)
    await message.answer(
        "🎓 <b>Fan nomini kiriting:</b>\n"
        "<i>(Masalan: Iqtisodiyot nazariyasi)</i>",
        reply_markup=back_to_menu_kb(),
        parse_mode="HTML"
    )

@router.message(DocumentStates.entering_subject)
async def enter_subject(message: Message, state: FSMContext):
    await state.update_data(subject=message.text.strip())
    await state.set_state(DocumentStates.choosing_quality)
    from keyboards.documents_kb import referat_quality_kb
    await message.answer(
        "💎 <b>Xizmat sifatini tanlang:</b>\n"
        "<i>(Pro versiya 2 barobar sifatli va qimmatroq)</i>",
        reply_markup=referat_quality_kb(),
        parse_mode="HTML"
    )

@router.callback_query(F.data.startswith("ref_quality:"), DocumentStates.choosing_quality)
async def choose_quality(callback: CallbackQuery, state: FSMContext):
    quality = callback.data.split(":")[1]
    is_pro = (quality == "pro")
    await state.update_data(quality=quality, is_pro=is_pro)
    await state.set_state(DocumentStates.choosing_pages)
    from keyboards.documents_kb import page_count_kb
    await callback.message.edit_text(
        f"📑 <b>{'💎 Pro' if is_pro else '✨ Standart'} | Hajmni tanlang:</b>",
        reply_markup=page_count_kb(is_pro=is_pro),
        parse_mode="HTML"
    )
    await callback.answer()

@router.callback_query(F.data.startswith("pages:"), DocumentStates.choosing_pages)
async def choose_pages(callback: CallbackQuery, state: FSMContext, db_user: User):
    _, pages, price = callback.data.split(":")
    await state.update_data(num_pages=int(pages), price=int(price))
    
    # Auto-generate a plan first
    data = await state.get_data()
    topic = data.get("topic")
    
    wait_msg = await callback.message.edit_text(
        "⏳ <b>AI eng yaxshi rejani tayyorlamoqda...</b>",
        parse_mode="HTML"
    )
    
    try:
        # For Pro version, we explicitly ask for a much longer plan
        is_pro = data.get("is_pro", False)
        plan_prompt = "7-10 ta juda batafsil bo'lim" if is_pro else "5-6 ta bo'lim"
        
        plan_data = await generate_document_plan(
            service_type="essay", 
            topic=topic, 
            language=data.get("language", "uz"),
            detail_level="pro" if is_pro else "standard"
        )
        plan = plan_data.get("plan", "")
        sections = [s.strip() for s in plan.split('\n') if s.strip()]
        
        await state.update_data(manual_plan=plan, num_chapters=len(sections))
    except:
        pass
    
    await wait_msg.delete()
    await state.set_state(DocumentStates.reviewing_summary)
    await show_referat_summary(callback.message, state, db_user)
    await callback.answer()

@router.message(F.web_app_data, DocumentStates.reviewing_summary)
async def handle_referat_webapp(message: Message, state: FSMContext, db_user: User):
    try:
        try: await message.delete() 
        except: pass
        data = json.loads(message.web_app_data.data)
        
        if data.get("action") == "plan_update":
            await state.update_data(manual_plan=data.get("manual_plan"), num_chapters=data.get("num_chapters"))
        elif data.get("action") == "content_update":
            await state.update_data(extra_info=data.get("extra_info"), tone=data.get("tone"))
        elif "author" in data:
            # Settings update
            await state.update_data(author=data.get("author"), language=data.get("language"))

        await show_referat_summary(message, state, db_user)
    except Exception as e:
        print(f"WebApp Error: {e}")

async def show_referat_summary(message: Message, state: FSMContext, db_user: User):
    data = await state.get_data()
    topic = data.get("topic")
    subject = data.get("subject")
    pages = data.get("num_pages")
    price = data.get("price")
    author = data.get("author", "Foydalanuvchi")
    quality_val = data.get("quality", "standard")
    quality_label = "💎 Pro" if quality_val == "pro" else "✨ Standart"
    
    text = (
        f"📝 <b>Hujjat Haqida | {quality_label}</b>\n\n"
        f"📚 <b>Mavzu:</b> {topic}\n"
        f"🎓 <b>Fan:</b> {subject}\n"
        f"👤 <b>Muallif:</b> {author}\n"
        f"📄 <b>Sahifalar:</b> {pages} bet\n"
        f"💰 <b>Narxi:</b> {format_price(price)}\n\n"
        f"<i>Reja yoki ma'lumotlarni o'zgartirish uchun tugmalardan foydalaning.</i>"
    )
    
    summary_id = data.get("summary_msg_id")
    if summary_id:
        try: await message.bot.delete_message(chat_id=message.chat.id, message_id=summary_id)
        except: pass

    msg = await message.answer(text, reply_markup=referat_summary_kb(db_user.balance or 0, db_user.full_name or ""), parse_mode="HTML")
    await state.update_data(summary_msg_id=msg.message_id)

@router.message(F.text == "🚀 Yaratish", DocumentStates.reviewing_summary)
async def generate_referat(message: Message, state: FSMContext, db_user: User):
    data = await state.get_data()
    price = data.get("price", 3000)
    topic = data.get("topic")
    pages = data.get("num_pages")
    plan = data.get("manual_plan", "")
    
    if (db_user.balance or 0) < price and db_user.id not in ADMIN_IDS:
        await message.answer("⚠️ Balansingiz yetarli emas!", reply_markup=main_menu_kb())
        return
    # Cancellation support
    cancel_flag = {"cancelled": False}
    await state.update_data(cancel_flag=cancel_flag)
    
    cancel_kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="❌ Bekor qilish", callback_data="doc_cancel_gen")]
    ])

    wait_msg = await message.answer(
        f"⏳ <b>Hujjat yaratilmoqda...</b>\n"
        f"📌 Mavzu: <i>{topic}</i>\n"
        f"📑 Hajmi: {pages} bet\n"
        f"🤖 AI juda batafsil yozmoqda, biroz kuting...",
        parse_mode="HTML",
        reply_markup=cancel_kb
    )

    def is_cancelled():
        return cancel_flag.get("cancelled", False)

    try:
        # Dynamic word count calculation to hit the page target
        sections = [s.strip() for s in plan.split('\n') if s.strip()] if plan else [topic]
        num_sections = len(sections)
        # 1 page is ~250-300 words with our 14pt 1.5 spacing format
        total_target_words = pages * 260 
        words_per_section = int(total_target_words / num_sections)
        
        # Load structure template + quality example
        doc_type = data.get("service_type", "referat")
        template_context = load_template_and_example(doc_type)

        source_text = data.get("source_text", "")
        mode        = data.get("mode", "normal")

        full_content = []

        if mode == "murakkab" and source_text:
            # ── MURAKKAB: execute file instructions exactly, no templates ──────
            import logging
            logging.getLogger(__name__).info(f"MURAKKAB mode: {topic}")
            try:
                await wait_msg.edit_text(
                    "🔴 <b>Murakkab topshiriq bajarilmoqda...</b>\n\n"
                    "📄 AI fayldagi barcha ko'rsatmalarni so'zsiz bajarmoqda.\n"
                    "<i>Biroz sabr qiling...</i>",
                    parse_mode="HTML"
                )
            except:
                pass
            result = await generate_document_section(
                topic=topic,
                section_title="To'liq mustaqil ish",
                extra_details=(
                    f"FOYDALANUVCHI FAYLI — barcha ko'rsatmalar:\n{source_text}\n\n"
                    "QOIDA: Yuqoridagi faylda yozilgan BARCHA ko'rsatmalarni "
                    "SO'ZSIZ bajaring. Hech qanday shablon, qo'shimcha qoida yoki "
                    "cheklov ishlatmang. Faqat faylda yozilganidek, aynan o'sha "
                    f"format va tuzilmada bajaring. Hajm: ~{pages * 280} so'z."
                ),
                language=data.get("language", "uz"),
                quality="pro",
                service_type="article"
            )
            full_content = [result]

        else:
            # ── NORMAL: template-guided section generation ────────────────────
            source_block = (
                f"\n\n--- YUBORILGAN FAYL MATNI (asosiy manba) ---\n"
                f"{source_text[:5000]}\n"
                "--- FAYL MATNI TUGADI ---\n\n"
                "MUHIM: Faylda berilgan ma'lumotlar asosida yozing.\n"
            ) if source_text else ""

            for i, sec in enumerate(sections, 1):
                if is_cancelled():
                    await state.clear()
                    return
                filled = int(((i - 1) / len(sections)) * 10)
                bar = "🟩" * filled + "⬜" * (10 - filled)
                try:
                    await wait_msg.edit_text(
                        f"🤖 <b>AI hujjat yozmoqda...</b>\n\n"
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
                    ) + source_block + f"Ushbu bo'lim aynan {words_per_section} ta so'zdan iborat bo'lsin. " + data.get("extra_info", ""),
                    language=data.get("language", "uz"),
                    quality=data.get("quality", "standard"),
                    service_type="article"
                )
                full_content.append(f"# {sec}\n\n{text}")


        # Add bibliography for Pro
        if data.get("quality") == "pro":
            try:
                bib_title = "FOYDALANILGAN ADABIYOTLAR RO'YXATI"
                try: await wait_msg.edit_text(f"⏳ <b>Hujjat yaratilmoqda...</b>\n📚 <i>Adabiyotlar ro'yxati shakllantirilmoqda...</i>")
                except: pass
                
                bib_text = await generate_document_section(
                    topic=topic, 
                    section_title=bib_title, 
                    extra_details="Mavzu bo'yicha kamida 10-15 ta ilmiy manba va kitoblarni ro'yxat shaklida yozing.",
                    language=data.get("language", "uz"),
                    quality="standard", # Just a list
                    service_type="article"
                )
                full_content.append(f"# {bib_title}\n\n{bib_text}")
            except: pass
        
        content = "\n\n".join(full_content)
        doc_type = data.get("service_type", "referat")
        
        if doc_type == "referat":
            # Use ready-made template for referat
            docx_bytes = await asyncio.get_event_loop().run_in_executor(
                None,
                generate_docx_from_template,
                topic,
                content,
                data.get("author", "Foydalanuvchi"),
                plan,
                data.get("subject", ""),
                "",  # reviewer
            )
        else:
            # Mustaqil ish uses code-generated DOCX
            docx_bytes = await asyncio.get_event_loop().run_in_executor(
                None, 
                generate_docx, 
                "report", 
                topic, 
                content,
                data.get("author", "Foydalanuvchi"),
                plan
            )
        
        if db_user.id not in ADMIN_IDS:
            await deduct_balance(db_user.id, price)
        
        await create_request(user_id=db_user.id, service_type="essay", topic=topic, options={"pages": pages})
        
        await wait_msg.delete()
        file_name = f"Referat_{topic[:20]}.docx"
        await message.answer_document(
            document=BufferedInputFile(docx_bytes, filename=file_name),
            caption=f"✅ <b>Hujjat tayyor!</b>\n\n📌 Mavzu: {topic}\n📑 Hajmi: {pages} bet\n💰 Narxi: {format_price(price)}",
            parse_mode="HTML"
        )
        await state.clear()
    except Exception as e:
        await message.answer(f"❌ Xatolik yuz berdi: {e}")

@router.callback_query(F.data == "doc_cancel_gen")
async def cancel_doc_gen(callback: CallbackQuery, state: FSMContext):
    data = await state.get_data()
    cancel_flag = data.get("cancel_flag")
    if cancel_flag and isinstance(cancel_flag, dict):
        cancel_flag["cancelled"] = True
    await state.clear()
    await callback.message.edit_text("❌ Hujjat yaratish to'xtatildi. Hech narsa hisobdan yechilmadi.")
    await callback.answer()

@router.message(F.text == "❌ Bekor qilish", DocumentStates.reviewing_summary)
async def cancel_referat(message: Message, state: FSMContext):
    await state.clear()
    await message.answer("❌ Amaliyot bekor qilindi.", reply_markup=main_menu_kb())
