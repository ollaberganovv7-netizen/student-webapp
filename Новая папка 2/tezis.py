import asyncio
import io
import os
import json
from aiogram import Router, F
from aiogram.types import (
    Message, CallbackQuery, BufferedInputFile, WebAppInfo,
    InlineKeyboardMarkup, InlineKeyboardButton,
    ReplyKeyboardMarkup, KeyboardButton, ReplyKeyboardRemove
)
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup

from database.models import User
from database.db import create_request, deduct_balance
from keyboards.main_kb import main_menu_kb, back_to_menu_kb
from services.ai_service import client, OPENAI_MODEL
from services.docx_service import generate_docx
from services.template_loader import load_template_and_example
from utils.helpers import format_price
from config import ADMIN_IDS, PRICING

router = Router()

class TezisStates(StatesGroup):
    reviewing_settings = State()   # Normal flow: configure in mini-app then generate
    waiting_for_file   = State()   # File flow: waiting for user to send the source file

# ── Helpers ───────────────────────────────────────────────────────────────────

async def extract_text_from_file(message: Message) -> str:
    """Download and extract text from .txt / .docx / .pdf files."""
    doc = message.document
    file = await message.bot.get_file(doc.file_id)
    buf = io.BytesIO()
    await message.bot.download_file(file.file_path, buf)
    buf.seek(0)
    name = (doc.file_name or "").lower()

    if name.endswith(".txt"):
        return buf.read().decode("utf-8", errors="ignore")

    elif name.endswith(".docx"):
        from docx import Document as DocxDocument
        doc_obj = DocxDocument(buf)
        return "\n".join(p.text for p in doc_obj.paragraphs if p.text.strip())

    elif name.endswith(".pdf"):
        import pdfplumber
        text_parts = []
        with pdfplumber.open(buf) as pdf:
            for page in pdf.pages:
                t = page.extract_text()
                if t:
                    text_parts.append(t)
        return "\n".join(text_parts)

    else:
        raise ValueError("Qo'llab-quvvatlanmaydigan fayl formati. Iltimos .txt, .docx yoki .pdf yuboring.")


def tezis_settings_kb():
    base_url = os.getenv("WEBAPP_URL", "https://arslon.github.io/student_bot/webapp/").split("?")[0]
    if not base_url.endswith("/"): base_url += "/"
    url = f"{base_url}tezis_settings.html"
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="⚙️ Sozlamalarni ochish", web_app=WebAppInfo(url=url))],
            [KeyboardButton(text="📎 Fayldan yaratish")],
            [KeyboardButton(text="🚀 Yaratish")],
            [KeyboardButton(text="❌ Bekor qilish")]
        ],
        resize_keyboard=True
    )

@router.message(F.text == "🎓 Tezis yaratish")
async def start_tezis(message: Message, state: FSMContext, db_user: User):
    balance = db_user.balance or 0
    await state.set_state(TezisStates.reviewing_settings)
    
    text = (
        "🎓 <b>Tezis yaratish</b>\n\n"
        f"💳 <b>Balansingiz:</b> {format_price(balance)}\n\n"
        "ℹ️ <i>1. Sozlamalarni oching va to'ldiring\n2. SAQLASH VA YARATISH tugmasini bosing</i>"
    )
    await message.answer(text, reply_markup=tezis_settings_kb(), parse_mode="HTML")

@router.message(F.web_app_data, TezisStates.reviewing_settings)
async def tezis_webapp_data(message: Message, state: FSMContext, db_user: User):
    try:
        data = json.loads(message.web_app_data.data)
        if data.get("type") != "tezis_settings":
            return
        await state.update_data(
            topic=data["topic"],
            name=data["author"],
            university=data.get("university", ""),
            structure=data["structure"],
            lang=data["language"],
            pages=data["pages"]
        )
        # Auto-start generation immediately after saving from WebApp
        await tezis_generate(message, state, db_user)
    except Exception as e:
        print("WebApp parsing error:", e)

# Also handle webapp data coming from waiting_for_file state
@router.message(F.web_app_data, TezisStates.waiting_for_file)
async def tezis_webapp_data_file(message: Message, state: FSMContext, db_user: User):
    await tezis_webapp_data(message, state, db_user)

async def show_tezis_settings(message_or_callback, state: FSMContext, db_user: User):
    data = await state.get_data()
    topic = data.get("topic")
    name = data.get("name")
    structure = data.get("structure", "Standart")
    lang = data.get("lang", "uz")
    pages = data.get("pages", 5)
    
    price = PRICING.get(f"tezis_{pages}", 4000)
    balance = db_user.balance or 0
    is_admin = db_user.id in ADMIN_IDS
    if is_admin: balance = 9999999
    
    lang_full = "O'zbekcha 🇺🇿" if lang == "uz" else "Русский 🇷🇺" if lang == "ru" else "English 🇬🇧"
    
    text = (
        "🎓 <b>Tezis ma'lumotlari</b>\n\n"
        f"📝 <b>Mavzu:</b> {topic}\n"
        f"👤 <b>Muallif:</b> {name}\n"
        f"📐 <b>Tuzilma:</b> {structure}\n"
        f"🌐 <b>Til:</b> {lang_full}\n"
        f"📄 <b>Sahifalar soni:</b> 1-{pages} ta\n\n"
        f"💰 <b>Narx:</b> {format_price(price)}\n"
        f"💳 <b>Sizning balansingiz:</b> {'Yetarli ✅' if balance >= price else 'Yetarli emas ❌'}"
    )
    
    kb = tezis_settings_kb(balance, price)
    
    if isinstance(message_or_callback, Message):
        try:
            await message_or_callback.delete()
        except:
            pass
        await message_or_callback.answer(text, reply_markup=kb, parse_mode="HTML")
    else:
        try:
            await message_or_callback.message.delete()
        except:
            pass
        await message_or_callback.message.answer(text, reply_markup=kb, parse_mode="HTML")

# ── File upload flow ──────────────────────────────────────────────────────

@router.message(F.text == "📎 Fayldan yaratish", TezisStates.reviewing_settings)
async def ask_for_file(message: Message, state: FSMContext):
    await state.set_state(TezisStates.waiting_for_file)
    await message.answer(
        "📎 <b>Faylingizni yuboring</b>\n\n"
        "Quyidagi format qo'llab-quvvatlanadi:\n"
        "• <b>.txt</b> — oddiy matn\n"
        "• <b>.docx</b> — Word hujjati\n"
        "• <b>.pdf</b> — PDF hujjati\n\n"
        "<i>Fayl yuborilgandan so'ng, sozlamalarni to'ldiring va yaratish boshlaydi.</i>",
        reply_markup=ReplyKeyboardMarkup(
            keyboard=[[KeyboardButton(text="❌ Bekor qilish")]],
            resize_keyboard=True
        ),
        parse_mode="HTML"
    )

@router.message(F.document, TezisStates.waiting_for_file)
async def receive_file(message: Message, state: FSMContext, db_user: User):
    proc = await message.answer("⏳ Fayl o'qilmoqda...")
    try:
        source_text = await extract_text_from_file(message)
        if not source_text.strip():
            await proc.edit_text("❌ Fayl bo'sh yoki o'qib bo'lmadi. Boshqa fayl yuboring.")
            return
        
        # Trim to reasonable context size (~12000 chars)
        source_text = source_text[:12000]
        await state.update_data(source_text=source_text)
        await state.set_state(TezisStates.reviewing_settings)
        
        await proc.delete()
        await message.answer(
            f"✅ <b>Fayl muvaffaqiyatli o'qildi!</b>\n"
            f"📝 Hajm: ~{len(source_text.split())} so'z\n\n"
            f"Endi sozlamalarni to'ldiring (mavzu, muallif, sahifalar soni).\n"
            f"ℹ️ Agar mavzu bo'sh qolsa, AI fayldan o'zi aniqlaydi.",
            reply_markup=tezis_settings_kb(),
            parse_mode="HTML"
        )
    except ValueError as e:
        await proc.edit_text(f"❌ {e}")
    except Exception as e:
        await proc.edit_text(f"❌ Xatolik: {str(e)}")

@router.message(F.text == "❌ Bekor qilish", TezisStates.reviewing_settings)
async def tezis_cancel(message: Message, state: FSMContext):
    await message.answer("❌ Amaliyot bekor qilindi.", reply_markup=main_menu_kb())
    await state.clear()

@router.message(F.text == "❌ Bekor qilish", TezisStates.waiting_for_file)
async def tezis_cancel_file(message: Message, state: FSMContext):
    await message.answer("❌ Amaliyot bekor qilindi.", reply_markup=main_menu_kb())
    await state.clear()

@router.message(F.text == "🚀 Yaratish", TezisStates.reviewing_settings)
async def tezis_generate(message: Message, state: FSMContext, db_user: User):
    data = await state.get_data()
    pages = data.get("pages", 5)
    topic = data.get("topic")
    name = data.get("name")
    
    if not topic or not name:
        await message.answer("❗ Iltimos, oldin <b>⚙️ Sozlamalarni ochish</b> tugmasini bosib, mavzu va ismingizni kiriting.", parse_mode="HTML")
        return
        
    structure = data.get("structure", "Standart")
    lang      = data.get("lang", "uz")
    university = data.get("university", "")
    
    price = PRICING.get(f"tezis_{pages}", 4000)
    is_admin = db_user.id in ADMIN_IDS
    
    if not is_admin and (db_user.balance or 0) < price:
        await message.answer("Balans yetarli emas!", reply_markup=main_menu_kb())
        return

    # Cancellation support
    cancel_flag = {"cancelled": False}
    await state.update_data(cancel_flag=cancel_flag)
    
    cancel_kb_inline = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="❌ Bekor qilish", callback_data="tezis_cancel_gen")]
    ])
    
    # Remove keyboard first, then send editable wait message
    try:
        await message.answer("⏳", reply_markup=ReplyKeyboardRemove())
    except:
        pass
    wait_msg = await message.bot.send_message(
        chat_id=message.chat.id,
        text="⏳ <b>Tezis yaratilmoqda...</b>\n📝 Bo'limlar ketma-ket yozilmoqda, biroz sabr qiling (3-5 daqiqa)...",
        parse_mode="HTML",
        reply_markup=cancel_kb_inline
    )
    
    def is_cancelled():
        return cancel_flag.get("cancelled", False)
    
    # Target words per page tier (1 page = ~250 words)
    total_words = pages * 250
    
    try:
        lang_instruction = "O'zbek tilida (Lotin alifbosida)" if lang == "uz" else "Rus tilida (На русском)" if lang == "ru" else "Ingliz tilida (In English)"
        source_text = data.get("source_text", "")
        
        template_context = load_template_and_example("tezis")
        
        university_line = f"Universitet/Muassasa: {university}" if university else ""
        system_prompt = (
            f"{template_context}\n"
            "You are a world-class academic professor with 30+ years of research experience. "
            "Before writing each section, deeply think through the key arguments, sub-points, "
            "historical context, and relevant data you will include. "
            "Then write a deeply analytical, evidence-based, and comprehensive text. "
            "Every paragraph must contain specific facts, real-world examples, or scholarly reasoning. "
            "Never write a short or superficial paragraph. Always expand and elaborate. "
            "Your writing must feel like it comes from a peer-reviewed academic journal.\n\n"
            "CRITICAL RULES FOR DOCUMENT BODY:\n"
            "- Do NOT write 'Mavzu:', 'Muallif:', 'Universiteti:', 'Author:', 'Topic:' labels anywhere.\n"
            "- Do NOT repeat the thesis title inside the body text.\n"
            "- The title page is handled separately — write ONLY the section content."
        )
        
        # Shared source context block injected into each section prompt
        source_block = ""
        if source_text:
            source_block = (
                f"\n\n--- SOURCE MATERIAL (use this as your primary reference) ---\n"
                f"{source_text}\n"
                f"--- END OF SOURCE MATERIAL ---\n\n"
                "IMPORTANT: Base your writing on the source material above. "
                "Expand, analyze, and restructure it into proper academic sections. "
                "Do NOT copy verbatim. Rewrite in your own academic voice."
            )
        
        # Define sections based on structure
        if structure == "IMRAD":
            words_per_section = total_words // 6
            sections = [
                ("KIRISH",
                 f"Write a deeply researched INTRODUCTION of exactly {words_per_section} words for: '{topic}'. "
                 f"Include: the historical evolution of the topic, its global and regional significance, "
                 f"key unresolved problems, specific statistics or data, and the research objectives. "
                 f"Write in {lang_instruction}. Do NOT repeat the thesis title."),
                ("ANNOTATSIYA",
                 f"Write a concise abstract of exactly {words_per_section // 2} words for a thesis on: '{topic}'. "
                 f"Cover purpose, methodology, key findings, and implications. Write ONLY in {lang_instruction}. "
                 f"Do NOT write 'Mavzu:', 'Muallif:', or the title at the top."),
                ("KALIT SO'ZLAR",
                 f"List exactly 7 relevant academic keywords for a thesis on: '{topic}'. "
                 f"Format: Kalit so'zlar: word1, word2, word3, word4, word5, word6, word7"),
                ("ADABIYOTLAR TAHLILI",
                 f"Write a thorough LITERATURE REVIEW of exactly {words_per_section} words for: '{topic}'. "
                 f"Critically analyze major scholarly perspectives, theories, and academic debates. "
                 f"Identify gaps in current research. Write in {lang_instruction}."),
                ("TADQIQOT METODOLOGIYASI",
                 f"Write a detailed METHODOLOGY of exactly {words_per_section} words for: '{topic}'. "
                 f"Describe research design, data sources, analytical methods with academic rationale. "
                 f"Write in {lang_instruction}."),
                ("NATIJALAR VA MUHOKAMA",
                 f"Write comprehensive RESULTS AND DISCUSSION of exactly {words_per_section} words for: '{topic}'. "
                 f"Present specific findings, compare with existing literature, discuss implications. "
                 f"Write in {lang_instruction}."),
                ("XULOSA",
                 f"Write a CONCLUSION of exactly {words_per_section // 2} words for: '{topic}'. "
                 f"Synthesize key findings, state contributions, give concrete recommendations. "
                 f"Write in {lang_instruction}."),
                ("FOYDALANILGAN ADABIYOTLAR",
                 f"List 8–12 real academic references relevant to: '{topic}'. "
                 f"Format: Author F.I. Title. Journal, Year. Vol(No). P. XX–XX."),
            ]
        else:
            words_per_section = total_words // 4
            sections = [
                ("KIRISH",
                 f"Write a deeply researched INTRODUCTION of exactly {words_per_section} words for: '{topic}'. "
                 f"Include historical background, global context, key problems, data, and thesis aims. "
                 f"Write in {lang_instruction}. Do NOT repeat the thesis title."),
                ("ANNOTATSIYA",
                 f"Write a concise abstract of exactly {words_per_section // 3} words for a thesis on: '{topic}'. "
                 f"Cover purpose, methods, key findings. Write ONLY in {lang_instruction}. "
                 f"Do NOT write 'Mavzu:', 'Muallif:', or the title at the top."),
                ("KALIT SO'ZLAR",
                 f"List exactly 7 relevant academic keywords for a thesis on: '{topic}'. "
                 f"Format: Kalit so'zlar: word1, word2, word3, word4, word5, word6, word7"),
                ("ASOSIY QISM",
                 f"Write a deeply analytical MAIN BODY of exactly {words_per_section} words for: '{topic}'. "
                 f"Use 2–3 subheadings. Each must include deep analysis, real case studies, comparative data. "
                 f"Write in {lang_instruction}."),
                ("XULOSA",
                 f"Write a CONCLUSION of exactly {words_per_section // 2} words for: '{topic}'. "
                 f"Synthesize key arguments, state contributions, give future recommendations. "
                 f"Write in {lang_instruction}."),
                ("FOYDALANILGAN ADABIYOTLAR",
                 f"List 8–12 real academic references relevant to: '{topic}'. "
                 f"Format: Author F.I. Title. Journal, Year. Vol(No). P. XX–XX."),
            ]
        
        # Generate each section separately
        full_parts = []
        async def show_progress(step: int, total: int, phase: str, section_name: str):
            done   = "🟩" * (step - 1)
            cur    = "🟨"
            left   = "⬜" * (total - step)
            bar    = done + cur + left
            if phase == "plan":
                text = (
                    f"🧠 <b>AI chuqur tahlil qilmoqda...</b>\n\n"
                    f"📄 <b>Bo'lim {step}/{total}</b>\n"
                    f"{bar}\n\n"
                    f"🔍 <b>1-bosqich — Reja:</b> <i>{section_name}</i>\n\n"
                    f"📝 <b>Mavzu:</b> {topic}\n"
                    f"<i>AI avval fikrlaydi, keyin yozadi — sifat uchun sabr qiling...</i>"
                )
            else:
                text = (
                    f"✍️ <b>AI yozmoqda...</b>\n\n"
                    f"📄 <b>Bo'lim {step}/{total}</b>\n"
                    f"{bar}\n\n"
                    f"📝 <b>2-bosqich — Yozish:</b> <i>{section_name}</i>\n\n"
                    f"📝 <b>Mavzu:</b> {topic}\n"
                    f"<i>Reja tayyor — endi to'liq matn yozilmoqda...</i>"
                )
            try:
                await wait_msg.edit_text(text, parse_mode="HTML")
                await asyncio.sleep(0.5)   # avoid Telegram rate limit
            except Exception as e:
                import logging
                logging.getLogger(__name__).warning(f"Progress edit failed (section {step}/{total}): {e}")

        for i, (section_title, section_prompt) in enumerate(sections, 1):
            # Check cancellation before each section
            if is_cancelled():
                await state.clear()
                return
            total = len(sections)

            await show_progress(i, total, "plan", section_title)

            full_prompt = (
                f"Topic of the full thesis: '{topic}'\n"
                f"Author: {name}\n\n"
                f"Your task: {section_prompt}\n"
                f"{source_block}\n"
                "MANDATORY RULES:\n"
                "1. Think deeply before writing. Consider all relevant angles, sub-arguments, and evidence.\n"
                "2. Write ONLY in full academic paragraphs. NO bullet points. NO numbered lists.\n"
                "3. Every paragraph must contain: a claim, supporting evidence or example, and analysis.\n"
                "4. Include real statistics, historical facts, or specific case studies wherever relevant.\n"
                "5. Do NOT use markdown symbols like **, #, or ```.\n"
                "6. Write EXACTLY the requested amount of words. Do not make it much longer or shorter."
            )

            # ── Step 1: Deep pre-analysis (chain-of-thought planning) ──────────
            planning_prompt = (
                f"You are preparing to write the '{section_title}' section of an academic thesis on: '{topic}'.\n\n"
                "Perform a DEEP pre-writing analysis before writing anything. Think like a professor:\n"
                "1. What are the 5 most critical academic arguments for this section? List each with supporting evidence.\n"
                "2. What specific statistics, dates, historical facts, or real case studies will you use?\n"
                "3. What are the leading scholarly theories and which scholars support each position?\n"
                "4. What are the counterarguments and how will you address them?\n"
                "5. What is the most logical paragraph order for maximum academic impact?\n\n"
                "Write your detailed analysis plan (250–400 words). This blueprint will guide the actual writing."
            )
            try:
                plan_response = await client.chat.completions.create(
                    model=OPENAI_MODEL,
                    messages=[
                        {"role": "system", "content": system_prompt},
                        {"role": "user",   "content": planning_prompt}
                    ],
                    temperature=0.5,
                    max_tokens=1500
                )
                section_plan = plan_response.choices[0].message.content.strip()
            except Exception:
                section_plan = ""

            # ── Step 2: Write the actual section using the plan ───────────────
            if section_plan:
                full_prompt += f"\n\nYOUR PRE-WRITING ANALYSIS (use this as your writing blueprint):\n{section_plan}"

            await show_progress(i, total, "write", section_title)

            response = await client.chat.completions.create(
                model=OPENAI_MODEL,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user",   "content": full_prompt}
                ],
                temperature=0.8,
                max_tokens=10000
            )

            section_text = response.choices[0].message.content.strip()
            full_parts.append(f"# {section_title}\n\n{section_text}")

        
        content = "\n\n".join(full_parts)
        
        try:
            await wait_msg.edit_text("⏳ <b>Hujjat (DOCX) tayyorlanmoqda...</b>", parse_mode="HTML")
        except:
            pass
        
        docx_bytes = await asyncio.get_event_loop().run_in_executor(
            None, generate_docx, "tezis", topic, content, name, ""
        )
        
        if not is_admin:
            await deduct_balance(db_user.id, price)
            status_note = f"💰 <b>Hisobdan yechildi:</b> {format_price(price)}"
        else:
            status_note = "🛡️ <b>Admin uchun tekin</b>"
            
        await create_request(db_user.id, "tezis", topic=topic)
        
        await wait_msg.delete()
        await message.answer_document(
            BufferedInputFile(docx_bytes, filename=f"tezis_{topic[:20]}.docx"),
            caption=(
                f"✅ <b>Tezis tayyor!</b>\n\n"
                f"📝 <b>Mavzu:</b> {topic}\n"
                f"👤 <b>Muallif:</b> {name}\n"
                + (f"🏫 <b>Universitet:</b> {university}\n" if university else "")
                + f"📄 <b>Hajmi:</b> {pages} sahifa (~{total_words} so'z)\n"
                f"────────────────────\n"
                f"{status_note}"
            ),
            reply_markup=main_menu_kb(),
            parse_mode="HTML"
        )
        await state.clear()
        
    except Exception as e:
        try:
            await wait_msg.edit_text(f"❌ <b>Xatolik yuz berdi:</b> {str(e)}", parse_mode="HTML")
        except:
            await message.answer(f"❌ <b>Xatolik yuz berdi:</b> {str(e)}", parse_mode="HTML")
        await state.clear()

@router.callback_query(F.data == "tezis_cancel_gen")
async def cancel_tezis_gen(callback: CallbackQuery, state: FSMContext):
    data = await state.get_data()
    cancel_flag = data.get("cancel_flag")
    if cancel_flag and isinstance(cancel_flag, dict):
        cancel_flag["cancelled"] = True
    await state.clear()
    await callback.message.edit_text("❌ Tezis yaratish to'xtatildi. Hech narsa hisobdan yechilmadi.")
    await callback.answer()
