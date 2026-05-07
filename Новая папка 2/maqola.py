"""
handlers/maqola.py

Ilmiy maqola yaratish handler — tezis kodi asosida qurilgan.
Foydalanuvchi mavzu va parametrlarni kiritadi, AI esa doc_templates/maqola_structure.txt
va doc_examples/maqola_example.* fayllarini o'qib, sifatli maqola yozadi.
"""

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
from keyboards.main_kb import main_menu_kb
from services.ai_service import client, OPENAI_MODEL
from services.docx_service import generate_docx
from services.template_loader import load_template_and_example
from utils.helpers import format_price
from config import ADMIN_IDS, PRICING

router = Router()

class MaqolaStates(StatesGroup):
    reviewing_settings = State()
    waiting_for_file   = State()


# ── File extractor (shared logic) ────────────────────────────────────────────

async def extract_text_from_file(message: Message) -> str:
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
        parts = []
        with pdfplumber.open(buf) as pdf:
            for page in pdf.pages:
                t = page.extract_text()
                if t:
                    parts.append(t)
        return "\n".join(parts)
    else:
        raise ValueError("Qo'llab-quvvatlanmaydigan fayl formati. Iltimos .txt, .docx yoki .pdf yuboring.")


# ── Keyboard ──────────────────────────────────────────────────────────────────

def maqola_settings_kb():
    base_url = os.getenv("WEBAPP_URL", "https://arslon.github.io/student_bot/webapp/").split("?")[0]
    if not base_url.endswith("/"): base_url += "/"
    url = f"{base_url}tezis_settings.html"   # maqola shares the same settings WebApp
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="⚙️ Sozlamalarni ochish", web_app=WebAppInfo(url=url))],
            [KeyboardButton(text="📎 Fayldan yaratish")],
            [KeyboardButton(text="🚀 Yaratish")],
            [KeyboardButton(text="❌ Bekor qilish")]
        ],
        resize_keyboard=True
    )


# ── Entry point ───────────────────────────────────────────────────────────────

@router.message(F.text == "📝 Maqola yaratish")
async def start_maqola(message: Message, state: FSMContext, db_user: User):
    balance = db_user.balance or 0
    await state.set_state(MaqolaStates.reviewing_settings)
    await message.answer(
        "📝 <b>Ilmiy maqola yaratish</b>\n\n"
        f"💳 <b>Balansingiz:</b> {format_price(balance)}\n\n"
        "ℹ️ <i>1. Sozlamalarni oching va to'ldiring\n"
        "2. SAQLASH VA YARATISH tugmasini bosing</i>",
        reply_markup=maqola_settings_kb(),
        parse_mode="HTML"
    )


# ── WebApp data handlers ──────────────────────────────────────────────────────

@router.message(F.web_app_data, MaqolaStates.reviewing_settings)
async def maqola_webapp_data(message: Message, state: FSMContext, db_user: User):
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
        await maqola_generate(message, state, db_user)
    except Exception as e:
        print("Maqola WebApp parsing error:", e)


@router.message(F.web_app_data, MaqolaStates.waiting_for_file)
async def maqola_webapp_data_file(message: Message, state: FSMContext, db_user: User):
    await maqola_webapp_data(message, state, db_user)


# ── File upload flow ──────────────────────────────────────────────────────────

@router.message(F.text == "📎 Fayldan yaratish", MaqolaStates.reviewing_settings)
async def maqola_ask_for_file(message: Message, state: FSMContext):
    await state.set_state(MaqolaStates.waiting_for_file)
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


@router.message(F.document, MaqolaStates.waiting_for_file)
async def maqola_receive_file(message: Message, state: FSMContext, db_user: User):
    proc = await message.answer("⏳ Fayl o'qilmoqda...")
    try:
        source_text = await extract_text_from_file(message)
        if not source_text.strip():
            await proc.edit_text("❌ Fayl bo'sh yoki o'qib bo'lmadi. Boshqa fayl yuboring.")
            return
        source_text = source_text[:12000]
        await state.update_data(source_text=source_text)
        await state.set_state(MaqolaStates.reviewing_settings)
        await proc.delete()
        await message.answer(
            f"✅ <b>Fayl muvaffaqiyatli o'qildi!</b>\n"
            f"📝 Hajm: ~{len(source_text.split())} so'z\n\n"
            "Endi sozlamalarni to'ldiring (mavzu, muallif, sahifalar soni).\n"
            "ℹ️ Agar mavzu bo'sh qolsa, AI fayldan o'zi aniqlaydi.",
            reply_markup=maqola_settings_kb(),
            parse_mode="HTML"
        )
    except ValueError as e:
        await proc.edit_text(f"❌ {e}")
    except Exception as e:
        await proc.edit_text(f"❌ Xatolik: {str(e)}")


# ── Cancel handlers ───────────────────────────────────────────────────────────

@router.message(F.text == "❌ Bekor qilish", MaqolaStates.reviewing_settings)
async def maqola_cancel(message: Message, state: FSMContext):
    await state.clear()
    await message.answer("❌ Amaliyot bekor qilindi.", reply_markup=main_menu_kb())


@router.message(F.text == "❌ Bekor qilish", MaqolaStates.waiting_for_file)
async def maqola_cancel_file(message: Message, state: FSMContext):
    await state.clear()
    await message.answer("❌ Amaliyot bekor qilindi.", reply_markup=main_menu_kb())


# ── Main generation ───────────────────────────────────────────────────────────

@router.message(F.text == "🚀 Yaratish", MaqolaStates.reviewing_settings)
async def maqola_generate(message: Message, state: FSMContext, db_user: User):
    data = await state.get_data()
    pages = data.get("pages", 5)
    topic = data.get("topic")
    name  = data.get("name")

    if not topic or not name:
        await message.answer(
            "❗ Iltimos, oldin <b>⚙️ Sozlamalarni ochish</b> tugmasini bosib, "
            "mavzu va ismingizni kiriting.",
            parse_mode="HTML"
        )
        return

    structure  = data.get("structure", "Standart")
    lang       = data.get("lang", "uz")
    university = data.get("university", "")

    price    = PRICING.get(f"maqola_{pages}", 4000)
    is_admin = db_user.id in ADMIN_IDS

    if not is_admin and (db_user.balance or 0) < price:
        await message.answer("Balans yetarli emas!", reply_markup=main_menu_kb())
        return

    # Cancellation support
    cancel_flag = {"cancelled": False}
    await state.update_data(cancel_flag=cancel_flag)
    
    cancel_kb_inline = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="❌ Bekor qilish", callback_data="maqola_cancel_gen")]
    ])
    
    # Remove keyboard first, then send editable wait message
    try:
        await message.answer("⏳", reply_markup=ReplyKeyboardRemove())
    except:
        pass
    wait_msg = await message.bot.send_message(
        chat_id=message.chat.id,
        text="⏳ <b>Maqola yaratilmoqda...</b>\n"
             "📝 Bo'limlar ketma-ket yozilmoqda, biroz sabr qiling (3–5 daqiqa)...",
        parse_mode="HTML",
        reply_markup=cancel_kb_inline
    )
    
    def is_cancelled():
        return cancel_flag.get("cancelled", False)

    total_words = pages * 250

    try:
        lang_instruction = (
            "O'zbek tilida (Lotin alifbosida)" if lang == "uz"
            else "Rus tilida (На русском)" if lang == "ru"
            else "Ingliz tilida (In English)"
        )
        source_text = data.get("source_text", "")

        # Load knowledge base rules + structure template + quality example
        template_context = load_template_and_example("maqola")

        university_line = f"Universitet/Muassasa: {university}\n" if university else ""
        system_prompt = (
            f"{template_context}\n"
            "You are a world-class academic researcher and journal editor with 30+ years of experience. "
            "You write rigorous, evidence-based academic articles that meet international publication standards. "
            "Every paragraph must contain a clear argument, supporting evidence, and critical analysis. "
            "Never write a superficial or generic paragraph. "
            "Your writing must meet the standards of a peer-reviewed academic journal.\n\n"
            "CRITICAL RULES FOR DOCUMENT BODY:\n"
            "- Do NOT write 'Mavzu:', 'Muallif:', 'Universiteti:', 'Author:', 'Topic:' labels anywhere in the body.\n"
            "- Do NOT repeat the article title inside the body text.\n"
            "- The title page is handled separately — write ONLY the section content."
        )

        source_block = ""
        if source_text:
            source_block = (
                f"\n\n--- SOURCE MATERIAL (use as primary reference) ---\n"
                f"{source_text}\n"
                "--- END OF SOURCE MATERIAL ---\n\n"
                "IMPORTANT: Base your writing on the source material above. "
                "Expand, analyse, and restructure it into proper academic sections. "
                "Do NOT copy verbatim. Rewrite in your own academic voice."
            )

        # ── Section definitions ───────────────────────────────────────────────
        if structure == "IMRAD":
            words_per_section = total_words // 6
            sections = [
                ("ANNOTATSIYA",
                 f"Write a concise abstract of exactly {words_per_section // 2} words for an academic article on: '{topic}'. "
                 f"Cover: problem, methodology, key findings, conclusion. Write ONLY in {lang_instruction}. "
                 f"Do NOT include title, author name, or labels like 'Mavzu:' or 'Muallif:'."),
                ("KALIT SO'ZLAR",
                 f"List 7 keywords relevant to the academic article on: '{topic}'. "
                 f"Format exactly: Kalit so'zlar: word1, word2, word3, word4, word5, word6, word7"),
                ("KIRISH",
                 f"Write a deeply researched INTRODUCTION of exactly {words_per_section} words for: '{topic}'. "
                 f"Include historical evolution, global significance, key problems, statistics, and research objectives. "
                 f"Write in {lang_instruction}. Do NOT repeat the article title."),
                ("ADABIYOTLAR TAHLILI",
                 f"Write a thorough LITERATURE REVIEW of exactly {words_per_section} words for: '{topic}'. "
                 f"Critically analyse major scholarly perspectives and academic debates. "
                 f"Write in {lang_instruction}."),
                ("METODOLOGIYA",
                 f"Write a detailed METHODOLOGY of exactly {words_per_section} words for: '{topic}'. "
                 f"Describe research design, data sources, analytical methods. Write in {lang_instruction}."),
                ("MUHOKAMA VA NATIJALAR",
                 f"Write comprehensive RESULTS AND DISCUSSION of exactly {words_per_section} words for: '{topic}'. "
                 f"Present specific findings, compare with literature, discuss implications. Write in {lang_instruction}."),
                ("XULOSA VA TAKLIFLAR",
                 f"Write a CONCLUSION of exactly {words_per_section // 2} words for: '{topic}'. "
                 f"Synthesise findings, state contributions, give recommendations. Write in {lang_instruction}."),
                ("FOYDALANILGAN ADABIYOTLAR RO'YXATI",
                 f"List 8–12 real academic references relevant to: '{topic}'. "
                 f"Format: Author F.I. Title. Journal, Year. Vol(No). P. XX–XX."),
            ]
        else:
            words_per_section = total_words // 4
            sections = [
                ("ANNOTATSIYA",
                 f"Write a concise abstract of exactly {words_per_section // 3} words for an academic article on: '{topic}'. "
                 f"Cover: problem, methodology, key findings. Write ONLY in {lang_instruction}. "
                 f"Do NOT include title, author name, or labels like 'Mavzu:' or 'Muallif:'."),
                ("KALIT SO'ZLAR",
                 f"List 7 keywords relevant to: '{topic}'. "
                 f"Format: Kalit so'zlar: word1, word2, word3, word4, word5, word6, word7"),
                ("KIRISH",
                 f"Write a deeply researched INTRODUCTION of exactly {words_per_section} words for: '{topic}'. "
                 f"Include historical background, key problems, statistics, and aims. "
                 f"Write in {lang_instruction}. Do NOT repeat the article title."),
                ("ASOSIY QISM",
                 f"Write a MAIN BODY of exactly {words_per_section} words for: '{topic}'. "
                 f"Use 2–3 subheadings. Each must include deep analysis, case studies, comparative data. "
                 f"Write in {lang_instruction}."),
                ("XULOSA VA TAKLIFLAR",
                 f"Write a CONCLUSION of exactly {words_per_section // 2} words for: '{topic}'. "
                 f"Synthesise arguments, state contributions, provide recommendations. Write in {lang_instruction}."),
                ("FOYDALANILGAN ADABIYOTLAR RO'YXATI",
                 f"List 8–12 real academic references relevant to: '{topic}'. "
                 f"Format: Author F.I. Title. Journal, Year. Vol(No). P. XX–XX."),
            ]

        # ── Section generation loop ───────────────────────────────────────────
        full_parts = []

        async def show_progress(step: int, total: int, phase: str, section_name: str):
            done  = "🟩" * (step - 1)
            cur   = "🟨"
            left  = "⬜" * (total - step)
            bar   = done + cur + left
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
                await asyncio.sleep(0.5)
            except Exception as e:
                import logging
                logging.getLogger(__name__).warning(f"Maqola progress edit failed ({step}/{total}): {e}")

        for i, (section_title, section_prompt) in enumerate(sections, 1):
            # Check cancellation before each section
            if is_cancelled():
                await state.clear()
                return
            total_sec = len(sections)

            await show_progress(i, total_sec, "plan", section_title)

            full_prompt = (
                f"Topic of the full academic article: '{topic}'\n"
                f"Author: {name}\n"
                + (f"University: {university}\n" if university else "")
                + f"\nYour task: {section_prompt}\n"
                f"{source_block}\n"
                "MANDATORY RULES:\n"
                "1. Think deeply before writing. Consider all relevant angles, arguments, and evidence.\n"
                "2. Write ONLY in full academic paragraphs. NO bullet points. NO numbered lists.\n"
                "3. Every paragraph must contain: a claim, supporting evidence or example, and analysis.\n"
                "4. Include real statistics, historical facts, or specific case studies wherever relevant.\n"
                "5. Do NOT use markdown symbols like **, #, or ```.\n"
                "6. Write EXACTLY the requested amount of words. Do not make it much longer or shorter."
            )

            # Step 1: planning
            planning_prompt = (
                f"You are preparing to write the '{section_title}' section of an academic article on: '{topic}'.\n\n"
                "Perform a DEEP pre-writing analysis before writing anything. Think like a senior journal editor:\n"
                "1. What are the 5 most critical academic arguments for this section? List each with supporting evidence.\n"
                "2. What specific statistics, dates, historical facts, or real case studies will you use?\n"
                "3. What are the leading scholarly theories and which scholars support each position?\n"
                "4. What are the counterarguments and how will you address them?\n"
                "5. What is the most logical paragraph order for maximum academic impact?\n\n"
                "Write your detailed analysis plan (250–400 words). This blueprint will guide the actual writing."
            )
            try:
                plan_resp = await client.chat.completions.create(
                    model=OPENAI_MODEL,
                    messages=[
                        {"role": "system", "content": system_prompt},
                        {"role": "user",   "content": planning_prompt}
                    ],
                    temperature=0.5,
                    max_tokens=1500
                )
                section_plan = plan_resp.choices[0].message.content.strip()
            except Exception:
                section_plan = ""

            if section_plan:
                full_prompt += f"\n\nYOUR PRE-WRITING ANALYSIS (use this as your writing blueprint):\n{section_plan}"

            await show_progress(i, total_sec, "write", section_title)

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
            None, generate_docx, "maqola", topic, content, name, "", {"university": university}
        )

        if not is_admin:
            await deduct_balance(db_user.id, price)
            status_note = f"💰 <b>Hisobdan yechildi:</b> {format_price(price)}"
        else:
            status_note = "🛡️ <b>Admin uchun tekin</b>"

        await create_request(db_user.id, "maqola", topic=topic)

        await wait_msg.delete()
        await message.answer_document(
            BufferedInputFile(docx_bytes, filename=f"maqola_{topic[:20]}.docx"),
            caption=(
                f"✅ <b>Maqola tayyor!</b>\n\n"
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

@router.callback_query(F.data == "maqola_cancel_gen")
async def cancel_maqola_gen(callback: CallbackQuery, state: FSMContext):
    data = await state.get_data()
    cancel_flag = data.get("cancel_flag")
    if cancel_flag and isinstance(cancel_flag, dict):
        cancel_flag["cancelled"] = True
    await state.clear()
    await callback.message.edit_text("❌ Maqola yaratish to'xtatildi. Hech narsa hisobdan yechilmadi.")
    await callback.answer()
