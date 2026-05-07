import os
import json
import asyncio
import logging
from openai import AsyncOpenAI
from config import OPENAI_API_KEY, OPENAI_MODEL

logger = logging.getLogger(__name__)
client = AsyncOpenAI(api_key=OPENAI_API_KEY)


async def _call_ai(messages, max_tokens=3000, temperature=0.8, json_mode=False, retries=3):
    """Call OpenAI API with automatic retry on failure."""
    kwargs = {
        "model": OPENAI_MODEL,
        "messages": messages,
        "max_tokens": max_tokens,
        "temperature": temperature,
    }
    if json_mode:
        kwargs["response_format"] = {"type": "json_object"}
    
    for attempt in range(retries):
        try:
            resp = await client.chat.completions.create(**kwargs)
            return resp.choices[0].message.content.strip()
        except Exception as e:
            logger.warning(f"AI call attempt {attempt+1}/{retries} failed: {e}")
            if attempt < retries - 1:
                await asyncio.sleep(2 ** attempt)  # 1s, 2s, 4s
            else:
                raise

async def generate_document_plan(service_type: str, topic: str, language: str = "uz", detail_level: str = "standard") -> dict:
    prompt = (
        f"Mavzu: {topic}\n"
        f"Til: {language}\n\n"
        "Vazifa: Ushbu mavzu bo'yicha akademik reja tuzing.\n"
        "REJA FAQAT 4-6 TA ASOSIY BO'LIMDAN iborat bo'lsin (Kirish, 3-4 ta asosiy qism, Xulosa).\n"
        "Bo'limlar nomi oddiy va tushunarli bo'lsin. Faqat bo'lim nomlarini yuboring."
    )
    response = await client.chat.completions.create(
        model=OPENAI_MODEL,
        messages=[{"role": "user", "content": prompt}]
    )
    plan_text = response.choices[0].message.content.strip()
    return {"plan": plan_text}

async def generate_document_section(topic: str, section_title: str, extra_details: str = "", language: str = "uz", quality: str = "standard", service_type: str = "") -> str:
    is_pro = (quality == "pro")
    
    style_instruction = (
        "Siz professional professor va akademik yozuvchisiz. "
        "Matnni juda yuqori ilmiy saviyada yozing, tahliliy yondashuvdan foydalaning. "
        "Matn ichida ilmiy manbalarga havolalar (masalan: [1], [3]) ishlating. "
        "Har bir fikrni faktlar va chuqur mantiq bilan isbotlang."
        if is_pro else 
        "Siz tajribali talabasiz. Matn tushunarli, akademik va mantiqiy bo'lsin."
    )
        
    prompt = (
        f"Mavzu: {topic}\n"
        f"Bo'lim nomi: {section_title}\n"
        f"Talablar: {extra_details}\n"
        f"Til: {language}\n\n"
        f"USLUB: {style_instruction}\n\n"
        "VAZIFA: Ushbu bo'lim uchun akademik matn yozing.\n"
        "DIQQAT: 'Talablar' qismidagi so'zlar soniga qat'iy amal qiling.\n\n"
        "QATIY QOIDALAR:\n"
        "1. Matn ichida HECH QACHON raqamli ro'yxatlar (1), 2), a), b) va hokazo) ISHLATMA! "
        "Matn faqat oddiy abzatslardan iborat bo'lsin. Ichki kichik sarlavhalar (1.1.1, 2.1.1) ham QO'SHMA.\n"
        "2. Bo'limni HECH QACHON bo'sh qoldirma. Har bir bo'limda kamida 3 abzats matn bo'lishi SHART.\n"
        "3. Faqat matnning o'zini yuboring, qo'shimcha izoh yoki sarlavha qo'shma."
    )
    response = await client.chat.completions.create(
        model=OPENAI_MODEL,
        messages=[{"role": "system", "content": style_instruction},
                  {"role": "user", "content": prompt}]
    )
    return response.choices[0].message.content.strip()

async def generate_presentation_plan(topic, slides, language="uz"):
    prompt = f"Mavzu: {topic}, Slaydlar soni: {slides}, Til: {language}. Taqdimot rejasini tuzing."
    response = await client.chat.completions.create(model=OPENAI_MODEL, messages=[{"role": "user", "content": prompt}])
    return response.choices[0].message.content.strip()


async def generate_presentation_content(topic, language="uz", num_slides=10, style="standard", quality="standard", num_chapters=4, forced_titles=None, progress_callback=None):
    is_premium = (quality == "premium")
    kb_context = ""
    slides_data = []

    total_steps = num_slides + 2  

    if progress_callback:
        try:
            await progress_callback(0, total_steps, "Mavzuni chuqur tahlil qilish (Mini-tadqiqot)...")
        except Exception:
            pass

    # ==========================================
    # ETAP 1: ANALIZ TEMI (DEEP RESEARCH)
    # ==========================================
    research_text = ""
    if is_premium:
        analysis_prompt = (
            f"{kb_context}\n"
            f"Mavzu: {topic}\nTil: {language}\n\n"
            "VAZIFA: Ushbu mavzu bo'yicha juda chuqur akademik va analitik tadqiqot (mini-research) o'tkazing.\n"
            "Ushbu tadqiqot keyinchalik yuqori sifatli taqdimot tayyorlash uchun asos (poydevor) bo'ladi.\n\n"
            "Tadqiqot quyidagilarni o'z ichiga olishi shart:\n"
            "1. Tarixiy va nazariy kontekst\n"
            "2. Eng muhim faktlar, sanalar va statistik ma'lumotlar\n"
            "3. Amaliy misollar va global / mahalliy bozor yoki jamiyatga ta'siri\n"
            "4. Turli xil qarashlar yoki muammoli jihatlar\n\n"
            "Kamida 600-800 so'zdan iborat to'liq matn yozing."
        )
        try:
            research_text = await _call_ai(
                [{"role": "user", "content": analysis_prompt}],
                max_tokens=3000, temperature=0.8
            )
        except Exception as e:
            logger.error(f"Research failed: {e}")
            research_text = "Tadqiqot ma'lumotlari topilmadi."
    else:
        research_text = "Asosiy faktlarga tayangan holda standart ma'lumotlar."

    if progress_callback:
        try:
            await progress_callback(1, total_steps, "Struktura va slaydlar rejasini tuzish...")
        except Exception:
            pass

    # ==========================================
    # ETAP 2: SOSTAVLENIE PLANA
    # ==========================================
    if forced_titles:
        chapters = []
        for t in forced_titles:
            clean_t = t.split('(')[0].strip()
            if clean_t not in chapters:
                chapters.append(clean_t)
        content_titles = forced_titles
    else:
        plan_sys = (
            "Siz professional metodist va akademik tuzuvchisiz. "
            "Taqdimot uchun eng mantiqiy va izchil strukturani ishlab chiqasiz."
        )
        plan_prompt = (
            f"Mavzu: {topic}\nTil: {language}\n\n"
            f"Tadqiqot bazasi:\n{research_text[:1500]}\n\n"
            f"Ushbu baza asosida {num_chapters} ta mantiqiy va kuchli bo'lim nomini tuzing.\n"
            "FAQAT nomlarni vergul bilan ajratib yozing, boshqa hech narsa qo'shmang."
        )
        try:
            plan_text = await _call_ai(
                [{"role": "system", "content": plan_sys}, {"role": "user", "content": plan_prompt}],
                max_tokens=300, temperature=0.7
            )
            chapters = [c.strip().strip('0123456789.-) ') for c in plan_text.split(',') if c.strip()]
            chapters = [c for c in chapters if len(c) > 2][:num_chapters]
        except Exception as e:
            logger.error(f"Plan generation failed: {e}")
            chapters = ["Kirish", "Asosiy qism", "Tahlil", "Xulosa"]

        if not chapters:
            chapters = ["Kirish", "Asosiy qism", "Xulosa"]

        content_slides_needed = num_slides - 2
        content_titles = []
        if content_slides_needed > 0 and chapters:
            base_count = content_slides_needed // len(chapters)
            remainder = content_slides_needed % len(chapters)
            for i, chapter in enumerate(chapters):
                count_for_this = base_count + (1 if i < remainder else 0)
                if count_for_this == 1:
                    content_titles.append(chapter)
                else:
                    for part in range(1, count_for_this + 1):
                        content_titles.append(f"{chapter} ({part}-qism)")
        else:
            content_titles = chapters[:max(0, content_slides_needed)]

    all_titles = ["Reja", "Kirish"] + content_titles + ["Xulosa"]

    # ==========================================
    # ETAP 3 & 4: KONTENT VA ADAPTATSIYA (SLAYD GENERATION)
    # ==========================================
    
    advanced_designer_rules = (
        "SEN PROFESSIONAL DIZAYNERSAN! Quyidagi 12 ta qoidaga QAT'IY amal qil:\n\n"
        "1. ROL: Sen kontent-meyker emas, DIZAYNERSAN. Ma'lumotni slaydlarga chiroyli taqsimlaysan.\n"
        "2. SLAYD TURLARI: Har xil turlardan foydalan — ro'yxat, taqqoslash, 1 ta yirik g'oya, xronologiya, ikki ustun.\n"
        "3. BLOKLAR: Slayddagi har bir matn blokini alohida fizik zona sifatida qabul qil.\n"
        "4. LIMITLAR: Senga berilgan so'zlar limitiga QAT'IY amal qil — ular fizik o'lchamga asoslangan.\n"
        "   - Kichik blok (15 gacha so'z) → qisqa tezis.\n"
        "   - O'rta blok (15-30 so'z) → 1-2 ta fikr.\n"
        "   - Katta blok (30-50 so'z) → batafsil tushuntirish.\n"
        "5. QISQARTIRISH: Matn sig'masa — qisqartir, soddalashtir, lekin ma'noni 100% saqla.\n"
        "6. AJRATISH: 1 ta slayd = 1 ta asosiy g'oya. Agar sig'masa — keyingi slaydga o'tkaz.\n"
        "7. MANTIQ: Kirish → tushuntirish → detallar → misollar → xulosa ketma-ketligida yoz.\n"
        "8. RITM: Hech qachon bir xil turdagi slaydlarni ketma-ket ishlatma!\n"
        "9. QAROR: Har bir slayd uchun — avval ma'noni tahlil qil, keyin turini tanla, keyin limitni tekshir.\n"
        "10. TEKSHIRUV: Yuklanmagan slayd bormi? Bo'sh slayd bormi? Limit buzilganmi?\n"
        "11. TAQIQ: Uzun abzatslar YOZMA. 1.1, 1.2 kabi raqamlash ISHLATMA. So'z limitini OSHIRMA.\n"
        "12. NATIJA: Har bir punkt MUSTAQIL, O'QILISHI OSON va MA'LUMOTGA BOY bo'lishi shart.\n"
    )

    if is_premium:
        system_persona = (
            f"{kb_context}\n"
            "Siz dunyo miqyosidagi akademik professor, tahlilchi va dizaynersiz.\n"
            f"{advanced_designer_rules}"
        )
        point_depth = (
            "Matnni slayd turiga qarab mantiqiy qismlarga (points) bo'ling. "
            "Punktlar soni (1 tadan 4 tagacha) va uzunligini o'zingiz hal qiling, "
            "lekin har bir punkt o'ta boy, aniq faktlar va raqamlar bilan to'ldirilgan bo'lsin."
        )
        max_tok = 3000
        temp = 0.9
    else:
        system_persona = (
            f"{kb_context}\n"
            "Siz malakali o'qituvchi, dizayner va kontent-meykersiz.\n"
            f"{advanced_designer_rules}"
        )
        point_depth = (
            "Matnni slayd turiga qarab mantiqiy qismlarga (points) bo'ling. "
            "Punktlar soni (1 tadan 4 tagacha) va uzunligini o'zingiz hal qiling. "
            "Faktlar va misollar bilan boyiting."
        )
        max_tok = 2500
        temp = 0.8

    completed_steps = 1
    for i, title in enumerate(all_titles):
        # Reja slide
        if title == "Reja":
            points = [f"{idx+1}. {c}" for idx, c in enumerate(chapters)]
            slides_data.append({"title": "Reja", "points": points})
            completed_steps += 1
            if progress_callback:
                next_title = all_titles[i + 1] if i + 1 < len(all_titles) else "Fizikaviy tekshiruv"
                try:
                    await progress_callback(completed_steps, total_steps, next_title)
                except Exception:
                    pass
            continue

        # Kirish slide
        if title == "Kirish":
            k_prompt = (
                f"Mavzu: {topic}\nTil: {language}\n"
                f"Tadqiqot bazasi: {research_text[:1000]}...\n\n"
                "Ushbu taqdimot uchun qisqa va qiziqarli 'Kirish' tayyorlang.\n"
                "Mavzuning dolzarbligi va ahamiyatini tushuntiring.\n"
                f"QOIDALAR: {point_depth}\n"
                "- FAQAT JSON formatida javob bering:\n"
                "{\"title\": \"Kirish\", \"points\": [\"1-qism...\", \"2-qism...\", \"...\"]}"
            )
            try:
                raw = await _call_ai(
                    [{"role": "system", "content": system_persona}, {"role": "user", "content": k_prompt}],
                    max_tokens=max_tok, temperature=temp, json_mode=True
                )
                import json
                data = json.loads(raw)
                slides_data.append({"title": "Kirish", "points": data.get("points", ["Kirish ma'lumotlari shakllantirilmoqda."])})
            except Exception as e:
                logger.error(f"Kirish generation error: {e}")
                slides_data.append({"title": "Kirish", "points": ["Mavzuning dolzarbligi", "Asosiy maqsad va vazifalar", "Kutilayotgan natijalar"]})
            
            completed_steps += 1
            if progress_callback:
                next_title = all_titles[i + 1] if i + 1 < len(all_titles) else "Xulosa"
                try:
                    await progress_callback(completed_steps, total_steps, next_title)
                except Exception:
                    pass
            continue

        # Xulosa slide
        if title == "Xulosa":
            x_rule = (
                "Xulosa juda chuqur bo'lsin. Kamida 3 ta yirik punkt: "
                "(1) Asosiy topilmalarni jamlash, (2) Amaliy va nazariy ahamiyati, (3) Kelajakdagi istiqbollar. "
                "Har bir punkt kamida 60 so'z."
            ) if is_premium else "Ikki yirik punkt: (1) asosiy natijalar va (2) ahamiyat. Har biri 40 so'z."
            
            x_prompt = (
                f"Mavzu: {topic}\nTil: {language}\n"
                f"Tadqiqot bazasi: {research_text[:1000]}...\n\n"
                "Ushbu taqdimot uchun chuqur Xulosa tayyorlang.\n"
                f"QOIDALAR: {x_rule}\n"
                "- FAQAT JSON formatida javob bering:\n"
                "{\"title\": \"Xulosa\", \"points\": [\"1-qism...\", \"2-qism...\", \"...\"]}"
            )
            try:
                raw = await _call_ai(
                    [{"role": "system", "content": system_persona}, {"role": "user", "content": x_prompt}],
                    max_tokens=max_tok, temperature=temp, json_mode=True
                )
                data = json.loads(raw)
                data["title"] = "Xulosa"
                slides_data.append(data)
            except Exception as e:
                logger.error(f"Xulosa generation failed: {e}")
                slides_data.append({"title": "Xulosa", "points": ["Xulosa yuzaga keldi."]})
            
            completed_steps += 1
            if progress_callback:
                try:
                    await progress_callback(completed_steps, total_steps, "Sifat nazorati va adaptatsiya")
                except Exception:
                    pass
            continue

        # Regular slide
        prompt = (
            f"Mavzu: {topic}\nSlayd sarlavhasi: {title}\nTil: {language}\n\n"
            f"ASOSIY TADQIQOT BAZASI (ushbu ma'lumotlarga tayaning):\n"
            f"{research_text}\n\n"
            f"VAZIFA: Faqat '{title}' mavzusiga bag'ishlangan bitta slayd uchun maksimal sifatli va chuqur matn yozing.\n\n"
            "ETAP 3: KONTENT YARATISH\n"
            f"{point_depth}\n\n"
            "ETAP 4: SHABLONGA ADAPTATSIYA\n"
            "Ushbu punktlar dizayn shabloniga avtomatik tushadi. Shuning uchun ular mantiqan mustaqil, "
            "o'qilishi oson va juda boy ma'lumotli bo'lishi kerak.\n\n"
            "- Matn ichida qo'shtirnoq (\") ISHLATMANG! Yakkalik tirnoq (') ishlating.\n"
            "- FAQAT JSON formatida javob bering:\n"
            "{\"title\": \"sarlavha\", \"points\": [\"1-qism...\", \"2-qism...\", \"...\"]}"
        )
        try:
            raw = await _call_ai(
                [{"role": "system", "content": system_persona}, {"role": "user", "content": prompt}],
                max_tokens=max_tok, temperature=temp, json_mode=True
            )
            data = json.loads(raw)
            data["title"] = title
            slides_data.append(data)
        except Exception as e:
            logger.error(f"Slide '{title}' generation failed after retries: {e}")
            slides_data.append({"title": title, "points": [f"'{title}' mavzusi bo'yicha ma'lumot."]})
        
        completed_steps += 1
        if progress_callback:
            next_title = all_titles[i + 1] if i + 1 < len(all_titles) else "Fizikaviy tekshiruv"
            try:
                await progress_callback(completed_steps, total_steps, next_title)
            except Exception:
                pass

    # ==========================================
    # ETAP 5: FINAL TEKSHIRUV (VALIDATION)
    # ==========================================
    for slide in slides_data:
        if "points" not in slide or not isinstance(slide["points"], list) or len(slide["points"]) == 0:
            slide["points"] = ["Ma'lumotlar qayta ishlanmoqda..."]
            
    completed_steps += 1
    if progress_callback:
        try:
            await progress_callback(completed_steps, total_steps, "")
        except Exception:
            pass

    return json.dumps({"slides": slides_data}, ensure_ascii=False)


async def generate_akademik_content(
    topic: str,
    language: str = "uz",
    subject_name: str = "",
    completed_by: str = "",
    training_session: str = "",
    quality: str = "standard",
    progress_callback=None,
) -> dict:
    """
    Generate structured content for akademik template.
    Returns a dict with all required tag values:
      TOPIC, SUBJECT_NAME, COMPLETED_BY, TRAINING_SESSION,
      EDUCATIONAL_GOALS, STUDY_QUESTIONS,
      QUESTION_1, QUESTION_1_CONTENT, QUESTION_1_CONCLUSION,
      QUESTION_2, QUESTION_2_CONTENT, QUESTION_2_CONCLUSION,
      QUESTION_3, QUESTION_3_CONTENT,
      GENERAL_CONCLUSION, REFERENCES_LIST
    """
    lang_name = {"uz": "o'zbek", "ru": "русский", "en": "English"}.get(language, "o'zbek")

    total_steps = 4
    completed = 0

    # Step 1: Generate the 3 questions and educational goals
    if progress_callback:
        try: await progress_callback(completed, total_steps, "Mavzuni tahlil qilish...")
        except: pass

    system_prompt = f"""Sen akademik professor va pedagogsan. Vazifang — berilgan mavzu bo'yicha ta'lim mashg'ulotining to'liq mazmunini tayyorlash.
Til: {lang_name}. 

FAQAT JSON formatida javob ber! Hech qanday izoh, tushuntirish yoki boshqa matn yozma!

JSON strukturasi:
{{
  "EDUCATIONAL_GOALS": "Mashg'ulotning ta'limiy maqsadlari (3-4 ta aniq maqsad, nuqtali vergul bilan ajratilgan)",
  "STUDY_QUESTIONS": "3 ta o'quv savolining nomlari (raqamlangan ro'yxat sifatida)",
  "QUESTION_1": "Birinchi o'quv savolining sarlavhasi",
  "QUESTION_1_CONTENT": "Birinchi o'quv savoli bo'yicha to'liq, chuqur ilmiy matn. Kamida 200 so'z. Akademik uslubda, dalillar va misollar bilan.",
  "QUESTION_1_CONCLUSION": "Birinchi savolga xulosa (2-3 jumlada)",
  "QUESTION_2": "Ikkinchi o'quv savolining sarlavhasi",
  "QUESTION_2_CONTENT": "Ikkinchi o'quv savoli bo'yicha to'liq, chuqur ilmiy matn. Kamida 200 so'z.",
  "QUESTION_2_CONCLUSION": "Ikkinchi savolga xulosa (2-3 jumlada)",
  "QUESTION_3": "Uchinchi o'quv savolining sarlavhasi",
  "QUESTION_3_CONTENT": "Uchinchi o'quv savoli bo'yicha to'liq, chuqur ilmiy matn. Kamida 200 so'z.",
  "GENERAL_CONCLUSION": "Umumiy xulosa — mavzuni yakunlovchi chuqur tahlil (4-5 jumla)",
  "REFERENCES_LIST": "5-7 ta akademik manba (kitoblar, maqolalar). Har birini yangi qatordan yozing."
}}

QOIDALAR:
1. Har bir CONTENT maydoni kamida 200 so'z bo'lishi SHART
2. Akademik, ilmiy uslubda yozing
3. Har bir javob faqat shu mavzuga tegishli bo'lsin
4. Faqat JSON qaytar, boshqa hech narsa yozma"""

    user_prompt = f"Mavzu: {topic}"

    try:
        response = await _call_ai(
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            max_tokens=4000,
            temperature=0.7,
            json_mode=True
        )
        
        completed = 2
        if progress_callback:
            try: await progress_callback(completed, total_steps, "Ma'lumotlar tahlil qilinmoqda...")
            except: pass

        # Parse JSON response
        clean = response.replace("```json", "").replace("```", "").strip()
        content = json.loads(clean)

    except Exception as e:
        logger.error(f"Akademik AI generation failed: {e}")
        # Fallback content
        content = {
            "EDUCATIONAL_GOALS": f"{topic} mavzusini o'rganish va tahlil qilish",
            "STUDY_QUESTIONS": f"1. {topic} asoslari\n2. {topic} amaliyoti\n3. {topic} istiqbollari",
            "QUESTION_1": f"{topic} asoslari",
            "QUESTION_1_CONTENT": f"{topic} mavzusining nazariy asoslari...",
            "QUESTION_1_CONCLUSION": f"{topic} nazariy jihatdan muhim.",
            "QUESTION_2": f"{topic} amaliyoti",
            "QUESTION_2_CONTENT": f"{topic} mavzusining amaliy qo'llanilishi...",
            "QUESTION_2_CONCLUSION": f"{topic} amaliyotda keng qo'llaniladi.",
            "QUESTION_3": f"{topic} istiqbollari",
            "QUESTION_3_CONTENT": f"{topic} mavzusining kelajakdagi rivojlanish yo'nalishlari...",
            "GENERAL_CONCLUSION": f"{topic} mavzusi bo'yicha umumiy xulosa.",
            "REFERENCES_LIST": "1. O'zbekiston Milliy Ensiklopediyasi\n2. Akademik tadqiqotlar jurnali"
        }

    # Add metadata fields
    content["TOPIC"] = topic
    content["SUBJECT_NAME"] = subject_name or topic
    content["COMPLETED_BY"] = completed_by or ""
    content["TRAINING_SESSION"] = training_session or "Ma'ruza mashg'uloti"

    completed = 3
    if progress_callback:
        try: await progress_callback(completed, total_steps, "Yakunlanmoqda...")
        except: pass

    completed = 4
    if progress_callback:
        try: await progress_callback(completed, total_steps, "")
        except: pass

    return content
