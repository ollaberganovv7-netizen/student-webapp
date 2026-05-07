from aiogram import Router, F, Bot
from aiogram.filters import Command
from aiogram.types import CallbackQuery, Message
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup

from database.db import (
    get_pending_payments, get_payment, update_payment_status,
    add_balance, get_user_count, get_premium_count,
    get_request_count, get_payment_count, get_recent_requests, get_all_users,
)
from keyboards.admin_kb import admin_main_kb, payment_action_kb, admin_back_kb
from keyboards.main_kb import main_menu_kb
from config import ADMIN_IDS, PRICING

router = Router()


def is_admin(user_id: int) -> bool:
    return user_id in ADMIN_IDS


class AdminStates(StatesGroup):
    broadcasting = State()
    adding_balance_id = State()
    adding_balance_amount = State()


# ─── Admin guard filter ───────────────────────────────────────────────────────

@router.message(Command("admin"))
async def cmd_admin(message: Message):
    if not is_admin(message.from_user.id):
        await message.answer("⛔ Sizda admin huquqlari yo'q.")
        return
    await message.answer(
        "🛡️ <b>Admin panel</b>\n\nXizmatni tanlang:",
        reply_markup=admin_main_kb(),
        parse_mode="HTML",
    )


@router.callback_query(F.data == "admin:back")
async def admin_back(callback: CallbackQuery):
    if not is_admin(callback.from_user.id):
        await callback.answer("⛔ Ruxsat yo'q", show_alert=True)
        return
    text = "🛡️ <b>Admin panel</b>\n\nXizmatni tanlang:"
    markup = admin_main_kb()

    try:
        if callback.message.photo or callback.message.document:
            await callback.message.delete()
            await callback.message.answer(text, reply_markup=markup, parse_mode="HTML")
        else:
            await callback.message.edit_text(text, reply_markup=markup, parse_mode="HTML")
    except Exception as e:
        # Fallback if edit fails
        await callback.message.delete()
        await callback.message.answer(text, reply_markup=markup, parse_mode="HTML")
        
    await callback.answer()


# ─── Pending Payments ─────────────────────────────────────────────────────────

@router.callback_query(F.data == "admin:pending_payments")
async def admin_pending_payments(callback: CallbackQuery, bot: Bot):
    if not is_admin(callback.from_user.id):
        await callback.answer("⛔ Ruxsat yo'q", show_alert=True)
        return

    payments = await get_pending_payments()

    if not payments:
        await callback.message.edit_text(
            "✅ <b>Kutayotgan to'lovlar yo'q.</b>",
            reply_markup=admin_back_kb(),
            parse_mode="HTML",
        )
        await callback.answer()
        return

    await callback.message.edit_text(
        f"💳 <b>{len(payments)} ta kutayotgan to'lov:</b>",
        reply_markup=admin_back_kb(),
        parse_mode="HTML",
    )

    for payment in payments:
        try:
            await bot.send_photo(
                chat_id=callback.from_user.id,
                photo=payment.screenshot_file_id,
                caption=(
                    f"🔑 To'lov ID: #{payment.id}\n"
                    f"👤 Foydalanuvchi ID: <code>{payment.user_id}</code>\n"
                    f"📦 Paket: {payment.package}\n"
                    f"💰 Summa: {payment.amount:,} so'm\n"
                    f"📅 Sana: {payment.created_at.strftime('%d.%m.%Y %H:%M')}"
                ),
                reply_markup=payment_action_kb(payment.id, payment.user_id),
                parse_mode="HTML",
            )
        except Exception:
            pass

    await callback.answer()


# ─── Approve payment ──────────────────────────────────────────────────────────

@router.callback_query(F.data.startswith("admin:approve:"))
async def admin_approve_payment(callback: CallbackQuery, bot: Bot):
    if not is_admin(callback.from_user.id):
        await callback.answer("⛔ Ruxsat yo'q", show_alert=True)
        return

    parts = callback.data.split(":")
    payment_id = int(parts[2])
    user_id = int(parts[3])

    payment = await get_payment(payment_id)
    if not payment:
        await callback.answer("To'lov topilmadi", show_alert=True)
        return

    if payment.status != "pending":
        await callback.answer(f"Bu to'lov allaqachon: {payment.status}", show_alert=True)
        return

    await update_payment_status(payment_id, "approved", "Admin tomonidan tasdiqlandi")
    await add_balance(user_id, payment.amount)

    await callback.message.edit_caption(
        caption=callback.message.caption + "\n\n✅ <b>TASDIQLANDI</b>",
        parse_mode="HTML",
    )
    await callback.answer("✅ To'lov tasdiqlandi!")

    # Notify user
    try:
        await bot.send_message(
            chat_id=user_id,
            text=(
                f"🎉 <b>Tabriklaymiz! To'lovingiz tasdiqlandi.</b>\n\n"
                f"💰 Balansingizga <b>{payment.amount:,} so'm</b> qo'shildi.\n"
                f"Endi xizmatlardan foydalanishingiz mumkin!"
            ),
            reply_markup=main_menu_kb(),
            parse_mode="HTML",
        )
    except Exception:
        pass


# ─── Reject payment ───────────────────────────────────────────────────────────

@router.callback_query(F.data.startswith("admin:reject:"))
async def admin_reject_payment(callback: CallbackQuery, bot: Bot):
    if not is_admin(callback.from_user.id):
        await callback.answer("⛔ Ruxsat yo'q", show_alert=True)
        return

    parts = callback.data.split(":")
    payment_id = int(parts[2])
    user_id = int(parts[3])

    payment = await get_payment(payment_id)
    if not payment:
        await callback.answer("To'lov topilmadi", show_alert=True)
        return

    if payment.status != "pending":
        await callback.answer(f"Bu to'lov allaqachon: {payment.status}", show_alert=True)
        return

    await update_payment_status(payment_id, "rejected", "Admin tomonidan rad etildi")

    await callback.message.edit_caption(
        caption=callback.message.caption + "\n\n❌ <b>RAD ETILDI</b>",
        parse_mode="HTML",
    )
    await callback.answer("❌ To'lov rad etildi!")

    try:
        await bot.send_message(
            chat_id=user_id,
            text=(
                "❌ <b>Afsuski, to'lovingiz tasdiqlanmadi.</b>\n\n"
                "Sabab: Chek aniq emas yoki to'lov summasi noto'g'ri.\n"
                "Qayta to'lov qilib, chek yuboring."
            ),
            reply_markup=main_menu_kb(),
            parse_mode="HTML",
        )
    except Exception:
        pass


# ─── Statistics ───────────────────────────────────────────────────────────────

@router.callback_query(F.data == "admin:stats")
async def admin_stats(callback: CallbackQuery):
    if not is_admin(callback.from_user.id):
        await callback.answer("⛔ Ruxsat yo'q", show_alert=True)
        return

    total_users = await get_user_count()
    total_requests = await get_request_count()
    approved_payments = await get_payment_count()

    text = (
        f"📊 <b>Bot statistikasi</b>\n\n"
        f"👥 Jami foydalanuvchilar: <b>{total_users}</b>\n"
        f"📄 Jami so'rovlar: <b>{total_requests}</b>\n"
        f"✅ Tasdiqlangan to'lovlar: <b>{approved_payments}</b>\n"
    )

    await callback.message.edit_text(text, reply_markup=admin_back_kb(), parse_mode="HTML")
    await callback.answer()


# ─── Pricing list ─────────────────────────────────────────────────────────────

@router.callback_query(F.data == "admin:pricing")
async def admin_pricing(callback: CallbackQuery):
    if not is_admin(callback.from_user.id):
        await callback.answer("⛔ Ruxsat yo'q", show_alert=True)
        return

    lines = [f"💰 <b>Joriy narxlar:</b>\n"]
    for key, price in PRICING.items():
        lines.append(f"• {key}: <b>{price:,} so'm</b>")

    await callback.message.edit_text(
        "\n".join(lines),
        reply_markup=admin_back_kb(),
        parse_mode="HTML",
    )
    await callback.answer()


# ─── Recent requests ──────────────────────────────────────────────────────────

@router.callback_query(F.data == "admin:recent_requests")
async def admin_recent_requests(callback: CallbackQuery):
    if not is_admin(callback.from_user.id):
        await callback.answer("⛔ Ruxsat yo'q", show_alert=True)
        return

    requests = await get_recent_requests(limit=10)
    if not requests:
        await callback.message.edit_text(
            "📋 So'rovlar yo'q.",
            reply_markup=admin_back_kb(),
            parse_mode="HTML",
        )
        await callback.answer()
        return

    lines = ["📋 <b>So'nggi 10 ta so'rov:</b>\n"]
    for req in requests:
        lines.append(
            f"• [{req.created_at.strftime('%d.%m %H:%M')}] "
            f"<code>{req.user_id}</code> → {req.service_type}: <i>{req.topic[:40]}</i>"
        )

    await callback.message.edit_text(
        "\n".join(lines),
        reply_markup=admin_back_kb(),
        parse_mode="HTML",
    )
    await callback.answer()


# ─── Add Balance Manually ───────────────────────────────────────────────────

@router.callback_query(F.data == "admin:add_balance")
async def admin_add_balance_start(callback: CallbackQuery, state: FSMContext):
    if not is_admin(callback.from_user.id):
        await callback.answer("⛔ Ruxsat yo'q", show_alert=True)
        return
    await state.set_state(AdminStates.adding_balance_id)
    await callback.message.edit_text(
        "➕ <b>Balans qo'shish</b>\n\n"
        "👤 Foydalanuvchining ID raqamini yozing:\n"
        "<i>(Bekor qilish uchun /cancel)</i>",
        reply_markup=admin_back_kb(),
        parse_mode="HTML",
    )
    await callback.answer()

@router.message(AdminStates.adding_balance_id)
async def admin_add_balance_get_id(message: Message, state: FSMContext):
    if not is_admin(message.from_user.id):
        return
    if message.text == "/cancel":
        await state.clear()
        await message.answer("❌ Bekor qilindi.", reply_markup=admin_back_kb())
        return

    try:
        user_id = int(message.text.strip())
    except ValueError:
        await message.answer("⚠️ Iltimos, faqat ID raqam yozing.")
        return

    await state.update_data(target_user_id=user_id)
    await state.set_state(AdminStates.adding_balance_amount)
    await message.answer(
        f"✅ ID: <code>{user_id}</code> qabul qilindi.\n\n"
        "💰 Qancha summa qo'shmoqchisiz?\n"
        "<i>(Masalan: 5000 yoki 10000)</i>\n"
        "<i>(Bekor qilish uchun /cancel)</i>",
        parse_mode="HTML",
        reply_markup=admin_back_kb()
    )

@router.message(AdminStates.adding_balance_amount)
async def admin_add_balance_finish(message: Message, state: FSMContext, bot: Bot):
    if not is_admin(message.from_user.id):
        return
    if message.text == "/cancel":
        await state.clear()
        await message.answer("❌ Bekor qilindi.", reply_markup=admin_back_kb())
        return

    try:
        amount = int(message.text.strip().replace(" ", "").replace(",", ""))
    except ValueError:
        await message.answer("⚠️ Iltimos, faqat raqam yozing.")
        return

    data = await state.get_data()
    user_id = data.get("target_user_id")

    try:
        await add_balance(user_id, amount)
        await state.clear()
        await message.answer(
            f"✅ <b>Muvaffaqiyatli!</b>\n\n"
            f"👤 ID: <code>{user_id}</code>\n"
            f"💰 Qo'shildi: {amount:,} so'm",
            parse_mode="HTML",
            reply_markup=admin_back_kb()
        )
        
        # Notify user
        try:
            await bot.send_message(
                chat_id=user_id,
                text=(
                    f"🎉 <b>Tabriklaymiz! Admin tomonidan balansingiz to'ldirildi.</b>\n\n"
                    f"💰 Balansingizga <b>{amount:,} so'm</b> qo'shildi.\n"
                    f"Xizmatlardan foydalanishingiz mumkin!"
                ),
                reply_markup=main_menu_kb(),
                parse_mode="HTML"
            )
        except Exception:
            await message.answer("⚠️ Foydalanuvchiga xabar yuborib bo'lmadi (botni bloklagan bo'lishi mumkin).")

    except Exception as e:
        await message.answer(f"❌ Xatolik yuz berdi: {e}")

# ─── Broadcast ────────────────────────────────────────────────────────────────

@router.callback_query(F.data == "admin:broadcast")
async def admin_broadcast_start(callback: CallbackQuery, state: FSMContext):
    if not is_admin(callback.from_user.id):
        await callback.answer("⛔ Ruxsat yo'q", show_alert=True)
        return

    await state.set_state(AdminStates.broadcasting)
    await callback.message.edit_text(
        "📢 <b>Xabar yuborish</b>\n\n"
        "Barcha foydalanuvchilarga yuboriladigan xabarni yozing:\n"
        "<i>(Bekor qilish uchun /cancel)</i>",
        reply_markup=admin_back_kb(),
        parse_mode="HTML",
    )
    await callback.answer()


@router.message(AdminStates.broadcasting)
async def admin_broadcast_send(message: Message, state: FSMContext, bot: Bot):
    if not is_admin(message.from_user.id):
        return

    if message.text and message.text.strip() == "/cancel":
        await state.clear()
        await message.answer("❌ Xabar bekor qilindi.", reply_markup=admin_back_kb())
        return

    await state.clear()
    users = await get_all_users()
    sent = 0
    failed = 0

    status_msg = await message.answer(f"📤 Yuborilmoqda... 0/{len(users)}")

    for i, user in enumerate(users):
        try:
            await bot.send_message(
                chat_id=user.id,
                text=f"📢 <b>Bot xabari:</b>\n\n{message.text or message.caption or ''}",
                parse_mode="HTML",
            )
            sent += 1
        except Exception:
            failed += 1

        if (i + 1) % 20 == 0:
            try:
                await status_msg.edit_text(f"📤 Yuborilmoqda... {i+1}/{len(users)}")
            except Exception:
                pass

    await status_msg.edit_text(
        f"✅ <b>Xabar yuborildi!</b>\n\n"
        f"✅ Muvaffaqiyatli: {sent}\n"
        f"❌ Xatolik: {failed}",
        parse_mode="HTML",
    )
