#!/usr/bin/env python3
"""Отправка рассылки с личного аккаунта через Telethon."""

import argparse
import asyncio
import os
from pathlib import Path
from typing import Iterable

from telethon import TelegramClient
from telethon.tl.functions.contacts import ImportContactsRequest
from telethon.tl.types import InputPhoneContact


def parse_targets(raw: str | None) -> list[str]:
    if not raw:
        return []
    parts: list[str] = []
    for chunk in raw.replace(";", ",").split(","):
        value = chunk.strip()
        if value:
            parts.append(value)
    return parts


async def ensure_phone_contacts(client: TelegramClient, targets: list[str]) -> dict[str, str]:
    """Добавляем телефоны в контакты и возвращаем карту phone -> peer id для отправки."""
    phones = [t for t in targets if t.startswith("+")]
    if not phones:
        return {}
    contacts = [
        InputPhoneContact(client_id=i, phone=phone, first_name=f"Broadcast{i}", last_name="")
        for i, phone in enumerate(phones)
    ]
    result = await client(ImportContactsRequest(contacts))
    mapping: dict[str, str] = {}
    for imported, phone in zip(result.imported, phones):
        if imported.user_id:
            mapping[phone] = imported.user_id
    return mapping


async def broadcast(client: TelegramClient, targets: Iterable[str], text: str, file_path: Path | None):
    targets = list(targets)
    phone_map = await ensure_phone_contacts(client, targets)
    for target in targets:
        peer = phone_map.get(target) if target.startswith("+") else target
        try:
            if file_path:
                await client.send_file(peer, file_path, caption=text)
            else:
                await client.send_message(peer, text)
            print(f"✅ Sent to {target}")
        except Exception as exc:
            print(f"❌ Failed to send to {target}: {exc}")


async def main():
    parser = argparse.ArgumentParser(description="Broadcast message from user account via Telethon")
    parser.add_argument("--text", required=True, help="Текст сообщения")
    parser.add_argument("--targets", help="Список получателей через запятую (@user, chat_id, phone)")
    parser.add_argument("--file", type=Path, help="Опциональный путь к файлу для отправки")
    parser.add_argument("--api-id", type=int, default=int(os.environ.get("TELETHON_API_ID", "0")), help="api_id из my.telegram.org")
    parser.add_argument("--api-hash", default=os.environ.get("TELETHON_API_HASH", ""), help="api_hash из my.telegram.org")
    parser.add_argument("--phone", default=os.environ.get("TELETHON_PHONE", ""), help="Номер аккаунта в формате +7...")
    args = parser.parse_args()

    targets = parse_targets(args.targets) or parse_targets(os.environ.get("TELETHON_TARGETS"))
    if not targets:
        raise SystemExit("Укажи получателей в --targets или TELETHON_TARGETS")
    if not args.api_id or not args.api_hash or not args.phone:
        raise SystemExit("Нужны api_id, api_hash и phone (аргументы или переменные TELETHON_API_ID/TELETHON_API_HASH/TELETHON_PHONE)")

    session_path = Path(os.environ.get("TELETHON_SESSION", "userbot_session"))
    async with TelegramClient(session_path, args.api_id, args.api_hash) as client:
        # при первом запуске спросит код / пароль
        await client.start(phone=args.phone)
        await broadcast(client, targets, args.text, args.file)


if __name__ == "__main__":
    asyncio.run(main())
