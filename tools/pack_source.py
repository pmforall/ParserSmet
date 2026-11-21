#!/usr/bin/env python3
"""Снимок исходников без смет, отчётов, секретов и эксплуатационных журналов.

python tools/pack_source.py --out ../releases/smeta-parser-source.zip
Публикация сайта и конфигурация служб не изменяются.
"""
from __future__ import annotations

import argparse
from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path, PurePosixPath
import re
import stat
import subprocess
import sys
import tempfile
import zipfile

ROOT = Path(__file__).resolve().parents[1]
FILES = (
    ".gitignore", "CHANGELOG.md", "build.py", "run_all.py",
    "app/schema.json", "app/template.html",
    "parser/smeta_parser.py", "parser/printed_form.py", "parser/printed_positions.py",
    "parser/reference.py", "parser/verify_parser.py",
    "tools/package_web.py", "tools/pack_source.py",
    "tests/requirements.txt", "tests/excel_contract.json", "tests/dom_stub.js",
    "tests/make_multi_ls.py", "tests/make_no_structure.py",
    "tests/test_printed_positions.py", "tests/test_row_groups.py",
    "tests/test_web_build.py", "tests/test_multi_ls.py",
    "tests/test_browser_contract.py", "tests/test_portable_browser.py", "tests/test_source_archive.py",
    "tests/test_sso_gateway.py", "tests/test_box1_auth.py", "tests/test_idp_registration.py",
    "deploy/prepare_sso.py", "deploy/stage_sso.py", "deploy/install_standalone.py",
    "deploy/register_idp_client.py", "deploy/call_registration.py",
    "deploy/activate_standalone.py", "deploy/update_web.py",
    "deploy/box1_auth.py", "deploy/activate_box1.py",
)
BLOCKED_PARTS = {".git", ".venv", "node_modules", "__pycache__", "prepared", "out",
                 "data", "example", "Исходные данные", "input", "backups", "artifacts"}
BLOCKED_SUFFIXES = {".xlsx", ".xls", ".xlsm", ".log", ".pem", ".key", ".pfx", ".p12",
                    ".jks", ".db", ".sqlite", ".sqlite3", ".zip", ".gz", ".tar", ".pyc"}
SECRET_RULES = {
    "private-key": rb"-----BEGIN (?:RSA |EC |DSA |OPENSSH |ENCRYPTED )?PRIVATE KEY-----",
    "jwt": rb"eyJ[A-Za-z0-9_-]{15,}\.[A-Za-z0-9_-]{15,}\.[A-Za-z0-9_-]{12,}",
    "credential-url": rb"(?:https?|ssh)://[^\s/:]+:[^\s/@]+@",
    "provider-token": rb"(?:gh[pousr]_[A-Za-z0-9]{30,}|AKIA[0-9A-Z]{16}|sk-[A-Za-z0-9_-]{25,})",
    "email-address": rb"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}",
    "credential-literal": rb'''(?im)\b(?:password|passwd|client_secret|api_key|access_token|refresh_token)\s*[:=]\s*["']([^"'\r\n]{8,})["']''',
}
PRIVACY = """# Состав и исключение приватных данных

В архив включены только явно перечисленные исходники, документы проекта,
актуальные тесты и заново собранные HTML с пустыми данными.

Не включены исходные сметы, справочники, разобранные данные, Excel-отчёты,
виртуальное окружение и зависимости, кэши, история Git, журналы запусков,
резервные копии, активные конфигурации и квитанции SSO, данные сессий,
client secret, ключи хранилища, сертификаты с закрытыми ключами, пароли и токены.
Исторические тесты, содержащие финансовые ожидания реальных смет, исключены.

Ссылки на реальные тестовые книги в копии исходников заменены обозначениями
fixture-NNN. Сами книги не читались и не включались. Адреса сервисов и пути
к внешним файлам секретов в коде интеграции сохранены как параметры окружения;
значений секретов в комплекте нет. Синтетические маркеры и генерация временных
ключей в тестовом коде относятся только к изолированным тестовым стендам.

Перед упаковкой проверяются допустимые пути, типы файлов, отсутствие ссылок,
материала закрытых ключей, JWT, типовых токенов, адресов электронной почты,
URL с паролями и литералов учётных данных. Состав и хеши указаны в MANIFEST.json.
"""


def allowed_name(name):
    path = PurePosixPath(name)
    return (not path.is_absolute() and ".." not in path.parts
            and not BLOCKED_PARTS.intersection(path.parts)
            and path.suffix.lower() not in BLOCKED_SUFFIXES
            and not any(p.startswith(".env") or p in ("client-secret", "store-key") for p in path.parts))


def scan(name, data):
    if not allowed_name(name):
        raise ValueError("запрещённый путь: " + name)
    for label, pattern in SECRET_RULES.items():
        if re.search(pattern, data):
            # Содержимое совпадения не выводится: оно может быть секретным.
            raise ValueError("проверка приватности: " + name + " (" + label + ")")


def sources(root):
    result = {name: root / name for name in FILES}
    for name, preferred in (("README.md", "docs/SOURCE-README.md"),
                            ("deploy/README.md", "docs/SOURCE-DEPLOY.md")):
        result[name] = root / preferred if (root / preferred).is_file() else root / name
    for name, path in result.items():
        if not allowed_name(name) or path.is_symlink() or not path.is_file():
            raise ValueError("небезопасный или отсутствующий исходник: " + name)
        if any(parent.is_symlink() for parent in path.parents if parent != root.parent):
            raise ValueError("символическая ссылка в пути исходника: " + name)
    return result


def book_aliases(root):
    names = set()
    for folder in ("Исходные данные", "example", "input"):
        directory = root / folder
        if directory.is_symlink():
            raise ValueError("символическая ссылка на каталог данных")
        if directory.is_dir():
            names.update(p.name for p in directory.glob("*.xlsx") if not p.is_symlink())
    result = {}
    for i, name in enumerate(sorted(names), 1):
        # Это стандартное имя формата справочника, а не идентификатор сметы.
        if name == "Справочник новый.xlsx":
            continue
        result[name] = f"fixture-{i:03}.xlsx"
        result[Path(name).stem] = f"fixture-{i:03}"
    return result


def neutral_text(data, aliases):
    value = data.decode("utf-8")
    for original in sorted(aliases, key=len, reverse=True):
        value = value.replace(original, aliases[original])
    value = re.sub(r"\b24-0\d{4}\b", "SAMPLE", value)
    value = value.replace("типовых книг", "типовых книг")
    return value.encode("utf-8")


def build_archive(root, output):
    root, output = root.resolve(), output.absolute()
    resolved_output = output.resolve()
    if resolved_output.is_relative_to(root):
        raise ValueError("архив передачи необходимо сохранять вне каталога проекта")
    if output.is_symlink() or output.exists():
        raise ValueError("файл архива уже существует: " + str(output))
    output.parent.mkdir(parents=True, exist_ok=True)
    selected, aliases = sources(root), book_aliases(root)
    now = datetime.now(timezone.utc)
    with tempfile.TemporaryDirectory(prefix=".smeta-source-", dir=output.parent) as temporary:
        stage = Path(temporary) / "smeta-parser"
        stage.mkdir()
        normalized = []
        for name, path in selected.items():
            original = path.read_bytes()
            data = neutral_text(original, aliases)
            scan(name, data)
            destination = stage / name
            destination.parent.mkdir(parents=True, exist_ok=True)
            destination.write_bytes(data)
            if data != original:
                normalized.append(name)
        (stage / "PRIVACY.md").write_text(PRIVACY, encoding="utf-8")
        for args in (("build.py",), ("tools/package_web.py", "--check")):
            subprocess.run([sys.executable, "-B", *(str(stage / a) if a.endswith(".py") else a for a in args)],
                           cwd=stage, check=True, capture_output=True, text=True)
        members = {}
        for path in sorted(stage.rglob("*")):
            if path.is_dir():
                continue
            name = path.relative_to(stage).as_posix()
            if not stat.S_ISREG(path.lstat().st_mode):
                raise ValueError("недопустимый тип файла: " + name)
            data = path.read_bytes()
            scan(name, data)
            members[name] = data
        manifest = {
            "schema": "smeta.source-archive/v1", "created_at": now.isoformat(),
            "type": "sanitized-source-snapshot", "git_history_included": False,
            "private_data_included": False, "live_site_modified": False,
            "normalized_source_files": sorted(normalized),
            "files": {name: {"bytes": len(data), "sha256": hashlib.sha256(data).hexdigest()}
                      for name, data in members.items()},
            "manifest_note": "MANIFEST.json не включает собственный хеш; ZIP имеет внешнюю SHA-256.",
        }
        members["MANIFEST.json"] = (json.dumps(manifest, ensure_ascii=False, indent=2) + "\n").encode()
        target = Path(temporary) / "source.zip"
        with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
            for name, data in sorted(members.items()):
                scan(name, data)
                info = zipfile.ZipInfo("smeta-parser/" + name, now.timetuple()[:6])
                info.compress_type = zipfile.ZIP_DEFLATED
                info.create_system = 3
                info.external_attr = 0o100644 << 16
                archive.writestr(info, data)
        with zipfile.ZipFile(target) as archive:
            if archive.testzip() is not None:
                raise ValueError("нарушена целостность ZIP")
        # Не перезаписываем ранее переданный архив.
        with output.open("xb") as stream:
            stream.write(target.read_bytes())
    checksum = hashlib.sha256(output.read_bytes()).hexdigest()
    checksum_file = output.with_name(output.name + ".sha256")
    with checksum_file.open("x", encoding="utf-8") as stream:
        stream.write(checksum + "  " + output.name + "\n")
    return {"archive": str(output), "bytes": output.stat().st_size,
            "files": len(members), "sha256": checksum, "checksum_file": str(checksum_file)}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out", type=Path, default=ROOT.parent / "releases" /
                        ("smeta-parser-source-" + datetime.now(timezone.utc).strftime("%Y-%m-%d") + ".zip"))
    args = parser.parse_args()
    try:
        result = build_archive(ROOT, args.out)
    except (ValueError, OSError, subprocess.CalledProcessError) as error:
        parser.exit(2, str(error) + "\n")
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
