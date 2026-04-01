import sys
import time
import tempfile
import subprocess
from pathlib import Path
from shutil import copyfileobj, which
from urllib.parse import quote

import requests
from docx import Document


API_DELAY_SECONDS = 0.4
MAX_RETRIES = 3
BACKOFF_BASE_SECONDS = 1.0
TARGET_COLOR = (255, 0, 0)
SUPPORTED_EXTENSIONS = [".docx", ".odt", ".rtf"]


def get_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def normalize_color(color: str) -> tuple[int, int, int]:
    color = color.strip().replace("(", "").replace(")", "")
    parts = [p.strip() for p in color.split(",")]
    return tuple(map(int, parts))


def run_rgb_tuple(run):
    rgb = run.font.color.rgb
    if rgb is None:
        return None
    return tuple(rgb)


def doc_scour(filename: str, color: str) -> list[str]:
    target_color = normalize_color(color)
    doc = Document(filename)
    found_text = []

    for para in doc.paragraphs:
        runs = para.runs
        i = 0

        while i < len(runs):
            current_color = run_rgb_tuple(runs[i])

            if current_color == target_color:
                text_chunk = runs[i].text
                i += 1

                while i < len(runs) and run_rgb_tuple(runs[i]) == target_color:
                    text_chunk += runs[i].text
                    i += 1

                cleaned = text_chunk.strip()
                if cleaned:
                    found_text.append(cleaned)
            else:
                i += 1

    return found_text


def safe_filename(name: str) -> str:
    bad_chars = '<>:"/\\|?*'
    for ch in bad_chars:
        name = name.replace(ch, "_")
    return name.strip()


def make_session() -> requests.Session:
    session = requests.Session()
    session.headers.update(
        {
            "User-Agent": "scriptHelperMagic/1.0",
            "Accept": "application/json",
        }
    )
    return session


def request_with_retry(
    session: requests.Session,
    url: str,
    *,
    stream: bool = False,
    timeout: int = 20,
):
    last_error = None

    for attempt in range(MAX_RETRIES):
        try:
            response = session.get(url, stream=stream, timeout=timeout)

            if response.status_code == 404:
                response.close()
                raise FileNotFoundError(f"404 Not Found: {url}")

            if response.status_code == 429:
                retry_after = response.headers.get("Retry-After")
                sleep_for = float(retry_after) if retry_after else BACKOFF_BASE_SECONDS * (2 ** attempt)
                response.close()

                if attempt == MAX_RETRIES - 1:
                    raise RuntimeError(f"429 Too Many Requests: {url}")

                print(f"Rate limited. Waiting {sleep_for:.1f}s before retrying...")
                time.sleep(sleep_for)
                continue

            if 500 <= response.status_code < 600:
                sleep_for = BACKOFF_BASE_SECONDS * (2 ** attempt)
                response.close()

                if attempt == MAX_RETRIES - 1:
                    raise RuntimeError(f"Server error {response.status_code}: {url}")

                print(f"Server error {response.status_code}. Waiting {sleep_for:.1f}s before retrying...")
                time.sleep(sleep_for)
                continue

            response.raise_for_status()
            return response

        except FileNotFoundError:
            raise

        except requests.RequestException as e:
            last_error = e

            if attempt == MAX_RETRIES - 1:
                break

            sleep_for = BACKOFF_BASE_SECONDS * (2 ** attempt)
            print(f"Request failed ({e}). Waiting {sleep_for:.1f}s before retrying...")
            time.sleep(sleep_for)

    raise RuntimeError(f"Failed after {MAX_RETRIES} attempts: {url}") from last_error


def exact_card_lookup(session: requests.Session, card_name: str):
    encoded_name = quote(card_name.strip(), safe="")
    url = f"https://api.scryfall.com/cards/named?exact={encoded_name}"
    response = request_with_retry(session, url, timeout=15)
    data = response.json()
    response.close()
    time.sleep(API_DELAY_SECONDS)
    return data


def fuzzy_card_lookup(session: requests.Session, card_name: str):
    encoded_name = quote(card_name.strip(), safe="")
    url = f"https://api.scryfall.com/cards/named?fuzzy={encoded_name}"
    response = request_with_retry(session, url, timeout=15)
    data = response.json()
    response.close()
    time.sleep(API_DELAY_SECONDS)
    return data


def card_to_image_and_id(
    session: requests.Session,
    card_name: str,
    cache: dict[str, tuple[object, str]],
):
    normalized = card_name.strip()
    if normalized in cache:
        return cache[normalized]

    try:
        data = exact_card_lookup(session, normalized)
    except FileNotFoundError:
        data = fuzzy_card_lookup(session, normalized)

    if data.get("object") == "error":
        raise ValueError(f"Bad card name: {card_name}")

    if "card_faces" in data:
        faces = data["card_faces"]
        image_urls = []

        for face in faces:
            if "image_uris" in face and "png" in face["image_uris"]:
                image_urls.append(face["image_uris"]["png"])

        if len(image_urls) >= 2:
            result = ((image_urls[0], image_urls[1]), data["id"])
        elif len(image_urls) == 1:
            result = (image_urls[0], data["id"])
        else:
            raise ValueError(f"No downloadable face images found for: {card_name}")
    else:
        if "image_uris" not in data or "png" not in data["image_uris"]:
            raise ValueError(f"No image found for: {card_name}")
        result = (data["image_uris"]["png"], data["id"])

    cache[normalized] = result
    return result


def download_file(session: requests.Session, url: str, out_path: Path) -> None:
    response = request_with_retry(session, url, stream=True, timeout=30)
    response.raw.decode_content = True

    with open(out_path, "wb") as f:
        copyfileobj(response.raw, f)

    response.close()
    time.sleep(API_DELAY_SECONDS)


def unique_preserve_order(items: list[str]) -> list[str]:
    seen = set()
    result = []

    for item in items:
        cleaned = item.strip()
        if cleaned and cleaned not in seen:
            seen.add(cleaned)
            result.append(cleaned)

    return result


def find_input_file(script_dir: Path, scriptname: str) -> Path | None:
    for ext in SUPPORTED_EXTENSIONS:
        candidate = script_dir / f"{scriptname}{ext}"
        if candidate.exists():
            return candidate
    return None


def find_soffice() -> str | None:
    cmd = which("soffice")
    if cmd:
        return cmd

    common_windows_paths = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]

    for path in common_windows_paths:
        if Path(path).exists():
            return path

    return None


def convert_to_docx(input_path: Path) -> Path:
    suffix = input_path.suffix.lower()

    if suffix == ".docx":
        return input_path

    if suffix not in {".odt", ".rtf"}:
        raise ValueError(f"Unsupported input format: {suffix}")

    soffice = find_soffice()
    if not soffice:
        raise RuntimeError(
            "This file is not .docx, and LibreOffice was not found for conversion.\n"
            "Install LibreOffice or save the file as .docx first."
        )

    temp_dir = Path(tempfile.mkdtemp(prefix="scriptHelperMagic_"))
    cmd = [
        soffice,
        "--headless",
        "--convert-to",
        "docx",
        "--outdir",
        str(temp_dir),
        str(input_path),
    ]

    result = subprocess.run(cmd, capture_output=True, text=True)

    if result.returncode != 0:
        raise RuntimeError(
            f"LibreOffice conversion failed.\nSTDOUT:\n{result.stdout}\nSTDERR:\n{result.stderr}"
        )

    converted_path = temp_dir / f"{input_path.stem}.docx"
    if not converted_path.exists():
        raise RuntimeError("Conversion appeared to succeed, but the .docx output was not found.")

    return converted_path


def main():
    base_dir = get_base_dir()
    script_dir = base_dir / "script"

    scriptname = input("Script name: ").strip()
    input_path = find_input_file(script_dir, scriptname)
    output_dir = base_dir / "img" / scriptname

    if input_path is None:
        print(f"Could not find input file for: {scriptname}")
        print(f"Checked in: {script_dir}")
        print(f"Supported extensions: {', '.join(SUPPORTED_EXTENSIONS)}")
        input("Press Enter to close...")
        return

    output_dir.mkdir(parents=True, exist_ok=True)

    try:
        docx_path = convert_to_docx(input_path)
        imglist = doc_scour(str(docx_path), str(TARGET_COLOR))
    except Exception as e:
        print(f"Failed to read document: {e}")
        print(f"Checked path: {input_path}")
        print("Make sure the file is a real document and, for .odt/.rtf, that LibreOffice is installed.")
        input("Press Enter to close...")
        return

    imglist = unique_preserve_order(imglist)

    if not imglist:
        print("No matching red text found.")
        input("Press Enter to close...")
        return

    session = make_session()
    api_cache: dict[str, tuple[object, str]] = {}
    seen_ids = set()

    for card_name in imglist:
        safe_name = safe_filename(card_name)

        try:
            card_data, card_id = card_to_image_and_id(session, card_name, api_cache)

            if card_id in seen_ids:
                print(f"Skipping duplicate card ID: {card_name}")
                continue

            if isinstance(card_data, tuple):
                front_url, back_url = card_data

                front_path = output_dir / f"{safe_name} Front.png"
                back_path = output_dir / f"{safe_name} Back.png"

                if not front_path.exists():
                    print(f"Downloading double-faced card front: {card_name}...")
                    download_file(session, front_url, front_path)

                if not back_path.exists():
                    print(f"Downloading double-faced card back: {card_name}...")
                    download_file(session, back_url, back_path)
            else:
                out_path = output_dir / f"{safe_name}.png"

                if out_path.exists():
                    print(f"Already exists, skipping file download: {card_name}")
                else:
                    print(f"Downloading {card_name}...")
                    download_file(session, card_data, out_path)

            seen_ids.add(card_id)

        except Exception as e:
            print(f"Skipped '{card_name}': {e}")

    session.close()
    print("All done! Feel free to close this window.")
    input("")


if __name__ == "__main__":
    main()