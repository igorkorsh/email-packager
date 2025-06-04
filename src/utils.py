import base64
from pathlib import Path
from typing import List
from constants import HTML_ENTITIES


def get_files(path: Path) -> List[Path]:
    files = list(path.glob("*.html"))
    return sorted(files, key=lambda x: x.name)


def replace_html_entities(html: str) -> str:
    for unicode_char, html_entity in HTML_ENTITIES.items():
        html = html.replace(unicode_char, html_entity)
    return html


def img_to_base64(file: Path) -> str:
    image = file.read_bytes()
    encoded_image = base64.b64encode(image).decode("utf-8")
    return f"data:image/{file.suffix[1:]};base64,{encoded_image}"