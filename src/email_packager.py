import re
import shutil
import logging
import email.message, email.generator
from bs4 import BeautifulSoup
from bs4.formatter import HTMLFormatter
from pathlib import Path
from win32com.client import Dispatch
from playwright.sync_api import sync_playwright
from typing import TypedDict, List, Optional
from utils import get_files, replace_html_entities, img_to_base64


class EmailData(TypedDict):
    path: Path
    content: BeautifulSoup
    subject: str
    images: dict[str, Path]
    is_oft: bool
    is_eml: bool
    is_pdf: bool


class ErrorHandler(logging.Handler):
    def __init__(self):
        super().__init__()
        self.errors = []
    
    def emit(self, record):
        if record.levelno == logging.ERROR:
            self.errors.append(record)


class EmailPackager:
    def __init__(self):
        self.data: List[EmailData] = []
        self.input_dir: Optional[Path] = None
        self.output_dir: Optional[Path] = None
        self.options = {
            "EMBED_IMAGES": True,
            "REMOVE_MINDBOX_VARIABLES": True
        }
        self.error_handler = ErrorHandler()
        self.logger = logging.getLogger(__name__)
        self.logger.setLevel(logging.ERROR)
        self.logger.addHandler(self.error_handler)
  
    def is_valid_html(self, file: Path) -> bool:
        try:
            content = file.read_text(encoding="utf-8")
            soup = BeautifulSoup(content, "html.parser")
            return bool(soup.find("html"))
        except FileNotFoundError:
            self.logger.error(f"[is_valid_html] File not found: {file}")
            return False
        except PermissionError:
            self.logger.error(f"[is_valid_html] Permission error: {file}")
            return False
        except UnicodeDecodeError:
            self.logger.error(f"[is_valid_html] Unicode decode error: {file}")
            return False
        except Exception as e:
            self.logger.error(f"[is_valid_html] Unknown error: {file}: {e}")
            return False

    @staticmethod
    def get_subject(soup: BeautifulSoup) -> str:
        title = soup.find("title")
        if title is None:
            return ""
        return title.text.strip()
    
    def get_images(self, soup: BeautifulSoup) -> dict[str, Path]:
        images = {}
        for img in soup.find_all("img"):
            src = img.get("src")
            if not src:
                continue
            
            if src.startswith(("http://", "https://")):
                continue
            elif src.startswith("data:"):
                continue
            elif src.startswith("/"):
                image_path = Path(self.input_dir / src.lstrip("/"))
            else:
                image_path = Path(self.input_dir / src)

            if image_path.exists():
                images[src] = image_path
        return images

    def get_data(self):
        files = get_files(self.input_dir)
        for file in files:
            if self.is_valid_html(file):
                content = file.read_text(encoding="utf-8")
                soup = BeautifulSoup(content, "html.parser")
                self.data.append({
                    "path": file,
                    "content": soup,
                    "subject": self.get_subject(soup),
                    "images": self.get_images(soup),
                    "is_oft": True,
                    "is_eml": False,
                    "is_pdf": False
                })

    @staticmethod
    def write_html(soup: BeautifulSoup, path: Path) -> None:
        try:
            content = BeautifulSoup(str(soup), "html.parser")
        
            for link in content.find_all("link", rel="preload"):
                link.decompose()
                
            for img in content.find_all("img"):
                img["src"] = re.sub(r"/static", "images", img["src"])

            formatter = HTMLFormatter(entity_substitution=replace_html_entities, indent=2)
            html = content.prettify(formatter=formatter)
            path.write_text(html, encoding="utf-8")
        except Exception as e:
            self.logger.error(f"[write_html] Unknown error: {path}: {e}")
    
    @staticmethod
    def embed_images(data: EmailData) -> BeautifulSoup:
        soup = data["content"]
        for img in soup.find_all("img"):
            src = img.get("src")
            if src in data["images"]:
                img["src"] = img_to_base64(data["images"][src])
        return soup

    @staticmethod
    def remove_mindbox(soup: BeautifulSoup) -> None:
        for element in soup.find_all(attrs={"data-mindbox": True}):
            element.decompose()

    def save_to_oft(self, data: EmailData, output_path: Path) -> None:
        if self.options["EMBED_IMAGES"]:
            self.embed_images(data)
        if self.options["REMOVE_MINDBOX_VARIABLES"]:
            self.remove_mindbox(data["content"])

        try:
            file_path = output_path / f"{data['path'].stem}.oft"
            outlook = Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)
            mail.Subject = data["subject"]
            mail.HTMLBody = str(data["content"])
            mail.SaveAs(str(file_path), 2)
            mail.Close(0)
            outlook.Quit()
        except Exception as e:
            self.logger.error(f"[save_to_oft] Unknown error: {path}: {e}")
        
    def save_to_eml(self, data: EmailData, output_path: Path) -> None:
        if self.options["EMBED_IMAGES"]:
            self.embed_images(data)
        if self.options["REMOVE_MINDBOX_VARIABLES"]:
            self.remove_mindbox(data["content"])

        try:
            file_path = output_path / f"{data['path'].stem}.emltpl"
            mail = email.message.EmailMessage()
            mail.set_content(str(data["content"]), subtype="html")
            mail.add_header("Subject", data["subject"])
            with open(file_path, "w", encoding="utf-8") as f:
                generator = email.generator.Generator(f)
                generator.flatten(mail)
        except Exception as e:
            self.logger.error(f"[save_to_eml] Unknown error: {path}: {e}")
    
    def save_to_pdf(self, data: EmailData, output_path: Path) -> None:
        file = (output_path / data["path"].name).as_uri()

        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            page = browser.new_page()
            page.goto(file)
            height = page.evaluate("document.querySelector('center').scrollHeight")
            try:
                page.pdf(
                    path=output_path / f"{data['path'].stem}.pdf",
                    print_background=True,
                    width="600",
                    height=str(height)
                )
            except Exception as e:
                self.logger.error(f"[save_to_pdf] Unknown error: {file}: {e}")
            finally:
                browser.close()

    def process(self, data: EmailData):        
        email_dir = self.output_dir / data["path"].stem
        email_dir.mkdir(exist_ok=True)
        
        file_path = email_dir / data["path"].name
        self.write_html(data["content"], file_path)
        
        if data["images"]:
            image_dir = email_dir / "images"
            image_dir.mkdir(exist_ok=True)
            for src, image_path in data["images"].items():
                try:
                    shutil.copy2(image_path, image_dir)
                except Exception as e:
                    self.logger.error(f"[shutil.copy2] Unknown error: {image_path}: {e}")
        
        if data["is_oft"]:
            self.save_to_oft(data, email_dir)
        if data["is_eml"]:
            self.save_to_eml(data, email_dir)
        if data["is_pdf"]:
            self.save_to_pdf(data, email_dir)