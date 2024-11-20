import re
import shutil
import logging
from pathlib import Path
from bs4 import BeautifulSoup
from urllib.parse import urlparse


class EmailPackager:
    def __init__(self):
        self.output = "out"

    def get_files(self, source_path: str) -> list[Path]:
        files = list(Path(source_path).glob("*.html"))
        valid_files = []

        for file in files:
            if self.parse_html(file) is not None:
                valid_files.append(file)

        return sorted(valid_files, key=lambda x: x.name)

    def create_directories(self, file_path: Path, is_image_dir: bool) -> dict:
        source_dir = file_path.parent
        output_dir = source_dir / Path(self.output)
        output_dir.mkdir(exist_ok=True)

        file_dir = output_dir / file_path.stem
        file_dir.mkdir(exist_ok=True)

        directories = {
            "file": file_dir,
            "source": source_dir
        }

        if is_image_dir:
            images_dir = file_dir / "images"
            images_dir.mkdir(exist_ok=True)
            directories["images"] = images_dir

        return directories

    @staticmethod
    def parse_html(file_path: Path) -> BeautifulSoup | None:
        try:
            with open(file_path, "r", encoding="utf-8") as file:
                content = BeautifulSoup(file, "html.parser")
            return content
        except Exception as err:
            logging.error(f"Ошибка при чтении файла {file}: {err}")
            return None

    @staticmethod
    def get_images(content: BeautifulSoup, base_path: Path) -> list[Path]:
        return [
            (base_path / Path(src)).resolve()
            for tag in content.find_all("img")
            if (src := tag.get("src")) and not urlparse(src).scheme in ("http", "https")
        ]

    @staticmethod
    def copy_files(files: list[Path], output_dir: Path):
        for file in files:
            if file.exists() and file.is_file():
                try:
                    shutil.copy2(file, output_dir)
                except Exception as err:
                    logging.error(f"Ошибка при копировании файла {file.name}: {err}")

    @staticmethod
    def update_file(file_path: Path, output_dir: Path):
        try:
            with open(file_path, "r", encoding="utf-8") as file:
                content = file.read()

            pattern = r'((src|href)=["\'])static/'

            if re.search(pattern, content):
                updated_content = re.sub(pattern, r"\1images/", content)

                with open(output_dir / file_path.name, "w", encoding="utf-8") as file:
                    file.write(updated_content)
            else:
                shutil.copy2(file_path, output_dir / file_path.name)
        except Exception as err:
            logging.error(f"Ошибка при записи файла {file}: {err}")

    def package(self, file_path: Path):
        html = self.parse_html(file_path)

        if html is not None:
            images = self.get_images(html, file_path.parent)
            directories = self.create_directories(file_path, is_image_dir=bool(len(images)))

            if images:
                self.copy_files(images, directories["images"])

            self.update_file(file_path, directories["file"])