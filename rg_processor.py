import os
import re
from typing import Tuple

import cv2
import numpy as np
from pdf2image import convert_from_path
import pytesseract


def _check_tessdata() -> bool:
    """Return True if por.traineddata exists in any known tessdata path."""
    lang_file = "por.traineddata"
    prefix = os.environ.get("TESSDATA_PREFIX")
    base_dir = os.path.dirname(__file__)
    possible_dirs = [
        prefix if prefix else "",
        os.path.join(base_dir, "tessdata"),
        base_dir,
        "/usr/share/tesseract-ocr/5/tessdata",
        "/usr/share/tesseract-ocr/4.00/tessdata",
        "/usr/share/tesseract-ocr/tessdata",
        "/usr/share/tesseract/tessdata",
    ]
    for d in filter(None, possible_dirs):
        if os.path.exists(os.path.join(d, lang_file)):
            os.environ["TESSDATA_PREFIX"] = d
            return True
    return False


def _ensure_tessdata():
    """Set TESSDATA_PREFIX to a directory containing por.traineddata."""
    if not _check_tessdata():
        raise RuntimeError(
            "Arquivo 'por.traineddata' não encontrado. Copie-o para a pasta "
            "'tessdata' do projeto ou defina a variável TESSDATA_PREFIX."
        )

def _load_image(file_path: str) -> "np.ndarray":
    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".pdf":
        images = convert_from_path(file_path, first_page=1, last_page=1)
        if not images:
            raise ValueError("PDF sem páginas")
        image = np.array(images[0])
        return cv2.cvtColor(image, cv2.COLOR_RGB2BGR)
    image = cv2.imread(file_path)
    if image is None:
        raise ValueError("Não foi possível abrir a imagem")
    return image


def _preprocess_image(image: "np.ndarray") -> "np.ndarray":
    if image.shape[0] > image.shape[1]:
        image = cv2.rotate(image, cv2.ROTATE_90_CLOCKWISE)
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    gray = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]
    return gray


def extract_text(file_path: str) -> str:
    _ensure_tessdata()
    image = _load_image(file_path)
    processed = _preprocess_image(image)
    try:
        return pytesseract.image_to_string(processed, lang="por")
    except pytesseract.TesseractNotFoundError:
        raise RuntimeError(
            "Tesseract OCR não encontrado. Verifique a instalação e se a "
            "variável TESSDATA_PREFIX aponta para o diretório 'tessdata'."
        )
    except pytesseract.TesseractError as e:
        msg = str(e)
        if "failed loading language" in msg.lower() and not _check_tessdata():
            raise RuntimeError(
                "O idioma português não está configurado no Tesseract. "
                "Instale o arquivo 'por.traineddata' e defina TESSDATA_PREFIX "
                "para o diretório tessdata."
            )
        raise RuntimeError(f"Erro no Tesseract: {e}")


def parse_rg_text(text: str) -> Tuple[str, str, str]:
    nome = ""
    cpf = ""
    rg = ""
    nome_match = re.search(r"nome[:\s-]*([A-ZÀ-Ú\s]+)", text, re.IGNORECASE)
    if nome_match:
        nome = nome_match.group(1).strip()
    cpf_match = re.search(r"(\d{3}\.?\d{3}\.?\d{3}-?\d{2})", text)
    if cpf_match:
        cpf = cpf_match.group(1)
    rg_match = re.search(r"(\d{2}\.?\d{3}\.?\d{3}-?\d)", text)
    if rg_match:
        rg = rg_match.group(1)
    return nome, cpf, rg


def extract_rg_data(file_path: str) -> Tuple[str, str, str]:
    text = extract_text(file_path)
    return parse_rg_text(text)