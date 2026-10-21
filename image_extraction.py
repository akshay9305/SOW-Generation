#image_extraction.py

import base64
import json
import os
import re
from pathlib import Path
from typing import Dict, List, Tuple

import fitz
from openai import AzureOpenAI

from key import *

aoai_client = AzureOpenAI(
    api_key=AZURE_OPENAI_API_KEY,
    api_version=AZURE_OPENAI_API_VERSION,
    azure_endpoint=AZURE_OPENAI_ENDPOINT,
)

IMAGE_STORE_PATH = "uploads/extracted_images/"
LOG_FILE_PATH = "Output/image_audit_log.txt"


def log_audit(message: str):
    """
    Append a message to the centralized audit log file and print to terminal.
    """
    os.makedirs(os.path.dirname(LOG_FILE_PATH), exist_ok=True)
    with open(LOG_FILE_PATH, "a", encoding="utf-8") as f:
        f.write(message + "\n")
    print(message)


def extract_images_from_pdf(
    pdf_path: str, output_images_folder: str = None
) -> List[Dict]:
    """Extract images from input PDF pages and generate Smart IDs."""
    if output_images_folder is None:
        output_images_folder = IMAGE_STORE_PATH

    pages_with_images = []
    os.makedirs(output_images_folder, exist_ok=True)

    # Sanitize filename for safe paths
    doc_name = Path(pdf_path).stem
    doc_name = re.sub(r"[^a-zA-Z0-9_-]", "_", doc_name)

    try:
        pdf_document = fitz.open(pdf_path)

        for page_num in range(pdf_document.page_count):
            page = pdf_document[page_num]
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
            image_bytes = pix.tobytes("png")

            # === SMART ID GENERATION ===
            smart_id = f"{doc_name}_Page{page_num + 1}"
            image_filename = f"{smart_id}.png"
            image_path = os.path.join(output_images_folder, image_filename)

            pix.save(image_path)

            image_base64 = base64.b64encode(image_bytes).decode("utf-8")

            pages_with_images.append(
                {
                    "page_number": page_num + 1,
                    "doc_name": doc_name,
                    "chunk_id": smart_id,
                    "image_path": image_path,
                    "image_base64": image_base64,
                    "image_bytes": image_bytes,
                }
            )

        pdf_document.close()
        print(f"Extracted {len(pages_with_images)} images to {output_images_folder}")

    except Exception as e:
        print(f"Error extracting images: {str(e)}")
        raise

    return pages_with_images


def generate_image_classification(image_base64: str) -> Tuple[bool, str]:
    """
    STRICT UI CLASSIFIER.
    Returns: (is_ui_screen: bool, description: str)
    """

    vision_prompt = """You are a Strict UI Classifier & Technical Analyst.

    TASK 1: CLASSIFY
    Determine if the image is a "User Interface Screen" (High-Fidelity Mockup, Dashboard, App Screen, Wireframe).
    - ACCEPT: High-quality UI mockups, dashboards, mobile app screens, wireframes.
    - REJECT: Stock photos, generic business images, process diagrams, flowcharts, slideshow titles, pure text screenshots.

    TASK 2: DESCRIBE
    If ACCEPTED, write a brief 1-line technical caption (e.g., 'Dashboard showing revenue metrics').
    If REJECTED, state the type (e.g., 'Process Diagram').

    OUTPUT FORMAT (JSON ONLY):
    {
        "is_ui": boolean,
        "type": "UI" or "Diagram" or "Stock" or "Text",
        "reason": "Why it was accepted or rejected",
        "caption": "Brief description"
    }
    """

    try:
        response = aoai_client.chat.completions.create(
            model=AZURE_OPENAI_VISION_DEPLOYMENT,
            messages=[
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": vision_prompt},
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:image/png;base64,{image_base64}"
                            },
                        },
                    ],
                }
            ],
            max_completion_tokens=500,
        )
        content = response.choices[0].message.content.strip()
        content = re.sub(
            r"^```json\s*|\s*```$", "", content, flags=re.MULTILINE
        ).strip()
        data = json.loads(content)

        is_ui = data.get("is_ui", False)
        description = data.get("caption", "No description")
        reason = data.get("reason", "")

        # Log the decision for debugging
        print(f"   [Vision] UI: {is_ui} | Type: {data.get('type')} | Reason: {reason}")

        return is_ui, description

    except Exception as e:
        print(f"Error classifying image: {e}")
        return False, f"Error: {e}"


def process_pdf_to_text(pdf_path: str, append_mode: bool = False):
    """
    1. Extract images.
    2. Classify: If NON-UI, Delete it.
    3. NO TEXT FILE GENERATED. We just keep the valid files in the folder.
    """
    print(f"Processing PDF: {pdf_path}")

    # Init Log
    if not append_mode and os.path.exists(LOG_FILE_PATH):
        os.remove(LOG_FILE_PATH)
    log_audit(f"=== PROCESSING PDF: {os.path.basename(pdf_path)} ===")

    images = extract_images_from_pdf(pdf_path)

    if not images:
        log_audit("No images extracted.")
        return

    approved_count = 0

    for i, img in enumerate(images):
        print(f"[{i+1}/{len(images)}] Classifying {img['chunk_id']}...")

        is_ui, _ = generate_image_classification(img["image_base64"])

        if is_ui:
            # KEEP IT
            log_audit(f"   [KEPT] {img['chunk_id']} (UI)")
            approved_count += 1
        else:
            # DELETE IT
            if os.path.exists(img["image_path"]):
                os.remove(img["image_path"])
            log_audit(f"   [DELETED] {img['chunk_id']} (Not UI)")

    log_audit(
        f"=== Done. Kept {approved_count} UI images. Deleted {len(images) - approved_count} others. ==="
    )


# PPT EXTRACTION LOGIC
def extract_images_from_ppt(
    ppt_path: str, output_images_folder: str = None
) -> List[Dict]:
    """Extract full slides as images from PPT using PowerPoint COM automation."""
    if output_images_folder is None:
        output_images_folder = IMAGE_STORE_PATH

    pages_with_images = []
    os.makedirs(output_images_folder, exist_ok=True)

    # Sanitize filename for safe paths
    doc_name = Path(ppt_path).stem
    doc_name = re.sub(r"[^a-zA-Z0-9_-]", "_", doc_name)

    try:
        import win32com.client

        # Initialize PowerPoint Application
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")

        # Open Presentation
        abs_path = os.path.abspath(ppt_path)
        presentation = powerpoint.Presentations.Open(abs_path, WithWindow=False)

        for i, slide in enumerate(presentation.Slides):
            page_num = i + 1
            smart_id = f"{doc_name}_Slide{page_num}"
            image_filename = f"{smart_id}.png"
            image_path = os.path.join(output_images_folder, image_filename)


            max_retries = 3
            for attempt in range(max_retries):
                try:
                    slide.Export(os.path.abspath(image_path), "PNG")
                    break  # Success, exit retry loop
                except Exception as e:
                    if attempt < max_retries - 1:
                        import time
                        print(f"   [Retry {attempt + 1}] Slide {page_num} export failed, retrying in 2s...")
                        time.sleep(2)
                    else:
                        print(f"   [Error] Failed to export Slide {page_num} after {max_retries} attempts: {e}")
                        continue  # Skip this slide

            if not os.path.exists(image_path):
                continue  # Skip if the slide wasn't saved

            with open(image_path, "rb") as f:
                image_bytes = f.read()
            image_base64 = base64.b64encode(image_bytes).decode("utf-8")

            pages_with_images.append(
                {
                    "page_number": page_num,
                    "doc_name": doc_name,
                    "chunk_id": smart_id,
                    "image_path": image_path,
                    "image_base64": image_base64,
                    "image_bytes": image_bytes,
                }
            )

        presentation.Close()
        print(f"Extracted {len(pages_with_images)} slides to {output_images_folder}")

    except ImportError:
        print(
            "Error: 'pywin32' library is missing. Please install it using 'pip install pywin32'."
        )
        raise
    except Exception as e:
        print(f"Error extracting PPT slides: {str(e)}")
        raise

    return pages_with_images


def process_ppt_to_text(ppt_path: str, append_mode: bool = False):
    """
    1. Extract slides as images.
    2. Classify: If NON-UI, Delete it.
    3. NO TEXT FILE GENERATED. We just keep the valid files in the folder.
    """
    print(f"Processing PPT: {ppt_path}")

    # Init Log
    if not append_mode and os.path.exists(LOG_FILE_PATH):
        os.remove(LOG_FILE_PATH)
    log_audit(f"=== PROCESSING PPT: {os.path.basename(ppt_path)} ===")

    images = extract_images_from_ppt(ppt_path)

    if not images:
        log_audit("No slides extracted.")
        return

    approved_count = 0

    for i, img in enumerate(images):
        print(f"[{i+1}/{len(images)}] Classifying {img['chunk_id']}...")

        is_ui, _ = generate_image_classification(img["image_base64"])

        if is_ui:
            # KEEP IT
            log_audit(f"   [KEPT] {img['chunk_id']} (UI)")
            approved_count += 1
        else:
            # DELETE IT
            if os.path.exists(img["image_path"]):
                os.remove(img["image_path"])
            log_audit(f"   [DELETED] {img['chunk_id']} (Not UI)")

    log_audit(
        f"=== Done. Kept {approved_count} UI slides. Deleted {len(images) - approved_count} others. ==="
    )


# GALLERY RETRIEVAL LOGIC
def get_valid_ui_images() -> List[Dict]:
    """
    Scans the output folder for any surviving images.
    Since we deleted the non-UIs, everything here is valid.
    Returns images sorted by page/slide number.
    """
    if not os.path.exists(IMAGE_STORE_PATH):
        return []

    valid_images = []
    # Scan directory
    for f in os.listdir(IMAGE_STORE_PATH):
        if f.lower().endswith((".png", ".jpg", ".jpeg")):
            valid_images.append(
                {
                    "id": Path(f).stem,
                    "description": "UI Screenshot",
                    "path": os.path.join(IMAGE_STORE_PATH, f).replace("\\", "/"),
                }
            )

    # Sort numerically by page/slide number extracted from filename
    def extract_number(item):
        # Match patterns like Page1, Slide10, etc.
        match = re.search(r"(?:Page|Slide)(\d+)", item["id"])
        if match:
            return int(match.group(1))
        return 0  # Default if no number found

    valid_images.sort(key=extract_number)

    return valid_images


def assign_images_to_sections(
    sections: Dict[str, str], descriptions_path: str = None
) -> Dict[str, List[Dict]]:
    """
    GALLERY MODE:
    1. Find the 'User Interface' section.
    2. Dump ALL available UI images into it.
    """
    log_audit("\n=== GALLERY ASSIGNMENT MODE ===")

    ui_images = get_valid_ui_images()
    if not ui_images:
        log_audit("No UI images available to assign. (Folder empty)")
        return {}

    log_audit(f"Found {len(ui_images)} Valid UI Images in folder.")

    target_key = None
    for key in sections.keys():

        k_normalized = key.lower().replace("_", " ")
        if (
            "user interface" in k_normalized
            or "ui design" in k_normalized
            or "screens" in k_normalized
        ):
            target_key = key
            break

    final_placements = {}

    if target_key:
        final_placements[target_key] = ui_images
        log_audit(
            f"   [ASSIGNED] All {len(ui_images)} images -> Section '{target_key}'"
        )
    else:
        log_audit(
            "   [WARNING] No 'User Interface' section found in SOW. Images unassigned."
        )

    return final_placements


if __name__ == "__main__":
    pass