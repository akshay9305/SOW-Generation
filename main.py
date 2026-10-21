# main.py

import argparse
import email
import json
import os
import re
import uuid
from email import policy
from pathlib import Path
from typing import List

import fitz
from azure.core.credentials import AzureKeyCredential
from azure.search.documents import SearchClient
from azure.search.documents.models import VectorizedQuery
from docx import Document
from docx.oxml import OxmlElement
from docx.shared import Inches
from docx.text.paragraph import Paragraph
from markdown_it import MarkdownIt
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

# Import image extraction logic
from image_extraction import (
    IMAGE_STORE_PATH,
    assign_images_to_sections,
    process_pdf_to_text,
    process_ppt_to_text,
)

# Import keys and prompt agents from your local modules
from key import *
from prompt import *

# Initialize Azure Clients
aoai_client = AzureOpenAI(
    api_key=AZURE_OPENAI_API_KEY,
    api_version=AZURE_OPENAI_API_VERSION,
    azure_endpoint=AZURE_OPENAI_ENDPOINT,
)

search_client = SearchClient(
    endpoint=AZURE_SEARCH_ENDPOINT,
    index_name=AZURE_SEARCH_INDEX,
    credential=AzureKeyCredential(AZURE_SEARCH_KEY),
)

# ==============================================================================
# DOCX WRITER CLASS
# ==============================================================================


class DocxWriter:
    """Helper class to write content section-wise into a Word document."""

    DEFAULT_SECTION_PLACEHOLDER = "[PLACEHOLDER]"

    @staticmethod
    def _resolve(path: Path) -> str:
        return str(path)

    @staticmethod
    def _normalize(text: str) -> str:
        """Normalize text for robust heading comparison."""
        return text.lower().strip().replace("_", " ")

    @staticmethod
    def _is_heading(para):
        """Check if paragraph is a heading."""
        return para.style.name.startswith("Heading")

    @staticmethod
    def _delete_paragraph(para):
        """Delete a paragraph from the document."""
        p = para._element
        p.getparent().remove(p)
        p._p = p._element = None

    @staticmethod
    def _set_table_borders(table):
        """
        Add borders to all cells in a table programmatically.
        This ensures tables always have visible grid lines, regardless of template styles.
        """
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn

        tbl = table._tbl
        tblPr = tbl.tblPr if tbl.tblPr is not None else OxmlElement("w:tblPr")
        if tbl.tblPr is None:
            tbl.insert(0, tblPr)

        # Create table borders element
        tblBorders = OxmlElement("w:tblBorders")
        for border_name in ("top", "left", "bottom", "right", "insideH", "insideV"):
            border = OxmlElement(f"w:{border_name}")
            border.set(qn("w:val"), "single")
            border.set(qn("w:sz"), "4")  # border width
            border.set(qn("w:space"), "0")
            border.set(qn("w:color"), "000000")  # black
            tblBorders.append(border)
        tblPr.append(tblBorders)

    @staticmethod
    def _parse_markdown_with_parser(md_text: str) -> list:
        """
        Parse Markdown using markdown-it-py and return structured content.
        Returns a list of dicts: {'type': 'paragraph'|'table', 'content': ...}
        """
        md = MarkdownIt("commonmark", {"typographer": True}).enable("table")
        tokens = md.parse(md_text)

        result = []
        i = 0
        while i < len(tokens):
            token = tokens[i]

            # Handle Tables
            if token.type == "table_open":
                table_data = []
                current_row = []
                i += 1
                while i < len(tokens) and tokens[i].type != "table_close":
                    t = tokens[i]
                    # Treat thead/tbody as transparent wrappers - just skip them
                    if t.type in (
                        "thead_open",
                        "thead_close",
                        "tbody_open",
                        "tbody_close",
                    ):
                        i += 1
                        continue
                    if t.type == "tr_open":
                        current_row = []
                    elif t.type == "tr_close":
                        table_data.append(current_row)
                    elif t.type in ("th_open", "td_open"):
                        # Get inline content
                        i += 1
                        if i < len(tokens) and tokens[i].type == "inline":
                            current_row.append(tokens[i].content)
                        else:
                            current_row.append("")
                    i += 1
                if table_data:
                    max_cols = max((len(r) for r in table_data), default=0)
                    if max_cols:
                        for row in table_data:
                            if len(row) < max_cols:
                                row.extend([""] * (max_cols - len(row)))
                    result.append({"type": "table", "content": table_data})

            # Handle Paragraphs
            elif token.type == "paragraph_open":
                i += 1
                if i < len(tokens) and tokens[i].type == "inline":
                    para_text = tokens[i].content
                    # Clean markdown bold
                    para_text = re.sub(r"\*\*(.*?)\*\*", r"\1", para_text)
                    result.append({"type": "paragraph", "content": para_text})

            # Handle Bullet Lists
            elif token.type == "bullet_list_open":
                list_items = []
                i += 1
                while i < len(tokens) and tokens[i].type != "bullet_list_close":
                    if tokens[i].type == "list_item_open":
                        i += 1
                        # Find paragraph inside list item
                        while i < len(tokens) and tokens[i].type != "list_item_close":
                            if tokens[i].type == "inline":
                                list_items.append(tokens[i].content)
                            i += 1
                    i += 1
                if list_items:
                    result.append({"type": "list", "content": list_items})

            # Handle Code Fences (Recursion)
            elif token.type == "fence":
                # Recursively parse content inside fence
                inner_content = DocxWriter._parse_markdown_with_parser(token.content)
                result.extend(inner_content)

            i += 1

        return result

    @staticmethod
    def append_under_heading_in_newdoc(
        new_doc_path: Path, heading_text: str, text: str, images: List[dict] = None
    ) -> None:
        """
        Appends content (Text, Tables, Images) under a specific heading in a DOCX.
        Uses markdown-it-py for robust Markdown parsing.
        """
        new_doc_path_str = DocxWriter._resolve(new_doc_path)
        doc = Document(new_doc_path_str)
        target_norm = DocxWriter._normalize(heading_text)

        # 1. Find or create the heading
        heading_para = None
        for p in doc.paragraphs:
            if (
                DocxWriter._is_heading(p)
                and DocxWriter._normalize(p.text) == target_norm
            ):
                heading_para = p
                break

        if not heading_para:
            heading_para = doc.add_heading(heading_text, level=1)

        # 2. Remove placeholder under this heading
        insert_point = heading_para._p
        current = heading_para._p.getnext()
        while current is not None:
            if current.tag.endswith("p"):
                para = Paragraph(current, heading_para._parent)
                if para.text.strip() == DocxWriter.DEFAULT_SECTION_PLACEHOLDER:
                    DocxWriter._delete_paragraph(para)
                    break
            current = current.getnext()

        def _create_paragraph_element() -> Paragraph:
            """Create a new paragraph element immediately after the insert point."""
            nonlocal insert_point
            new_p = OxmlElement("w:p")
            insert_point.addnext(new_p)
            para = Paragraph(new_p, heading_para._parent)
            insert_point = new_p
            return para

        # 3. Parse Markdown content using robust parser
        parsed_content = DocxWriter._parse_markdown_with_parser(text)

        # 4. Write parsed content to DOCX
        for item in parsed_content:
            if item["type"] == "paragraph":
                p = _create_paragraph_element()
                p.add_run(item["content"])

            elif item["type"] == "table":
                table_data = item["content"]
                if table_data and len(table_data) > 0:
                    num_cols = len(table_data[0])
                    table = doc.add_table(rows=len(table_data), cols=num_cols)
                    # Add borders programmatically (always works, no style dependency)
                    DocxWriter._set_table_borders(table)
                    for i, row_data in enumerate(table_data):
                        for j, cell_text in enumerate(row_data):
                            cell = table.rows[i].cells[j]

                            # Clear default paragraph
                            cell.text = ""

                            # Split numbered items: 1. 2. 3. → separate lines
                            items = re.split(r"\s*(?=\d+\.\s)", cell_text.strip())

                            for idx, item in enumerate(items):
                                if not item.strip():
                                    continue

                                if idx == 0:
                                    p = cell.paragraphs[0]
                                else:
                                    p = cell.add_paragraph()

                                p.add_run(item.strip())

                    # Move table to the insertion point under the heading
                    insert_point.addnext(table._tbl)
                    insert_point = table._tbl

            elif item["type"] == "list":
                for list_item in item["content"]:
                    p = _create_paragraph_element()
                    clean_text = re.sub(r"\*\*(.*?)\*\*", r"\1", list_item)
                    p.add_run(f"• {clean_text}")

        # 5. Insert Images if any
        if images:
            for img in images:
                try:
                    abs_img_path = os.path.abspath(img["path"])
                    if os.path.exists(abs_img_path):
                        p = _create_paragraph_element()
                        run = p.add_run()
                        run.add_picture(abs_img_path, width=Inches(6.0))
                except Exception as e:
                    print(f"Error inserting image {img.get('path')}: {e}")

        doc.save(new_doc_path_str)


# ==============================================================================
# EMBEDDINGS & ANONYMIZATION
# ==============================================================================


def get_embedding(text: str) -> list:
    max_chars = 30000
    if len(text) > max_chars:
        text = text[:max_chars]
        print(
            f"⚠️ Text truncated to {max_chars} characters for embedding to avoid token limit (max ~8192 tokens)"
        )

    response = aoai_client.embeddings.create(
        model=AZURE_OPENAI_EMBEDDING_DEPLOYMENT, input=text
    )
    return response.data[0].embedding


def _call_gpt_for_anonymize(system_prompt: str, user_text: str) -> str:
    """Helper to call Azure OpenAI for anonymization."""
    try:
        response = aoai_client.chat.completions.create(
            model=AZURE_OPENAI_DEPLOYMENT_CHAT,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_text},
            ],
            temperature=0.1,
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        print(f"  [Error] GPT Call Failed: {e}")
        return ""


def anon_writer_agent(text: str, feedback: str = None) -> str:
    """
    Dynamic Writer Agent: Handles both drafting AND fixing.
    """
    if len(text) < 50:
        return ""

    if feedback is None:
        system = (
            "You are a professional legal technical writer. Your goal is to rewrite the following SOW page content "
            "to be COMPLETELY GENERIC and UNIVERSAL. \n"
            "RULES:\n"
            "1. REMOVE all specific Client Names, Vendor/Company Names, Individual Names, Locations, and Addresses.\n"
            "2. REMOVE all specific Dates, Durations. Replace with 'the Effective Date' or 'a defined period'.\n"
            "3. REMOVE all specific Monetary Amounts. Replace with 'the Fees' or 'agreed compensation'.\n"
            "4. REMOVE specific Project Names, Internal Codes, or Proprietary System Names.\n"
            "   - Use generic terms like 'the Project', 'the Legacy System'.\n"
            "   - KEEP standard industry terms (e.g., 'AWS', 'Python', 'React', 'SAP').\n"
            "5. DO NOT use placeholders like '[Client Name]'. Use generic nouns like 'the Client', 'the Service Provider'.\n"
            "6. KEEP the technical intent, scope description, deliverables list, and legal structure intact.\n"
            "7. Do not include introductory phrases. Just output the text.\n"
            "8. If the text is just a footer/header or noise, output 'N/A'."
        )
        return _call_gpt_for_anonymize(system, text)
    else:
        system = (
            "You are a SURGICAL TEXT FIXER. You have been given a draft and a specific list of issues from an Auditor.\n"
            "YOUR TASK:\n"
            "1. Read the list of issues carefully.\n"
            "2. Find ONLY the specific text quoted in the issues.\n"
            "3. Replace ONLY those specific items with generic terms.\n"
            "4. DO NOT rewrite, rephrase, or change any other part of the text.\n"
            "5. Output the FULL corrected text.\n"
            "\nIMPORTANT: Keep everything else EXACTLY as it was. Only fix what is listed."
        )
        user_input = f"CURRENT DRAFT:\n{text}\n\nAUDITOR ISSUES TO FIX:\n{feedback}"
        return _call_gpt_for_anonymize(system, user_input)


def anon_auditor_agent(text: str) -> tuple:
    """
    Enhanced Auditor Agent: Returns (True, 'SAFE') or (False, 'Detailed Issue List').
    """
    if text == "N/A" or not text:
        return True, "SAFE"

    system = (
        "You are a strict PII Auditor. Check the following text for PII or Specific Entity Data.\n"
        "LOOK FOR:\n"
        "- Specific Company Names (e.g., 'Acme Corp', 'Microsoft')\n"
        "- Specific Dates (e.g., 'July 4, 1776', 'Q1 2024')\n"
        "- Specific Monetary Amounts (e.g., '$150,000')\n"
        "- Specific Names of People\n"
        "- Specific Proprietary Project Names or Internal Codes\n\n"
        "IGNORE generic terms like 'the Client', 'the Provider', 'TBD', 'As agreed', 'the Fees'.\n\n"
        "RESPONSE FORMAT:\n"
        "- If the text is fully generic: Reply strictly 'SAFE'.\n"
        "- If specific info is found: Reply with a numbered list in this exact format:\n"
        "  FLAGGED:\n"
        '  1. [Type]: "[Exact Quote]" - found in the sentence starting with "[Context]"\n'
        '  2. [Type]: "[Exact Quote]" - found in the sentence starting with "[Context]"\n'
        "  ... (list ALL issues found)"
    )
    result = _call_gpt_for_anonymize(system, text)

    if "SAFE" in result.upper() and "FLAGGED" not in result.upper():
        return True, "SAFE"
    else:
        return False, result


def anonymize_text(raw_text: str) -> str:
    """Run the 2-Agent Write <-> Audit loop on a single page."""
    current_draft = anon_writer_agent(raw_text, feedback=None)
    if not current_draft or current_draft == "N/A":
        return ""

    for attempt in range(3):
        is_safe, feedback = anon_auditor_agent(current_draft)
        if is_safe:
            return current_draft
        else:
            current_draft = anon_writer_agent(current_draft, feedback=feedback)

    return current_draft


# ==============================================================================
# CHUNKING & EXTRACTION
# ==============================================================================


def extract_pages_as_chunks(pdf_path, doc_id, doc_role):
    doc = fitz.open(pdf_path)
    chunks = []

    for pno in range(doc.page_count):
        page = doc.load_page(pno)
        text = page.get_text("text").strip()
        if not text:
            continue

        chunks.append(
            {
                "id": str(uuid.uuid4()),
                "doc_id": doc_id,
                "doc_role": doc_role,
                "chunk_id": f"{doc_id}_page_{pno+1}",
                "page": pno + 1,
                "text": text,
            }
        )

    doc.close()
    return chunks


def clean_text(t: str) -> str:
    t = re.sub(r"[\x00-\x1F\x7F]", " ", t)
    t = re.sub(r"\s+", " ", t)
    return t.strip()


def extract_all_text_from_shape(shape) -> str:
    texts = []
    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        for child_shape in shape.shapes:
            texts.append(extract_all_text_from_shape(child_shape))
    elif shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                if cell.text_frame:
                    texts.append(cell.text_frame.text)
    elif shape.has_text_frame:
        texts.append(shape.text_frame.text)
    elif hasattr(shape, "text"):
        texts.append(shape.text)
    return "\n".join(t for t in texts if t.strip())


def extract_ppt_slides_as_chunks(ppt_path: str, doc_id: str, doc_role: str) -> list:
    prs = Presentation(ppt_path)
    chunks = []
    for slide_no, slide in enumerate(prs.slides, start=1):
        text_blocks = []
        for shape in slide.shapes:
            extracted = extract_all_text_from_shape(shape)
            if extracted:
                text_blocks.append(extracted)
        slide_text = clean_text("\n".join(text_blocks))
        if slide_text:
            chunks.append(
                {
                    "id": str(uuid.uuid4()),
                    "doc_id": doc_id,
                    "doc_role": doc_role,
                    "chunk_id": f"{doc_id}_slide{slide_no}",
                    "page": slide_no,
                    "text": slide_text,
                }
            )
    return chunks


def _clean_email_text(raw_text: str) -> str:
    lines = raw_text.split("\n")
    cleaned_lines = []

    header_patterns = r"^(From|To|Cc|Date|Sent|Subject|Reply-To):"
    contact_patterns = r"(\+?\d[\d -]{7,}\d|www\.|https?://|\S+@\S+)"

    for line in lines:
        line_strip = line.strip()
        if not line_strip:
            continue

        if re.search(header_patterns, line_strip, re.IGNORECASE):
            continue

        if re.fullmatch(r"<.*@.*>|\[image\]", line_strip, re.IGNORECASE):
            continue

        if re.search(contact_patterns, line_strip) and len(line_strip.split()) < 5:
            continue

        cleaned_lines.append(line_strip)

    return "\n".join(cleaned_lines).strip()


def extract_eml_as_chunks(eml_path: str, doc_id: str, doc_role: str) -> list:
    with open(eml_path, "rb") as f:
        msg = email.message_from_binary_file(f, policy=policy.default)

    subject = msg.get("Subject", "No Subject")

    body_content = ""
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/html":
                body_content = part.get_content()
                break
        if not body_content:
            for part in msg.walk():
                if part.get_content_type() == "text/plain":
                    body_content = part.get_content()
                    break
    else:
        body_content = msg.get_content()

    if not body_content:
        return []

    doc = fitz.open("html", body_content.encode("utf-8"))
    chunks = []

    for pno in range(doc.page_count):
        page = doc.load_page(pno)
        raw_text = page.get_text().strip()
        if not raw_text:
            continue

        cleaned_text = _clean_email_text(raw_text)
        if not cleaned_text:
            continue

        if pno == 0:
            cleaned_text = f"Subject: {subject}\n\n{cleaned_text}"

        chunks.append(
            {
                "id": str(uuid.uuid4()),
                "doc_id": doc_id,
                "doc_role": doc_role,
                "chunk_id": f"{doc_id}_email_page_{pno+1}",
                "page": pno + 1,
                "text": cleaned_text,
            }
        )

    doc.close()
    return chunks


# ==============================================================================
# TEMPLATE PARSING
# ==============================================================================


def extract_template_sections(template_path: str) -> dict:
    sections = {}
    ext = os.path.splitext(template_path.lower())[1]

    if ext == ".docx":
        doc = Document(template_path)
        for para in doc.paragraphs:
            if para.style.name.startswith("Heading"):
                key = para.text.strip().lower().replace(" ", "_")
                key = re.sub(r"[^a-z0-9_]", "", key)
                if key:
                    sections[key] = ""
        if not sections:
            raise ValueError("No headings found in Word template")

    elif ext in [".pdf", ".ppt", ".pptx"]:
        raw_text = ""
        try:
            if ext == ".pdf":
                doc = fitz.open(template_path)
                for page in doc:
                    raw_text += page.get_text("text") + "\n"
                doc.close()
            elif ext in [".ppt", ".pptx"]:
                prs = Presentation(template_path)
                for slide in prs.slides:
                    for shape in slide.shapes:
                        if hasattr(shape, "text"):
                            raw_text += shape.text + "\n"
        except Exception as e:
            print(f"Warning: Could not read text from template file: {e}")
            raise ValueError(f"Failed to parse template file {template_path}")

        if len(raw_text) > 12000:
            raw_text = raw_text[:12000] + "..."

        system_prompt = (
            "You are a document structure analyzer. "
            "Identify all major section headings in the provided text. "
            "Ignore page numbers, headers, footers, or noise. "
            "Return ONLY a valid JSON list of strings representing the section titles."
        )

        try:
            response = aoai_client.chat.completions.create(
                model=AZURE_OPENAI_DEPLOYMENT_CHAT,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": raw_text},
                ],
                temperature=0.1,
            )
            content = response.choices[0].message.content.strip()
            content = re.sub(
                r"^```json\s*|\s*```$", "", content, flags=re.MULTILINE
            ).strip()
            headers = json.loads(content)

            for h in headers:
                if isinstance(h, str):
                    key = h.strip().lower().replace(" ", "_")
                    key = re.sub(r"[^a-z0-9_]", "", key)
                    sections[key] = ""

        except Exception as e:
            print(
                f"Warning: LLM section extraction failed: {e}. Falling back to empty sections."
            )
            raise ValueError(
                f"Could not automatically determine sections from template: {e}"
            )
    else:
        raise ValueError(f"Unsupported template format: {ext}")

    return sections


# ==============================================================================
# AZURE SEARCH
# ==============================================================================


def upload_chunks_to_search(chunks: List[dict]):
    documents = []

    for c in chunks:
        documents.append(
            {
                "id": c["id"],
                "doc_id": c["doc_id"],
                "doc_role": c["doc_role"],
                "chunk_id": c["chunk_id"],
                "page": c["page"],
                "text": c["text"],
                "text_vector": get_embedding(c["text"]),
            }
        )

    if documents:
        search_client.upload_documents(documents=documents)


def retrieve_terms_from_search(
    query_text: str, valid_sow_doc_ids: List[str], top_k: int = 5
) -> str:

    sow_filter = " or ".join([f"doc_id eq '{d}'" for d in set(valid_sow_doc_ids)])
    combined_filter = f"doc_role eq 'TERMS_SOURCE' and ({sow_filter})"

    query_vector = get_embedding(query_text)

    vector_query = VectorizedQuery(
        vector=query_vector, fields="text_vector", k_nearest_neighbors=top_k
    )

    results = search_client.search(
        search_text="", vector_queries=[vector_query], filter=combined_filter
    )

    seen = set()
    texts = []
    for r in results:
        if r["chunk_id"] in seen:
            continue
        seen.add(r["chunk_id"])
        texts.append(r["text"])

    return "\n\n".join(texts)


# ==============================================================================
# RETRIEVAL
# ==============================================================================

SECTION_QUERY_MAP = {
    "introduction": "statement of work agreement parties engagement",
    "problem_statement": "business problem challenges current state gaps",
    "objectives": "project objectives goals outcomes",
    "scope_of_work": "scope of work activities responsibilities included",
    "out_of_scope": "out of scope exclusions not included",
    "approach": "project approach methodology phases execution",
    "deliverables": "project deliverables outputs artifacts",
    "data_requirements": "data requirements inputs formats access",
    "assumptions": "project assumptions dependencies constraints",
    "roles_and_responsibilities": "roles responsibilities client service provider",
    "payment_terms": "payment terms schedule milestones invoice payment deadline",
}


def retrieve_chunks_per_section(
    template_sections: dict,
    valid_project_doc_ids: List[str],
    valid_sow_doc_ids: List[str],
    top_k: int = 5,
) -> dict:

    SECTION_SOURCE_MAP = {
        "introduction": "TERMS_SOURCE",
    }

    section_chunks = {}

    for section_name in template_sections.keys():
        normalized = section_name.lower().strip()
        normalized = re.sub(r"^\d+\s*[._-]+\s*", "", normalized)
        normalized = re.sub(r"[_\-]+", " ", normalized).strip()
        source = SECTION_SOURCE_MAP.get(normalized, "PROJECT_INPUT")

        cleaned_title = re.sub(r"^\d+\s*[._-]+\s*", "", section_name).strip()
        cleaned_title = re.sub(r"[_\-]+", " ", cleaned_title)
        cleaned_title = re.sub(r"\s+", " ", cleaned_title).strip()
        base_query = SECTION_QUERY_MAP.get(normalized, normalized.replace("_", " "))
        query_text = base_query
        if cleaned_title and cleaned_title.lower() not in base_query:
            query_text = f"{base_query} {cleaned_title}"

        if source == "PROJECT_INPUT":
            doc_filter = " or ".join(
                [f"doc_id eq '{d}'" for d in set(valid_project_doc_ids)]
            )
            combined_filter = f"doc_role eq 'PROJECT_INPUT' and ({doc_filter})"
        else:
            sow_filter = " or ".join(
                [f"doc_id eq '{d}'" for d in set(valid_sow_doc_ids)]
            )
            combined_filter = f"doc_role eq 'TERMS_SOURCE' and ({sow_filter})"

        query_vector = get_embedding(query_text)

        vector_query = VectorizedQuery(
            vector=query_vector, fields="text_vector", k_nearest_neighbors=top_k
        )

        results = search_client.search(
            search_text="", vector_queries=[vector_query], filter=combined_filter
        )

        seen = set()
        chunks = []

        for r in results:
            if r["chunk_id"] in seen:
                continue
            seen.add(r["chunk_id"])

            chunks.append(
                {
                    "doc_role": r["doc_role"],
                    "source_file": r["doc_id"],
                    "chunk_id": r["chunk_id"],
                    "page": r["page"],
                    "text": r["text"],
                }
            )

        section_chunks[section_name] = {
            "source": source,
            "query_used": query_text,
            "chunks": chunks,
        }

    return section_chunks


def load_documents_from_dir(folder: str, doc_role: str) -> List[dict]:
    chunks = []
    for file in os.listdir(folder):
        path = os.path.join(folder, file)
        raw_name, ext = os.path.splitext(file)
        name = f"{doc_role}_{raw_name.lower()}"

        if ext.lower() == ".pdf":
            chunks.extend(extract_pages_as_chunks(path, name, doc_role))
        elif ext.lower() in [".ppt", ".pptx"]:
            chunks.extend(extract_ppt_slides_as_chunks(path, name, doc_role))
        elif ext.lower() == ".eml":
            chunks.extend(extract_eml_as_chunks(path, name, doc_role))

    return chunks


def extract_project_entities(project_chunks: List[dict]) -> dict:
    full_text = "\n".join(c["text"] for c in project_chunks)

    if len(full_text) > 12000:
        full_text = full_text[:12000] + "..."

    system_prompt = (
        "You are an entity extraction expert. Analyze the text and identify the Client Name, "
        "Service Provider (Vendor) Name, and Project Title. "
        "If a value is not found or unclear, return null for that key. "
        "Return ONLY a valid JSON object with keys: CLIENT_NAME, SERVICE_PROVIDER, PROJECT_TITLE."
    )

    try:
        response = aoai_client.chat.completions.create(
            model=AZURE_OPENAI_DEPLOYMENT_CHAT,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": full_text},
            ],
            temperature=0.1,
        )
        content = response.choices[0].message.content.strip()
        content = re.sub(
            r"^```json\s*|\s*```$", "", content, flags=re.MULTILINE
        ).strip()

        entities = json.loads(content)

        final_entities = {
            "CLIENT_NAME": entities.get("CLIENT_NAME"),
            "SERVICE_PROVIDER": entities.get("SERVICE_PROVIDER"),
            "PROJECT_TITLE": entities.get("PROJECT_TITLE"),
        }
        return {k: v for k, v in final_entities.items() if v}

    except Exception as e:
        print(f"  [Error] Entity extraction failed: {e}")
        return {}


# ==============================================================================
# MAIN PIPELINE
# ==============================================================================


def main():
    parser = argparse.ArgumentParser(description="Agentic SOW Generator")
    parser.add_argument("--SOW-dir", required=True)
    parser.add_argument("--Input-dir", required=True)
    parser.add_argument("--output", default="sow.docx")
    args = parser.parse_args()

    template_sections = {}
    template_dir = "uploads/template"
    base_template_path = None

    if os.path.exists(template_dir) and os.listdir(template_dir):
        template_file = os.path.join(template_dir, os.listdir(template_dir)[0])
        template_sections = extract_template_sections(template_file)
        base_template_path = template_file

    # Load Project Input chunks
    input_chunks = load_documents_from_dir(args.Input_dir, "PROJECT_INPUT")

    if not input_chunks:
        raise ValueError("No project input documents uploaded")

    # ===IMAGE EXTRACTION PHASE (Parallel Prep) ===
    print("--- Starting Image Extraction Phase ---")

    desc_path = os.path.join(IMAGE_STORE_PATH, "descriptions.txt")
    if os.path.exists(desc_path):
        os.remove(desc_path)

    if os.path.exists(IMAGE_STORE_PATH):
        for f in os.listdir(IMAGE_STORE_PATH):
            full_path = os.path.join(IMAGE_STORE_PATH, f)
            try:
                if os.path.isfile(full_path):
                    os.remove(full_path)
            except Exception as e:
                print(f"Warning: Could not delete {full_path}: {e}")

    for file in os.listdir(args.Input_dir):
        if file.lower().endswith(".pdf"):
            full_path = os.path.join(args.Input_dir, file)
            process_pdf_to_text(full_path, append_mode=True)
        elif file.lower().endswith((".ppt", ".pptx")):
            full_path = os.path.join(args.Input_dir, file)
            process_ppt_to_text(full_path, append_mode=True)

    os.makedirs("Output", exist_ok=True)

    # === ENTITY EXTRACTION ===
    print("Extracting Project Entities dynamically...")
    project_entities = extract_project_entities(input_chunks)

    json.dump(
        project_entities,
        open("Output/project_entities.json", "w", encoding="utf-8"),
        indent=2,
    )
    print(f"Extracted entities: {project_entities}")

    # === SOW PROCESSING: Extract -> Anonymize -> Chunk ===
    print("Processing SOW files (Extract -> Anonymize -> Chunk)...")
    SOW_chunks = []

    sow_dir = args.SOW_dir
    for file in os.listdir(sow_dir):
        path = os.path.join(sow_dir, file)
        raw_name, ext = os.path.splitext(file)
        doc_id = f"TERMS_SOURCE_{raw_name.lower()}"

        # Step 1: Extract raw pages
        raw_pages = []
        if ext.lower() == ".pdf":
            doc = fitz.open(path)
            for pno in range(doc.page_count):
                page = doc.load_page(pno)
                text = page.get_text("text").strip()
                if text:
                    raw_pages.append({"page": pno + 1, "text": text})
            doc.close()
        elif ext.lower() in [".ppt", ".pptx"]:
            prs = Presentation(path)
            for slide_no, slide in enumerate(prs.slides, start=1):
                text_blocks = []
                for shape in slide.shapes:
                    extracted = extract_all_text_from_shape(shape)
                    if extracted:
                        text_blocks.append(extracted)
                slide_text = clean_text("\n".join(text_blocks))
                if slide_text:
                    raw_pages.append({"page": slide_no, "text": slide_text})

        print(f"  File: {file} - {len(raw_pages)} pages extracted")

        # Step 2 & 3: Anonymize then Chunk
        for page_info in raw_pages:
            raw_text = page_info["text"]
            page_num = page_info["page"]

            anonymized_text = anonymize_text(raw_text)

            if anonymized_text:
                SOW_chunks.append(
                    {
                        "id": str(uuid.uuid4()),
                        "doc_id": doc_id,
                        "doc_role": "TERMS_SOURCE",
                        "chunk_id": f"{doc_id}_page_{page_num}",
                        "page": page_num,
                        "text": anonymized_text,
                    }
                )
                print(f"    Page {page_num}: Anonymized & Chunked")
            else:
                print(f"    Page {page_num}: Skipped (Empty/Noise)")

    if not SOW_chunks:
        raise ValueError("No SOW documents uploaded or all pages were empty/noise")

    # Save anonymized SOW chunks
    json.dump(
        SOW_chunks,
        open("Output/out_sowjson.json", "w", encoding="utf-8"),
        indent=2,
        ensure_ascii=False,
    )
    print(f"Saved {len(SOW_chunks)} anonymized SOW chunks to Output/out_sowjson.json")

    # Save raw input chunks for reference
    json.dump(
        input_chunks,
        open("Output/out.json", "w", encoding="utf-8"),
        indent=2,
        ensure_ascii=False,
    )

    # Upload to Azure Search
    upload_chunks_to_search(SOW_chunks)
    upload_chunks_to_search(input_chunks)

    valid_project_doc_ids = [c["doc_id"] for c in input_chunks]
    valid_sow_doc_ids = [c["doc_id"] for c in SOW_chunks]

    # Retrieve section chunks
    section_trace = retrieve_chunks_per_section(
        template_sections, valid_project_doc_ids, valid_sow_doc_ids
    )

    json.dump(
        section_trace,
        open("Output/section_chunks.json", "w", encoding="utf-8"),
        indent=2,
        ensure_ascii=False,
    )

    # Retrieve terms and project text
    project_text = "\n\n".join(c["text"] for c in input_chunks)
    retrieved_terms = retrieve_terms_from_search(project_text, valid_sow_doc_ids)

    # 1️⃣ Generate formatting rules
    formatting_rules = formatting_rules_agent(
        template_sections,
        project_text,
        retrieved_terms,
    )
    json.dump(
        formatting_rules,
        open("Output/section_formatting_rules.json", "w", encoding="utf-8"),
        indent=2,
        ensure_ascii=False,
    )
    print("✅ Section formatting rules generated")

    # 2️ Draft SOW JSON (Section-by-Section)
    print(" Drafting SOW section-by-section...")

    draft_json = {}
    total_sections = len(template_sections)

    for idx, section_key in enumerate(template_sections.keys(), start=1):
        print(f"  Drafting section {idx}/{total_sections}: {section_key}")

        # Get section-specific data
        section_template = template_sections[section_key]
        formatting_rule = formatting_rules.get(section_key, {})
        section_chunks = section_trace.get(section_key, {})

        # Retry logic for robustness
        for attempt in range(3):
            try:
                section_result = drafting_agent_single_section(
                    section_key=section_key,
                    section_template=section_template,
                    formatting_rule=formatting_rule,
                    section_chunks=section_chunks,
                    retrieved_terms=retrieved_terms,
                    project_inputs=project_text,
                    entities=project_entities,
                )
                # Merge into final draft
                draft_json.update(section_result)
                print(f"   Section '{section_key}' drafted successfully")
                break
            except Exception as e:
                print(f"     Attempt {attempt+1} failed for '{section_key}': {e}")
                if attempt == 2:
                    print(
                        f"     Section '{section_key}' failed after 3 attempts, skipping..."
                    )

    print(" Section-wise drafting completed")

    # 3️⃣ Image Retrieval
    print("--- Running Image Retrieval Matrix ---")
    image_placements = assign_images_to_sections(
        draft_json, os.path.join(IMAGE_STORE_PATH, "descriptions.txt")
    )
    print(f"Final Image Assignments: {len(image_placements)} sections received images.")

    # 4️⃣ Compliance check
    compliance = compliance_agent(draft_json, template_sections)
    compliance_data = json.loads(compliance)
    if not compliance_data.get("approved"):
        raise ValueError(f"SOW failed compliance: {compliance_data}")

    # 5️⃣ FINAL DOCX GENERATION (Section-wise)
    print("--- Generating Final DOCX ---")

    # Initialize output file: Copy template if exists, else create new
    if base_template_path:
        import shutil

        shutil.copy(base_template_path, args.output)
    else:
        Document().save(args.output)

    for index, (key, content) in enumerate(draft_json.items(), 1):
        # Format key (e.g., "1_introduction") into a clean Title
        display_title = key.replace("_", " ").title()
        # Restore dot if it starts with a number (e.g. "1 Introduction" -> "1. Introduction")
        display_title = re.sub(r"^(\d+)\s+(.*)", r"\1. \2", display_title)

        # Check if section has images
        images = image_placements.get(key, [])

        # Handle content type
        if isinstance(content, list):
            content_text = "\n\n".join(str(item) for item in content)
        else:
            content_text = str(content)

        # Special handling for UI section fallback
        key_normalized = key.lower().replace("_", " ")
        is_ui_section = (
            "user interface" in key_normalized
            or "ui design" in key_normalized
            or "screens" in key_normalized
        )

        if is_ui_section and not images:
            content_text = "Currently there are no user interface screenshots provided."

        # Append content to the Word document
        try:
            DocxWriter.append_under_heading_in_newdoc(
                Path(args.output), display_title, content_text, images=images
            )
            print(f"  [Section {index}/{len(draft_json)}] Processed: {display_title}")
        except Exception as e:
            print(f"  [Error] Failed to process section {display_title}: {e}")

    print(f"✅ DOCX Generated at {args.output}")

    # 6️⃣ MARKDOWN GENERATION (For Streamlit UI Preview)
    # This generates a string based on the exact same data used for Docx
    print("--- Generating Markdown Preview ---")
    markdown_lines = []

    for index, (key, content) in enumerate(draft_json.items(), 1):
        # Format key to Title
        display_title = key.replace("_", " ").title()
        display_title = re.sub(r"^(\d+)\s+(.*)", r"\1. \2", display_title)

        # Construct the Markdown structure
        markdown_lines.append(f"## {display_title}")
        markdown_lines.append("")  # Empty line spacing

        # Handle content format
        if isinstance(content, list):
            content_str = "\n".join(str(item) for item in content)
        else:
            content_str = str(content)

        # Normalize key for matching
        key_normalized = key.lower().replace("_", " ")
        is_ui_section = (
            "user interface" in key_normalized
            or "ui design" in key_normalized
            or "screens" in key_normalized
        )

        # If it's a UI section and no images, overwrite content (Logic consistent with Docx)
        if is_ui_section:
            if key in image_placements and len(image_placements[key]) > 0:
                pass  # Keep content
            else:
                content_str = (
                    "Currently there are no user interface screenshots provided."
                )

        markdown_lines.append(content_str)
        markdown_lines.append("")  # Empty line after section

        # ===INSERT IMAGE IF ASSIGNED ===
        if key in image_placements:
            assigned_images = image_placements[key]
            for img in assigned_images:
                # Add Markdown Image Syntax
                markdown_lines.append(f"![Reference Image]({img['path']})")
                markdown_lines.append("")

        markdown_lines.append("")  # Empty line after section block

    # Join everything into a single string and write to the output file
    final_sow_md = "\n".join(markdown_lines)
    # Save alongside the docx
    md_output_path = args.output.replace(".docx", ".md")
    open(md_output_path, "w", encoding="utf-8").write(final_sow_md)
    print(f"✅ Markdown Preview Generated at {md_output_path}")


if __name__ == "__main__":
    main()
