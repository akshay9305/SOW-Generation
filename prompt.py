# prompt.py

import json

from openai import AzureOpenAI

from key import *

llm = AzureOpenAI(
    api_key=AZURE_OPENAI_API_KEY,
    api_version=AZURE_OPENAI_API_VERSION,
    azure_endpoint=AZURE_OPENAI_ENDPOINT,
)


def formatting_rules_agent(
    template_sections: dict, project_inputs: str, retrieved_terms: str
) -> dict:
    """
    Generate formatting rules for each SOW section.
    Keeps introduction exact; other sections formatting can be inferred.
    """
    prompt = f"""
ROLE:
You are a SOW Section Formatting & Policy Agent.

OBJECTIVE:
For each section in the provided SOW template, define HOW it must be written.

INPUTS:
PROJECT_INPUTS:
{project_inputs}

APPROVED_LEGAL_CLAUSES:
{retrieved_terms}

SOW TEMPLATE SECTIONS:
{json.dumps(template_sections, indent=2)}

OUTPUT FORMAT (STRICT JSON):
Return a JSON object with a key per section, each containing:
- source
- structure
- paragraph_count
- bullet_count
- tone
- notes

MANDATORY RULES:
- "introduction": source="LEGAL_ONLY", structure="paragraphs", paragraph_count=2, tone="legal"
- Project execution sections: source="PROJECT_ONLY"
- Objectives, scope, deliverables, assumptions, roles: bullets
- Problem statement: paragraphs, short
- "deliverables": structure="markdown_table", source="PROJECT_ONLY"
- For "deliverables" table: limit to 5-7 rows total, and keep each cell to one short sentence (max 100 characters) so the table fits on one page in Word.
- "payment_terms": source="LEGAL_ONLY", structure="paragraphs", paragraph_count=2, tone="legal"

RETURN ONLY VALID JSON
"""

    response = llm.chat.completions.create(
        model="gpt-5.2", messages=[{"role": "user", "content": prompt}], temperature=0.1
    )

    # Correctly access the content depending on Azure SDK
    content = response.choices[0].message.content

    try:
        rules_json = json.loads(content)
    except json.JSONDecodeError:
        print("⚠️ Failed to parse JSON from LLM response. Raw response:")
        print(content)
        raise

    return rules_json


def compliance_agent(draft_json, template_sections):
    prompt = f"""
Verify SOW JSON.

Template keys:
{list(template_sections.keys())}

SOW:
{draft_json}

Return JSON ONLY:
{{"missing_sections":[],"empty_sections":[],"approved":true}}
"""
    r = llm.chat.completions.create(
        model="gpt-5.2",
        messages=[{"role": "user", "content": prompt}],
        response_format={"type": "json_object"},
        temperature=0.1,
    )
    return r.choices[0].message.content


def drafting_agent_single_section(
    section_key: str,
    section_template,
    formatting_rule: dict,
    section_chunks: dict,
    retrieved_terms: str,
    project_inputs: str,
    entities: dict,
) -> dict:
    """
    Draft ONLY ONE SOW section.
    Returns: {section_key: content}
    """

    section_lower = section_key.lower()

    # BYPASS LLM for User Interface section - return hardcoded text directly
    if (
        "user interface" in section_lower
        or "user_interface" in section_lower
        or "ui" in section_lower
        or "screens" in section_lower
    ):
        project_name = (
            entities.get("project_name")
            or entities.get("PROJECT_NAME")
            or "the project"
        )
        ui_text = f"User Interface – Below screenshots provides and illustrative of how the user interface will be billed. It doesn't represent the final design. The UI will be updated to incorporate feedback from {project_name} tool users during the PoC tool development. Product will account for all features listed above as a part of this SoW."
        return {section_key: ui_text}

    # Build section-specific instructions based on section name
    special_instructions = ""

    if "introduction" in section_lower:
        special_instructions = """
SPECIAL INSTRUCTION FOR THIS SECTION:
- This section MUST consist of EXACTLY TWO PARAGRAPHS.
- Each paragraph must be written in complete, formal SOW sentences.
- Do NOT use bullet points, numbering, or lists.
- Do NOT include headings or labels inside the section.
- Use ONLY agreement-style legal language derived from APPROVED_LEGAL_CLAUSES.
- Generate TWO coherent paragraphs that together describe the purpose and contractual engagement context.
- Use the specific party names from EXTRACTED ENTITIES if available.
"""
    elif "deliverables" in section_lower:
        special_instructions = """
SPECIAL INSTRUCTION FOR THIS SECTION:
- The "deliverables" section value MUST be a raw Markdown table string ONLY.
- Do NOT include any introductory text (e.g., "The table below shows...").
- Do NOT include concluding text.
- Do NOT include any text before or after the table.
- The table MUST have exactly 2 columns with headers: | Deliverable Category | Deliverables |
- The separator row MUST be: |---|---|
- REQUIRED: Create rows for ONLY these categories (if applicable): "Functional", "Non-Functional", "Backend".
- CRITICAL: Each cell MUST contain actual content (no empty cells).
- CRITICAL: All rows MUST have exactly 2 cells separated by pipes (|).
- CRITICAL: Start the table with the header row, immediately followed by the separator row.
- Fill column values based on PROJECT_INPUTS:
  * Column 1 (Category): Exact category name from list above
  * Column 2 (Deliverables): MERGE deliverable items AND notes/scope into ONE column using serial numbering. Format as: "1. [deliverable item] 2. [another deliverable] 3. [scope/constraint note]" etc. Each numbered point represents either a specific deliverable or a scope note. NO DUPLICATION between items and notes. Keep each numbered point concise and distinct.
- VALIDATION: Ensure NO extra pipes, NO blank rows, NO malformed cells.
- Example output:
| Deliverable Category | Deliverables |
|---|---|
| Functional | 1. User Dashboard with responsive UI design 2. Reporting Module with data export capability 3. Includes accessibility features compliant with WCAG 2.1 4. Tested across multiple browsers and devices |
| Non-Functional | 1. Latency < 200ms performance requirement 2. 99.9% uptime guarantee 3. Data encryption using AES-256 standard 4. 24/7 monitoring with automated alerts |
| Backend | 1. RESTful API with Python 3.9+ support 2. PostgreSQL database with daily backups 3. Cache layer reducing query time by 60% 4. Horizontal scaling capability for future growth |
"""
    elif "payment" in section_lower:
        special_instructions = """
SPECIAL INSTRUCTION FOR THIS SECTION:
- MUST include complete payment information using APPROVED_LEGAL_CLAUSES.
- Structure: 2-3 opening paragraphs (NO markdown tables).
- Include: total cost, invoicing requirements, payment timeline, dispute process, taxes.
- If PROJECT_INPUTS contains specific amount, use it; otherwise use "the Fees".
"""
    elif "out_of_scope" in section_lower or "out of scope" in section_lower:
        special_instructions = """
SPECIAL INSTRUCTION FOR THIS SECTION:
- Extract exclusions/out-of-scope items from PROJECT_INPUTS.
- Structure as bullet points or comma-separated list.
- Do NOT prefix each bullet with "Out of scope:" - the header makes this redundant.
- Remove duplicate or repetitive items; consolidate similar concepts.
"""

    prompt = f"""
ROLE:
You are a Statement of Work (SOW) Drafting Agent.

OBJECTIVE:
Draft ONLY the section "{section_key}".

SECTION TEMPLATE (DO NOT CHANGE KEY):
{json.dumps({section_key: section_template}, indent=2)}

FORMATTING RULE (AUTHORITATIVE):
{json.dumps({section_key: formatting_rule}, indent=2)}

SECTION CONTEXT (RAG):
{json.dumps(section_chunks, indent=2)}

PROJECT_INPUTS:
{project_inputs}

APPROVED_LEGAL_CLAUSES:
{retrieved_terms}

EXTRACTED ENTITIES:
{json.dumps(entities, indent=2)}

{special_instructions}

RULES:
- Generate content ONLY for "{section_key}"
- Follow formatting EXACTLY
- Use only allowed source (LEGAL_ONLY or PROJECT_ONLY as per formatting rule)
- Do NOT infer missing info
- Return VALID JSON with ONLY this key: "{section_key}"

RETURN JSON ONLY
"""

    response = llm.chat.completions.create(
        model="gpt-5.2", messages=[{"role": "user", "content": prompt}], temperature=0.1
    )

    content = response.choices[0].message.content

    try:
        return json.loads(content)
    except json.JSONDecodeError:
        print(f"❌ JSON parse failed for section: {section_key}")
        print(content)
        raise
