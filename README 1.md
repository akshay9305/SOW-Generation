                        SOW Generation Project

## Overview

Intertek SOW ─► TERMS_SOURCE ─► RAG
Crayola PDFs ─► PROJECT_INPUT
                     ↓
                Formatting Agent
                     ↓
                Drafting Agent
                     
                Compliance Agent
                     ↓
                Final SOW

PDF/PPT → chunks → embeddings → Azure AI Search (pgvector equivalent)
                                   ↓
                          semantic + vector retrieval
Azure AI Search becomes your RAG memory layer.


Your index must have these fields:
{
  "name": "sow-generation",
  "fields": [
    { "name": "id", "type": "Edm.String", "key": true },
    { "name": "doc_id", "type": "Edm.String", "filterable": true },
    { "name": "doc_role", "type": "Edm.String", "filterable": true },
    { "name": "chunk_id", "type": "Edm.String" },
    { "name": "content", "type": "Edm.String", "searchable": true },
    {
      "name": "text_vector",
      "type": "Collection(Edm.Single)",
      "searchable": true,
      "vectorSearchDimensions": 1536,
      "vectorSearchProfile": "default"
    }
  ]
}


## High-Level System Overview

Your system is an Agentic, Retrieval-Augmented SOW Generator.

It:
Accepts multiple heterogeneous inputs (PDFs, PPTs)
Separates legal precedent from project requirements
Uses Azure AI Search (vector DB) for clause retrieval
Uses multiple AI agents, each with a single responsibility
Produces auditable, structured, professional SOWs
Exports results as Word
This matches the paper’s Retrieval-Augmented Multi-Agent SOW Drafting System.


## Project Structure

```
SOW_Generation/
├── app.py                     # streamlit ui
├── prompt.py                  # prompt instructions
├── image_extraction.py        # image extraction
├── main.py                    # remaining all main functions
├── key.py                     # key credentials
├── requirements.txt           # Python dependencies
└── README.md                  # This file
```

## FULL SYSTEM FLOWCHART

┌──────────────────────────────┐
│        Streamlit UI          │
│  (app.py – User Interface)   │
└──────────────┬───────────────┘
               │
               ▼
┌──────────────────────────────┐
│        File Upload           │
│  uploads/sample SOW   → historical SOWs        
│  uploads/crayola      → new client docs       
|  Template             → structure only       
└──────────────┬───────────────┘
               │
               ▼
┌──────────────────────────────┐
│  Document Extraction Layer   │
│  - PDF parsing (PyMuPDF)     │
│  - PPT parsing (python-pptx) │
|  - Email parsing (.eml → HTML → text)
│  - Raw page/slide extraction │
└──────────────┬───────────────┘
               │
               ▼
┌──────────────────────────────────────────┐
│        Image Extraction & Gating         │
│        (PARALLEL PIPELINE)               │
│                                          │
│  - PDF → Page images (PyMuPDF)           │
│  - Vision LLM Classifier                 │
│      • UI Screens → KEPT                 │
│      • Diagrams / Stock / Text → DELETED │
│  - Audit logging of decisions            │
└───────────────┬──────────────────────────┘
                │
                ▼
┌──────────────────────────────────────────┐
│  SOW NORMALIZATION (AGENTIC LOOP)        │
│  Writer → Auditor                        |
|                                          │
│Writer Agent                              │
│   - Removes names, dates, money, places  │
│   - Converts to universal legal language │
│                                          │
│Auditor Agent                             │
│   - Detects residual PII / specificity   │
│   - Feeds precise fix instructions       │
│                                          │
│   (Max 3 Writer ↔ Auditor iterations)    │
│  (APPLIED ONLY TO sample SOWs)           │
└──────────────────────┬───────────────────┘
                       │
                       ▼
┌──────────────────────────────────────────┐
│      Chunking & Role Assignment          │
│  - Page/slide-level chunks               │
│  - doc_role tagging:                     │
│      • TERMS_SOURCE (normalized SOW)     │
│      • PROJECT_INPUT (new client docs)   │
└──────────────────────┬───────────────────┘
                       │
                       ▼
┌──────────────────────────────┐
│  Embedding Generation        │
│  - Azure OpenAI embeddings   │
│  - Vector creation           │
└──────────────┬───────────────┘
               │
               ▼
┌──────────────────────────────────────────┐
│        Azure AI Search Index             │
│  - text                                  │
│  - text_vector                           │
│  - doc_id                                │
│  - doc_role                              │
│  - page / chunk_id                       │
└──────────────┬───────────────────────────┘
               │
               ▼
┌──────────────────────────────────────────┐
│      Section-wise RAG Retrieval          │
│  - Vector similarity search              │
│  - Filter by doc_role                    │
│      • introduction → TERMS_SOURCE       │
│      • others sections → PROJECT_INPUT   │
│  - Top-K relevant chunks                 │
└──────────────────────────────────────────|
               │
               ▼
┌──────────────────────────────────────────┐
│     Entity Extraction Agent (LLM)        │
│  - CLIENT_NAME                           │
│  - SERVICE_PROVIDER                      │
│  - PROJECT_TITLE                         │
│  (ONLY from PROJECT_INPUT docs)          │
└──────────────────────┬───────────────────┘
                       │
                       ▼
┌──────────────────────────────────────────┐
│     Formatting Rules Agent (LLM)         │
│  - Section structure                     │
│  - Paragraph / bullet rules / table rules│
│  - Tone & source constraints             │
└──────────────────────┬───────────────────┘
                       │
                       ▼
┌──────────────────────────────────────────┐
│        Drafting Agent (LLM)              │
│  - Inject extracted entities             │
│  - Reuse normalized legal clauses        │
│  - Generate structured SOW JSON          │
│  - NO hardcoded replacements             │
└──────────────────────┬───────────────────┘
                       │
                       ▼

┌──────────────────────────────────────────┐
│     Image-to-Section Assignment          │
│  - Gallery mode                          │
│  - All valid UI images                   │
│    → "User Interface / Screens" section  │
│  - If section missing → images ignored   │
└───────────────┬──────────────────────────┘
                │
                ▼
┌──────────────────────────────────────────┐
│        Compliance Agent (LLM)            │
│  - Section completeness                  │
│  - Structure validation                  │
│  - Approve / Reject                      │
└──────────────────────┬───────────────────┘
                       │
                       ▼
┌──────────────────────────────┐
│  Output Layer                │
│  - Word (.docx)              │
└──────────────────────────────┘


## High-Level Architecture

INPUT FILES + SAMPLE SOWs + TEMPLATE
        ↓
Extraction (text + images)
        ↓
Anonymization (writer ↔ auditor loop)
        ↓
Chunking + Embeddings
        ↓
Azure AI Search (Vector RAG)
        ↓
Formatting Rules Agent (global)
        ↓
Drafting Agent (section-by-section, multiple LLM calls)
        ↓
Compliance Agent
        ↓
Section-wise DOCX + Markdown writer

The key design principle: Each SOW section is drafted independently using its own RAG context and its own LLM call.

## Sample SOW Files (Legal Source)

Extract raw pages
   ↓
anonymize_text()
   ↓
writer agent
   ↓
auditor agent
   ↓
re-write loop if not anonymized properly(max 3)

## Agents

1. Anonymization Writer Agent: Converts raw SOW pages into fully generic legal text. Removes Client / Vendor names, Dates, Costs, Locations, Proprietary project names. Preserves legal intent & structure. Call Frequency: Once per SOW page, Additional calls if auditor flags issues.

2. Anonymization Auditor Agent: Acts as a strict PII compliance checker. Flags: Company names, Dates, Money, People, Internal codes. Call Frequency: Once per anonymized page, Repeated up to 3 times via the write ↔ audit loop.

3. Project Entity Extraction: Extracts: CLIENT_NAME, SERVICE_PROVIDER, PROJECT_TITLE.

4. Formatting Rules Agent: Purpose: Defines how each section must be written. Controls: Source (LEGAL_ONLY / PROJECT_ONLY), Structure (paragraphs, bullets, table), Tone (legal, technical), Paragraph & bullet counts. Role: Acts as a policy / governance agent, Prevents hallucination & formatting drift. Call Frequency: Exactly ONCE per SOW.

5. Drafting Agent (Section-wise Core Agent): Generates content for ONE section ONLY. Uses: Section-specific RAG chunks, Formatting rules, Project inputs, Legal clauses, Extracted entities. Call Frequency: Once per section, Retries up to 3 times per section.

6. Compliance Agent: Verifies Missing sections, Empty sections, Structural compliance, Final gate before document generation. Role: Final auditor, No rewriting, only validation. Call Frequency: Once per run.

7. Vision Image Classification: Vision-based classifier which decides UI Screen. Role: Image quality gate, Deletes non-UI images physically. Call Frequency: Once per extracted image / slide, Can be dozens of calls depending on input.

| # | Agent Name       | Function                   | Calls             |
| - | ---------------- | -------------------------- | ----------------- |
| 1 | Anon Writer      | Rewrite to generic         | Per SOW page      |
| 2 | Anon Auditor     | PII detection              | Per page (looped) |
| 3 | Template Parser  | Section detection          | Once              |
| 4 | Entity Extractor | Client / Vendor / Project  | Once              |
| 5 | Formatting Rules | Section policy             | Once              |
| 6 | Drafting Agent   | Section content generation | **Per section**   |
| 7 | Compliance Agent | Final validation           | Once              |
| 8 | Vision Agent     | UI image classifier        | Per image         |


## Overll Summary

This system automates the creation of a Statement of Work (SOW) using a Retrieval-Augmented, multi-agent AI architecture. Users upload sample SOWs (e.g., Intertek) and project-specific input documents (e.g., Crayola PDFs/PPTs) through a Streamlit UI.

All documents are parsed and converted into text chunks, embedded using Azure OpenAI embeddings, and stored in Azure AI Search as a vector index. During generation, the system retrieves only legally approved clauses from the sample SOWs using semantic search, ensuring accuracy and compliance.

The generation process is handled by three specialized agents:

Drafting Agent – Creates a structured JSON SOW by combining project requirements with retrieved legal clauses.

Compliance Agent – Validates the draft for mandatory legal sections (confidentiality, termination, payment, liability).

Formatting Agent – Converts the approved JSON into a professional, human-readable SOW.

The final SOW is presented in the UI and can be downloaded as Word.
This approach minimizes legal risk, prevents hallucinations, and significantly reduces SOW creation time while maintaining enterprise-grade reliability.

