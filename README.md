# pharma-marketing-localization-agent
AI-powered pharmaceutical marketing localization agent using AWS Bedrock (Claude Sonnet) + LangGraph. Translates DOCX, PPTX, PDF &amp; TXT into multiple languages, preserving native file formats and flagging regulatory/clinical content for human review.


# 🌐 INT Marketing Translation Agent

<div align="center">

[![Python](https://img.shields.io/badge/Python-3.10%2B-blue?logo=python)](https://python.org)
[![Streamlit](https://img.shields.io/badge/Streamlit-1.x-FF4B4B?logo=streamlit)](https://streamlit.io)
[![AWS Bedrock](https://img.shields.io/badge/AWS-Bedrock-FF9900?logo=amazonaws)](https://aws.amazon.com/bedrock/)
[![Claude Sonnet](https://img.shields.io/badge/Claude-Sonnet%204-6B46C1?logo=anthropic)](https://anthropic.com)
[![LangGraph](https://img.shields.io/badge/LangGraph-0.2%2B-00A67E)](https://github.com/langchain-ai/langgraph)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow)](LICENSE)

**AI-powered pharmaceutical marketing localization agent.**
Translates DOCX · PPTX · PDF · TXT into 12 languages — preserving native file formats,
flagging regulatory content, and maintaining brand/glossary consistency.

</div>

---

## 📋 Table of Contents

- [Overview](#-overview)
- [Features](#-features)
- [Architecture](#-architecture)
- [Tech Stack](#-tech-stack)
- [Project Structure](#-project-structure)
- [Getting Started](#-getting-started)
- [Configuration](#-configuration)
- [Supported Languages](#-supported-languages)
- [Translation Pipelines](#-translation-pipelines)
- [File Format Support](#-file-format-support)
- [Flagging Categories](#-flagging-categories)
- [Roadmap](#-roadmap)

---

## 🔍 Overview

The **INT Marketing Translation Agent** is a Streamlit-based POC that automates the localization of pharmaceutical promotional materials — campaign decks, product brochures, and regulatory communications — into multiple target languages.

It is designed for:
- **Regulatory compliance** — flags clinical data, contraindications, and safety statements for mandatory human review
- **Brand consistency** — preserves product/brand names (e.g., CARDIOMAX, cardiomaxib) exactly as written
- **Format fidelity** — output files match the input format (DOCX → translated DOCX, PPTX → translated PPTX)

---

## ✨ Features

| Feature | Description |
|---|---|
| 🔄 **Multi-format I/O** | Translates DOCX, PPTX, PDF, TXT; output preserves original format |
| 🌍 **12 Locales** | FR, DE, PT-BR, JA, ES, IT, NL, KO, ZH, PL, AR, RU |
| 🤖 **3 Agent Modes** | LangGraph pipeline · Tool-Use Agent · Direct single-shot |
| ⚕️ **Regulatory Flagging** | Auto-flags clinical data, regulatory claims, low-confidence segments |
| 📊 **Confidence Scoring** | Per-translation confidence score (0.0–1.0) |
| 🔤 **Glossary Enforcement** | Brand/product terms preserved across all translations |
| 📝 **Audit Trail** | Full timestamped log of all translation jobs |
| ☁️ **Dual Provider** | AWS Bedrock or direct Anthropic API |
| 📦 **Batch Export** | ZIP download of all translated files |
| 👁️ **Vision Support** | Extracts and translates text embedded in PPTX images |

---

## 🏗️ Architecture

    ┌─────────────────────────────────────────────────────────────┐
    │                    Streamlit UI (app.py)                     │
    │  ┌──────────┐  ┌───────────┐  ┌──────────┐  ┌──────────┐  │
    │  │ Source   │  │ Translate │  │ Results  │  │ Audit    │  │
    │  │ Files    │  │ Tab       │  │ Tab      │  │ Log      │  │
    │  └──────────┘  └─────┬─────┘  └──────────┘  └──────────┘  │
    └────────────────────── │ ────────────────────────────────────┘
                            │
              ┌─────────────▼──────────────┐
              │     Translation Router      │
              └──────┬──────────┬──────────┘
                     │          │
        ┌────────────▼──┐  ┌────▼───────────────────────┐
        │  Direct Mode  │  │    LangGraph Pipeline       │
        │  (single-shot)│  │  1. Translate Node          │
        └───────────────┘  │  2. Regulatory Review Node  │
                           │  3. QA Check Node           │
        ┌──────────────┐   │  (auto-revision loop)       │
        │ Tool-Use     │   └─────────────────────────────┘
        │ Agent Mode   │
        └──────────────┘
                     │
        ┌────────────▼────────────┐
        │  AWS Bedrock / Anthropic │
        │  Claude Sonnet 4         │
        └────────────┬────────────┘
                     │
        ┌────────────▼────────────┐
        │  Segment Extractor      │  <- PPTX/DOCX paragraph & table cells
        │  Native File Rebuilder  │  <- rebuild_pptx / rebuild_docx
        └────────────┬────────────┘
                     │
        ┌────────────▼────────────┐
        │   translated/           │
        │   ├── France/           │
        │   ├── Germany/          │
        │   ├── Brazil/           │
        │   └── Japan/ ...        │
        └─────────────────────────┘

---

## 🛠️ Tech Stack

| Layer | Technology |
|---|---|
| **UI** | Streamlit |
| **LLM** | Claude Sonnet 4 via AWS Bedrock or Anthropic API |
| **Orchestration** | LangGraph 0.2+ |
| **DOCX I/O** | python-docx |
| **PPTX I/O** | python-pptx |
| **PDF I/O** | PyPDF2 |
| **Cloud** | AWS Boto3 (Bedrock Runtime) |
| **Language** | Python 3.10+ |

---

## 📁 Project Structure

    int_mkt_translation_agent_poc/
    │
    ├── int_mkt_translator/
    │   ├── app.py                  # Main Streamlit application
    │   ├── requirements.txt        # Python dependencies
    │   ├── README.md
    │   │
    │   ├── source_materials/       # Input documents (organized by type)
    │   │   ├── campaign_decks/
    │   │   ├── product_brochures/
    │   │   ├── regulatory_docs/
    │   │   └── other/
    │   │
    │   └── translated/             # Output files (organized by country)
    │       ├── France/
    │       ├── Germany/
    │       ├── Brazil/
    │       └── Japan/
    │
    └── multilingualenv/            # Python virtual environment

---

## 🚀 Getting Started

### Prerequisites

- Python 3.10+
- AWS account with Bedrock access (Claude Sonnet enabled in us-east-1)
  OR Anthropic API key

### Installation

    # 1. Clone the repository
    git clone https://github.com/<your-username>/int-mkt-translation-agent.git
    cd int-mkt-translation-agent/int_mkt_translator

    # 2. Create and activate virtual environment
    python -m venv venv
    source venv/bin/activate        # macOS/Linux
    venv\Scripts\activate           # Windows

    # 3. Install dependencies
    pip install -r requirements.txt

    # 4. Run the app
    streamlit run app.py

### requirements.txt

    streamlit
    anthropic
    boto3
    python-docx
    python-pptx
    PyPDF2
    langgraph>=0.2.0

---

## ⚙️ Configuration

### AWS Bedrock (default)

Update these constants in app.py:

    AWS_PROFILE_NAME = "your-aws-profile"
    AWS_REGION       = "us-east-1"
    AWS_MODEL_ID     = "us.anthropic.claude-sonnet-4-20250514-v1:0"

Ensure the profile has bedrock:InvokeModel permissions.

### Anthropic API

Select **Anthropic API** from the sidebar dropdown and paste your sk-ant-... key.

---

## 🌍 Supported Languages

| Code | Language | Country |
|---|---|---|
| fr-FR | 🇫🇷 French | France |
| de-DE | 🇩🇪 German | Germany |
| pt-BR | 🇧🇷 Portuguese | Brazil |
| ja-JP | 🇯🇵 Japanese | Japan |
| es-ES | 🇪🇸 Spanish | Spain |
| it-IT | 🇮🇹 Italian | Italy |
| nl-NL | 🇳🇱 Dutch | Netherlands |
| ko-KR | 🇰🇷 Korean | South Korea |
| zh-CN | 🇨🇳 Chinese | China |
| pl-PL | 🇵🇱 Polish | Poland |
| ar-SA | 🇸🇦 Arabic | Saudi Arabia |
| ru-RU | 🇷🇺 Russian | Russia |

---

## 🤖 Translation Pipelines

### 1. LangGraph (Recommended)

3-node agentic pipeline with auto-revision loop:

    Translate Node  →  Regulatory Review Node  →  QA Check Node
          ↑_______________________revision if needed____________|

- **Translate Node** — Produces fluent target-language translation
- **Regulatory Review Node** — Flags clinical data, safety statements, regulatory claims
- **QA Check Node** — Final quality gate; triggers revision if issues found (max 2 passes)

### 2. Tool-Use Agent

Single agent with structured tool calls:
- flag_segment — marks regulatory/clinical content
- apply_glossary_term — records preserved brand terms
- submit_translation — finalizes output

Supports **vision** (reads text embedded in PPTX slide images).

### 3. Direct

Single-shot JSON response. Fastest mode, text-only.

---

## 📄 File Format Support

| Input | Extraction | Output |
|---|---|---|
| .txt | Full text | .txt |
| .docx | Paragraphs + table cells | .docx (formatting preserved) |
| .pptx | Text frames + table cells + images | .pptx (layout preserved) |
| .pdf | Text via PyPDF2 | .txt fallback |

> **Native format rebuild**: For DOCX and PPTX, the agent translates at paragraph/cell level
> and writes back into the original file structure — preserving fonts, styles, layouts, and images.

---

## ⚠️ Flagging Categories

Segments are automatically flagged for **mandatory human review**:

| Flag | Trigger |
|---|---|
| regulatory_claim | Contraindications, safety statements |
| clinical_data | Trial names, statistics, p-values, CIs |
| low_confidence | Translation confidence < 85% |
| brand_term_missing | Brand term with no clear local equivalent |

---

## 🗺️ Roadmap

- [ ] Azure OpenAI provider support
- [ ] Translation memory / glossary management UI
- [ ] PDF native output (ReportLab)
- [ ] Side-by-side diff viewer (source vs translation)
- [ ] Webhook / REST API mode
- [ ] XLIFF export for CAT tool integration

---

<div align="center">
  <sub>Built as a POC · Not for production regulatory submission without human review</sub>
</div>
