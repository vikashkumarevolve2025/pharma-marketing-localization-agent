# INT Marketing Translation Agent — Streamlit POC

A proof-of-concept for automated promotional material localization pipeline using Claude AI for pharmaceutical marketing translations.

## Features

- **Dual LLM Provider Support**: Choose between AWS Bedrock or direct Anthropic API
- **Multi-language Translation**: Support for 12+ languages (French, German, Portuguese, Japanese, Spanish, etc.)
- **Document Processing**: Handle .txt, .docx, .pdf, and .pptx files
- **Quality Flagging**: Automatic flagging of regulatory claims, clinical data, and low-confidence translations
- **Batch Processing**: Translate multiple files across multiple languages in one job
- **Export**: Download all translations as a ZIP archive

## Prerequisites

### For AWS Bedrock (Recommended)

1. AWS account with Bedrock access
2. AWS CLI configured with a profile (default: `foresight`)
3. Access to Claude Sonnet 4 model in AWS Bedrock (us-east-1 region)

Configure your AWS credentials:
```bash
aws configure --profile foresight
```

### For Anthropic API

1. Anthropic API account
2. API key from [Anthropic Console](https://console.anthropic.com/)

## Installation

1. Install dependencies:
```bash
pip install -r requirements.txt
```

2. Configure AWS Bedrock settings (if using AWS):
   - Edit `app.py` to set your AWS profile name and region:
   ```python
   AWS_PROFILE_NAME = "your-profile-name"  # Default: "foresight"
   AWS_REGION = "us-east-1"                # Your preferred region
   ```

## Usage

1. Run the Streamlit app:
```bash
streamlit run app.py
```

2. In the sidebar:
   - Select **LLM Provider** (AWS Bedrock or Anthropic API)
   - If using Anthropic API, enter your API key
   - Select target languages for translation
   - Review source files in the folder tree

3. Use the tabs:
   - **Source Files**: Upload and preview source documents
   - **Translate**: Configure and run translation jobs
   - **Results**: Review translations and flagged segments
   - **Audit Log**: Track all operations and events

## AWS Bedrock Configuration

The app uses the following AWS Bedrock configuration:

- **Profile**: `foresight` (configurable)
- **Region**: `us-east-1` (configurable)
- **Model**: `us.anthropic.claude-sonnet-4-20250514-v1:0`

Make sure your AWS profile has the necessary permissions:
```json
{
  "Version": "2012-10-17",
  "Statement": [
    {
      "Effect": "Allow",
      "Action": [
        "bedrock:InvokeModel"
      ],
      "Resource": "arn:aws:bedrock:*::foundation-model/anthropic.claude-*"
    }
  ]
}
```

## Project Structure

```
int_mkt_translator/
├── app.py                 # Main Streamlit application
├── requirements.txt       # Python dependencies
├── README.md             # This file
├── source_materials/     # Source documents for translation
│   ├── campaign_decks/
│   ├── product_brochures/
│   └── regulatory_docs/
└── translated/           # Generated translations (by country)
    ├── France/
    ├── Germany/
    └── ...
```

## Translation Features

### COSTAR Prompt Framework

The translation agent uses a structured prompt with:
- **Context**: Pharmaceutical marketing materials
- **Objective**: Accurate localization with content integrity
- **Style**: Natural, fluent marketing copy
- **Tone**: Authored in target language (not word-for-word)
- **Audience**: Marketing teams and regulatory reviewers
- **Response**: Structured JSON with confidence scores and flags

### Automatic Flagging

The system flags segments containing:
- Regulatory claims or contraindications
- Clinical trial data or statistics
- Safety statements
- Low-confidence translations (<85%)
- Brand terms with no local equivalent

## Requirements

- Python 3.8+
- Streamlit 1.32+
- Anthropic API SDK 0.8+ (for direct API)
- Boto3 1.34+ (for AWS Bedrock)
- Document processing libraries (python-docx, PyPDF2, python-pptx)

## License

Internal use only — Pharmaceutical company POC

A proof-of-concept for the automated promotional material localization pipeline described in the PRD.

## Quick Start

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Run the app
streamlit run app.py
```

Then open http://localhost:8501 in your browser.

## Setup

1. Get an Anthropic API key from https://console.anthropic.com
2. Paste it in the sidebar under **Configuration**
3. Select target languages in the sidebar (defaults: fr-FR, de-DE, pt-BR, ja-JP)

## What's Included

### Sample Source Files (source_materials/)
```
source_materials/
  campaign_decks/
    Q3_Campaign_Launch.txt          ← Campaign copy + safety information
  product_brochures/
    CARDIOMAX_Product_Overview.txt  ← Clinical data + dosing information
  regulatory_docs/
    Safety_Communication_Update.txt ← Regulatory safety communication
```

### Output Structure
```
translated/
  France/
    Q3_Campaign_Launch_fr-FR.txt
    CARDIOMAX_Product_Overview_fr-FR.txt
    Safety_Communication_Update_fr-FR.txt
  Germany/
    ...
  Brazil/
    ...
  Japan/
    ...
```

## Features

| Feature | Description |
|---------|-------------|
| Folder monitoring | Scans `source_materials/` for .txt, .docx, .pdf, .pptx |
| Upload new files | Drag & drop via the UI into any folder |
| 12 target locales | fr-FR, de-DE, pt-BR, ja-JP, es-ES, it-IT, nl-NL, ko-KR, zh-CN, pl-PL, ar-SA, ru-RU |
| CO-STAR prompt | Full system prompt with flagging rules applied to every call |
| Regulatory flagging | All safety/clinical segments flagged for mandatory human review |
| Brand term preservation | CARDIOMAX, cardiomaxib, PharmaCorp preserved exactly |
| Audit trail | Timestamped log of every translation job |
| Bulk download | Download all translations as a ZIP |

## Architecture (POC vs Production)

| Component | POC | Production (PRD) |
|-----------|-----|-----------------|
| LLM | Anthropic API (direct) | AWS Bedrock — Claude Sonnet |
| Vector DB | N/A | AWS OpenSearch (k-NN) |
| Embeddings | N/A | Azure OpenAI text-embedding-3-large |
| Source | Local filesystem | Microsoft SharePoint via Graph API |
| Frontend | Streamlit | React Native |
| Backend | Streamlit (monolith) | FastAPI async |
| Notifications | In-app | Push (React Native) |

## Supported File Types

- `.txt` — Plain text
- `.docx` — Word documents (requires `python-docx`)
- `.pdf` — PDF files (requires `PyPDF2`)
- `.pptx` — PowerPoint (requires `python-pptx`)
