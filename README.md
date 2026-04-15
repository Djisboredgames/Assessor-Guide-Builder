# CDU Assessor Guide Builder

A tool that generates VET Assessor Guides by pulling unit data from training.gov.au and filling out the CDU template automatically.

## What it does

1. **Enter a unit code** (e.g., `AURHTF102`) and the tool fetches data from training.gov.au
2. **Review & edit** the pulled data — unit details, knowledge evidence, performance evidence, assessment conditions, elements & PCs
3. **Configure assessment tasks** — choose methods (questioning, observation, third party, etc.), set contexts and attempts
4. **Generate a .docx** — complete Assessor Guide with all boilerplate sections filled, mapping matrix scaffolded, and knowledge questions auto-generated from evidence items

## Setup (macOS with Apple Silicon)

```bash
# 1. Clone or download this folder
cd assessor-guide-builder

# 2. Create a virtual environment (recommended)
python3 -m venv venv
source venv/bin/activate

# 3. Install dependencies
pip install -r requirements.txt

# 4. Run the app
streamlit run app.py
```

The app opens in your browser at `http://localhost:8501`.

## What gets auto-filled vs. what you fill in

### Auto-filled from training.gov.au:
- Unit code, title, application/descriptor
- Prerequisites and licensing info
- Knowledge evidence (becomes question scaffolds)
- Performance evidence
- Assessment conditions
- Elements & performance criteria
- Foundation skills
- Mapping matrix structure

### Boilerplate (always the same):
- Sections 5-10 (Reasonable Adjustment, RPL, Feedback, Record Keeping, Resulting Assessments, AI Statement)

### You fill in:
- Benchmark answers for knowledge questions
- Mapping matrix entries (which questions map to which PCs)
- Observation checklist specifics
- Assessment condition responses (how CDU meets them)
- CDU logo (add to the doc in Word after)

## Troubleshooting

**"Couldn't auto-fetch data"** — training.gov.au may be blocking automated requests, or the unit code might use a different URL pattern. Use the manual entry option to paste data directly from the website.

**Missing elements/PCs** — If only assessment requirements downloaded but not the unit file, paste elements from the training.gov.au unit page into the Elements field.

## File structure

```
assessor-guide-builder/
├── app.py              # Streamlit UI (run this)
├── tga_fetcher.py      # training.gov.au data fetching & parsing
├── docx_generator.py   # .docx generation using python-docx
├── requirements.txt    # Python dependencies
└── README.md           # This file
```
