# RD6 Report Generator

Streamlit app that auto-fills the SOCOTEC Arabia RD6 Completion of Works report.

## Setup

```bash
pip install -r requirements.txt
```

## Required Files in Same Folder

| File | Description |
|------|-------------|
| `rd6_template.docx` | The standard RD6 Word template |
| `malath_log.xlsx` | RD6 Master Sheet (Sheet1 + Tuw-Mlth tabs) |
| `rd6_app.py` | Main Streamlit app |
| `rd6_extractor.py` | PDF + Excel data extraction |
| `rd6_generator.py` | DOCX generation logic |

## Run

```bash
streamlit run rd6_app.py
```

## Workflow

1. **Engineer Details** — name, phone, email, degree  
2. **Policy Upload** — upload Malath or Tawuniya IDI policy PDF → auto-extracts fields  
3. **Review & Edit** — verify/correct all extracted data  
4. **Site Visits** — pre-filled from Excel, add/edit/remove rows  
5. **Requirements** — upload client docs (missing ones auto-appear in Section VIII)  
6. **Generate** — download filled RD6 .docx  

## Reference Number Formula

`{ENG_INITIALS}-RD6-{NT/FT}{IDI_NO}-{TAW_POL}-1`  
Example: Yousef Younis + NT + 358273 + 750580 → **YYO-RD6-NT358273-750580-1**

## Excel Columns Used (Sheet1)

| Column | Field |
|--------|-------|
| A: IDI_No | Policy/proposal number (lookup key) |
| B: NT/FT | Prefix type |
| C: Eng | Engineer full name |
| Q: RD0_Ref | RD0 report reference |
| X–AY | Visit refs, dates, inspectors, parts (up to 7 visits) |
| AZ: Taw Pol. | Tawuniya policy number |

## Notes

- **Insulation Certificate is mandatory** — the tool will not proceed to Step 6 without it  
- Documents not uploaded appear automatically as missing in Section VIII and the conclusion  
- All extracted fields are editable in Step 3 before generating  
- Template must be named exactly `rd6_template.docx`
