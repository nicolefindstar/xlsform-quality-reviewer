# XLSForm Quality Reviewer

A locally-run quality assurance tool for household survey instruments built with XLSForm. Upload your `.xlsx` form and get a structured report of structural issues, logic errors, simulation-based diagnostics, and survey design recommendations — before the form reaches the field.

> **Data privacy:** Everything runs on your machine. No files or results are ever uploaded to any server.

---

## What It Does

The tool performs three complementary types of analysis:

### 1. Structural Inspection
Examines the form's design statically, checking for:
- Duplicate or invalid variable names
- Broken group / repeat structures (`begin group` without `end group`, etc.)
- Missing metadata fields (`deviceid`, `start`, `end`)
- Integer fields without constraints, constraints without messages
- Outdated syntax (e.g. `${variable}` inside `selected()`)
- Unused choice lists, duplicate choice codes, missing choice labels
- Fields with no `required` rule, disabled/read-only fields
- Variable names that are too long or contain invalid characters

### 2. Simulation-Based Analysis
Runs hundreds of synthetic respondents through the form to expose logic errors that structural inspection alone cannot detect:
- Unreachable questions (never triggered across all simulated paths)
- Rarely-reached questions (reached in fewer than 5% of runs)
- Circular dependencies in relevance conditions
- Deep dependency chains that may cause performance issues
- Skip logic errors (conditions that always evaluate to the same result)

Synthetic respondents are generated with correlated demographic profiles (age, education, location, marital status, income) and realistic response styles (neutral, agreeable, cautious, extreme) to maximise path coverage.

### 3. Design Improvement Suggestions
Analyses the form against applied economics and social science survey methodology best practices, as well as the full [XLSForm specification](https://xlsform.org/en/) and [SurveyCTO documentation](https://docs.surveycto.com/02-designing-forms/). Returns prioritised recommendations across 13 categories:

| Category | What it checks |
|---|---|
| **Reference Integrity** | Broken `${varname}` references (typo detection); `select_multiple` choice names with spaces (data corruption) |
| **Choice Logic** | Exclusive options (Don't know / None / Refuse) selectable alongside other answers without a constraint |
| **Respondent Experience** | Sensitive questions (income, religion, health) missing a 'Prefer not to answer' option |
| **Validation** | Age/year/count/income fields without plausible bounds; required fields without a `required_message` |
| **Skip Logic** | Follow-up questions without `relevant` conditions; `read_only` fields with no `default` or calculation |
| **Performance** | `calculate` fields that should use `once()`; `pulldata()` results used in arithmetic without type conversion; choice lists >200 items (should move to CSV); forms >300 questions; large repeat groups |
| **Question Design** | Likert scales with only a positive pole (acquiescence bias); Likert questions without randomised choice order (order bias); excessively long or double-barreled labels |
| **Enumerator Guidance** | Complex or sensitive questions without `hint` text; open-ended text fields without `multiline` appearance |
| **Data Quality** | Missing `audit` field for interview timing; `geopoint` fields without `capture-accuracy` parameters |
| **Form Metadata** | Missing `form_id`, `form_title`, or `version` in settings sheet; version not in `yyyymmddrr` format |
| **Multilingual** | Incomplete label translations — flags each field missing a translation per language |
| **Survey Flow** | Long question blocks (15+) without any `note` row for context |
| **Efficiency** | Choice lists >20 options without `search`/`autocomplete` appearance; numbered question series (`var_1`, `var_2`…) that should be a repeat group |

Each suggestion includes a plain-language description, a concrete action step, and a ready-to-copy example expression.

### Reporting
- Interactive results tables grouped by issue type — no repeated rows for the same problem
- Severity badges: **Critical / High / Medium / Low** (structural/simulation) and **High / Medium / Low** priority (design suggestions)
- Exportable HTML report with full issue details and fix suggestions

---

## Standards

Structural checks are aligned with two widely-used survey QA frameworks:

| Standard | Organisation |
|---|---|
| [`ietestform`](https://github.com/PovertyAction/high-frequency-checks) | World Bank DIME |
| [`ipacheckscto`](https://github.com/PovertyAction/high-frequency-checks) | IPA |

Design suggestions are grounded in the [XLSForm specification](https://xlsform.org/en/) and SurveyCTO documentation on [core concepts](https://docs.surveycto.com/02-designing-forms/01-core-concepts/), [advanced topics](https://docs.surveycto.com/02-designing-forms/03-advanced-topics/), and [performance optimisation](https://docs.surveycto.com/02-designing-forms/05-performance-and-debugging/).

---

## Quick Start

### Option A — One-click launcher (recommended for non-technical users)

| Platform | File |
|---|---|
| macOS | Double-click `launch_mac.command` |
| Windows | Double-click `launch_windows.bat` |

The launcher will automatically create a virtual environment, install dependencies, and open the app in your browser. No Terminal knowledge required.

> **macOS note:** On first run, right-click `launch_mac.command` → Open → Open, to allow execution.

### Option B — Manual setup

**Requirements:** Python 3.10 or higher

```bash
# 1. Clone the repository
git clone https://github.com/nicolefindstar/xlsform-quality-reviewer.git
cd xlsform-quality-reviewer

# 2. Create and activate a virtual environment
python3 -m venv venv
source venv/bin/activate        # macOS / Linux
venv\Scripts\activate           # Windows

# 3. Install dependencies
pip install -r requirements.txt

# 4. Launch the app
streamlit run app.py
```

The app opens at `http://localhost:8501`.

---

## Usage

1. Launch the app using either method above
2. Upload your XLSForm `.xlsx` file using the sidebar
3. Review **Structural Analysis** — issues run automatically on upload
4. Run **Respondent Path Simulation** to test skip logic across diverse synthetic respondents
5. Review **Design Improvement Suggestions** — prioritised best-practice recommendations
6. Download the HTML report for sharing or archiving

A detailed setup guide (including troubleshooting and FAQ) is available in [`SETUP.md`](SETUP.md) and [`SETUP.pdf`](SETUP.pdf).

---

## File Structure

```
xlsform-quality-reviewer/
├── app.py                 # Main Streamlit application
├── requirements.txt       # Python dependencies
├── launch_mac.command     # One-click launcher for macOS
├── launch_windows.bat     # One-click launcher for Windows
├── make_setup_pdf.py      # Script to regenerate SETUP.pdf
├── SETUP.md               # Setup guide (Markdown)
└── SETUP.pdf              # Setup guide (formatted PDF)
```

---

## Dependencies

| Package | Purpose |
|---|---|
| `streamlit >= 1.35` | Web interface |
| `pandas >= 2.0` | XLSForm parsing and data handling |
| `openpyxl >= 3.1` | Reading `.xlsx` files |
| `numpy >= 1.24` | Numerical operations in simulation |

---

## Troubleshooting

| Problem | Solution |
|---|---|
| `python3: command not found` | Install Python 3.10+ from [python.org](https://www.python.org/downloads/) |
| `ModuleNotFoundError: streamlit` | Activate the virtual environment first (`source venv/bin/activate`) |
| App shows 0 questions after upload | Ensure the file has a `survey` sheet with `type` and `name` columns |
| Port 8501 already in use | Run `streamlit run app.py --server.port 8502` |
| macOS: `launch_mac.command` blocked | Right-click → Open → Open to bypass Gatekeeper |
| Windows: activation policy error | Run `Set-ExecutionPolicy RemoteSigned -Scope CurrentUser` in PowerShell as Administrator |

For the full troubleshooting guide, see [`SETUP.md`](SETUP.md).

---

## License

This project is intended for research use. Please contact the author before redistribution.
