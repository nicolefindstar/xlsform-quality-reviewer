"""
XLSForm Reviewer — Comprehensive quality assurance tool for household survey instruments.
Consolidates structural and simulation checks from World Bank ietestform and IPA ipacheckscto.
"""

import streamlit as st
import pandas as pd
import numpy as np
import random
import re
import html as html_module
from datetime import datetime
from collections import Counter, defaultdict

# ── PAGE CONFIG ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="XLSForm Quality Reviewer",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
)
st.markdown("""
<style>
  [data-testid="stDownloadButton"] button {
    background:#6a9cc8;color:white;border:none;width:100%;
    font-weight:600;padding:.65rem 1rem;border-radius:8px;
  }
  [data-testid="stDownloadButton"] button:hover{background:#5a8cb8}
  /* ── Grouped issue table ─────────────────────────────────────── */
  .xit-table{width:100%;border-collapse:collapse;font-size:.84rem;margin-top:.5rem}
  .xit-table th{background:#f1f5f9;text-align:left;padding:.5rem .75rem;font-weight:600;
    color:#64748b;text-transform:uppercase;letter-spacing:.04em;font-size:.7rem;
    border-bottom:2px solid #e2e8f0}
  .xit-table td{padding:.65rem .75rem;border-bottom:1px solid #e2e8f0;vertical-align:top}
  .xit-table tr:last-child td{border-bottom:none}
  .xit-table tr:hover td{background:#f8fafc}
  .xit-badge{display:inline-block;padding:.2rem .6rem;border-radius:20px;
    font-size:.7rem;font-weight:700;white-space:nowrap}
  .xit-src{display:inline-block;font-size:.65rem;font-weight:600;
    padding:.1rem .4rem;border-radius:4px;margin-top:.3rem}
  .xit-src-s{background:#e8f2fc;color:#6a9cc8}
  .xit-src-m{background:#f2eefc;color:#9888c8}
  .xit-count{text-align:center;font-weight:800;font-size:1.05rem}
  .xit-vars code{font-family:"SF Mono","Fira Code",monospace;background:#f1f5f9;
    padding:.1rem .35rem;border-radius:3px;font-size:.78rem;margin:1px 2px;
    display:inline-block}
  .xit-fix{color:#64748b;font-size:.81rem;margin-top:.4rem;line-height:1.5}
  .xit-more{color:#94a3b8;font-style:italic;font-size:.78rem}
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# XLSFORM PARSER
# ══════════════════════════════════════════════════════════════════════════════

class XLSFormParser:
    """Parse an XLSForm .xlsx into structured question and choices data."""

    STRUCTURAL = {
        "begin group", "end group", "begin repeat", "end repeat",
        "begin_group", "end_group", "begin_repeat", "end_repeat",
    }

    def __init__(self, file_obj):
        self.survey_df: pd.DataFrame  = None
        self.choices_df: pd.DataFrame = None
        self.choices_dict: dict       = {}   # list_name → [choice_names]
        self.questions: list          = []
        self.settings: dict           = {}   # settings sheet key→value
        self.label_langs: list        = []   # detected language suffixes e.g. ["English","French"]
        self._parse(file_obj)

    def _parse(self, file_obj):
        xl = pd.ExcelFile(file_obj)
        sheet_map = {n.lower(): n for n in xl.sheet_names}

        if "survey" in sheet_map:
            df = xl.parse(sheet_map["survey"])
            df.columns = [c.strip().lower() for c in df.columns]
            self.survey_df = df
            # Detect multilingual label columns
            self.label_langs = [
                col.split(":", 1)[1].strip()
                for col in df.columns
                if col.startswith("label:") and ":" in col
            ]

        if "choices" in sheet_map:
            df = xl.parse(sheet_map["choices"])
            df.columns = [c.strip().lower() for c in df.columns]
            self.choices_df = df
            self._build_choices()

        if "settings" in sheet_map:
            df = xl.parse(sheet_map["settings"])
            df.columns = [c.strip().lower() for c in df.columns]
            for _, row in df.iterrows():
                for col in df.columns:
                    v = self._val(row, col)
                    if v:
                        self.settings[col] = v

        if self.survey_df is not None:
            self._build_questions()

    def _build_choices(self):
        for _, row in self.choices_df.iterrows():
            ln   = self._val(row, "list_name")
            name = self._val(row, "name")
            if ln and name:
                self.choices_dict.setdefault(ln, []).append(name)

    def _build_questions(self):
        for _, row in self.survey_df.iterrows():
            q_type = self._val(row, "type")
            q_name = self._val(row, "name")
            if not q_type or not q_name:
                continue

            label = ""
            for col in self.survey_df.columns:
                if col == "label" or col.startswith("label:"):
                    v = self._val(row, col)
                    if v:
                        label = v
                        break

            q = {
                "type":               q_type,
                "name":               q_name,
                "label":              label,
                "relevant":           self._val(row, "relevant") or self._val(row, "relevance"),
                "constraint":         self._val(row, "constraint"),
                "constraint_message": self._first_val(row, "constraint_message"),
                "required":           self._val(row, "required"),
                "required_message":   self._first_val(row, "required_message"),
                "hint":               self._first_val(row, "hint"),
                "calculation":        self._val(row, "calculation"),
                "default":            self._val(row, "default"),
                "appearance":         self._val(row, "appearance"),
                "disabled":           self._val(row, "disabled"),
                "read_only":          self._val(row, "read_only") or self._val(row, "readonly"),
                "parameters":         self._val(row, "parameters"),
                "choices":            [],
                "list_name":          None,
                "_label_cols":        {
                    col: self._val(row, col)
                    for col in (self.survey_df.columns if self.survey_df is not None else [])
                    if col == "label" or col.startswith("label:")
                },
            }

            base = q_type.split()[0]
            if base in ("select_one", "select_multiple"):
                parts = q_type.split()
                if len(parts) >= 2:
                    ln = parts[1]
                    q["list_name"] = ln
                    q["choices"]   = self.choices_dict.get(ln, [])

            self.questions.append(q)

    @staticmethod
    def _val(row, col):
        v = row.get(col, None)
        if v is None or (isinstance(v, float) and np.isnan(v)):
            return ""
        return str(v).strip()

    def _first_val(self, row, col_prefix):
        """Return first non-empty value from col_prefix or col_prefix:* columns."""
        if self.survey_df is None:
            return ""
        for col in self.survey_df.columns:
            if col == col_prefix or col.startswith(col_prefix + ":"):
                v = row.get(col, None)
                if v is not None and not (isinstance(v, float) and np.isnan(v)):
                    s = str(v).strip()
                    if s and s.lower() != "nan":
                        return s
        return ""

    def answerable_questions(self):
        return [q for q in self.questions if q["type"].split()[0] not in self.STRUCTURAL]

    def form_stats(self):
        types = Counter(q["type"].split()[0] for q in self.questions)
        aq    = self.answerable_questions()
        return {
            "total_rows":         len(self.questions),
            "answerable":         len(aq),
            "types":              dict(types),
            "skip_logic_count":   sum(1 for q in aq if q["relevant"]),
            "constraint_count":   sum(1 for q in aq if q["constraint"]),
            "required_count":     sum(1 for q in aq if q["required"].lower() in ("yes", "true", "1")),
            "choice_lists":       len(self.choices_dict),
        }


# ══════════════════════════════════════════════════════════════════════════════
# CONDITION EVALUATOR
# ══════════════════════════════════════════════════════════════════════════════

class ConditionEvaluator:
    """Safely evaluate XLSForm XPath-like expressions."""

    def __init__(self, responses: dict):
        self.responses = responses

    def evaluate(self, condition: str, current_val=None) -> bool:
        if not condition or condition.strip() in ("", "nan", "true", "1", "True"):
            return True
        if condition.strip() in ("false", "0", "False"):
            return False
        try:
            expr   = self._to_python(condition, current_val)
            result = eval(expr, {"__builtins__": {}}, {
                "True": True, "False": False,
                "selected":       self._selected,
                "count_selected": self._count_selected,
                "string_length":  lambda s: len(str(s)),
                "__xlsnot":       lambda x: not x,
                # XPath if(cond, yes, no) — must rename because `if` is a keyword
                "__xlsif":        lambda c, a, b: a if c else b,
                "int": self._safe_int, "float": self._safe_float,
                "number":         self._safe_float,
                "string":         lambda v: "" if v is None else str(v),
                "boolean":        bool,
                "str": str, "len": len, "abs": abs, "min": min, "max": max,
                "floor":          lambda v: int(self._safe_float(v)),
                "ceiling":        lambda v: -int(-self._safe_float(v)),
                "round":          lambda v, d=0: round(self._safe_float(v), int(d)),
                "concat":         lambda *a: "".join(str(x) for x in a),
                "contains":       lambda s, sub: str(sub).strip("'\"") in str(s),
                "starts_with":    lambda s, p: str(s).startswith(str(p).strip("'\"")),
                "ends_with":      lambda s, p: str(s).endswith(str(p).strip("'\"")),
                "selected_at":    lambda s, i: (str(s).split()[int(i)] if str(s).split() else ""),
                "coalesce":       lambda *a: next((x for x in a if x not in (None, "", "nan")), ""),
                "__today":        "2024-06-01",
                "__now":          "2024-06-01T08:00:00",
            })
            return bool(result)
        except Exception:
            return True  # default to visible on parse error

    def _to_python(self, expr: str, current_val=None) -> str:
        # Strip SurveyCTO-specific namespace prefixes (jr:itext, jr:choice-name, etc.)
        expr = re.sub(r"\bjr:[a-zA-Z\-]+\s*\([^)]*\)", repr(""), expr)
        # Strip pulldata() — external CSV lookups, return empty string
        expr = re.sub(r"\bpulldata\s*\([^)]*\)", repr(""), expr)

        # Replace . (current-value reference in constraints)
        if current_val is not None:
            expr = re.sub(r"(?<![.\w])\.(?![.\w])", repr(str(current_val)), expr)

        # Replace ${varname} references — return numeric literal when possible
        # so that comparisons like ${age} > 18 work correctly
        def sub_var(m):
            val = self.responses.get(m.group(1))
            if val is None:
                return repr("")
            if isinstance(val, list):
                return repr(" ".join(str(v) for v in val))
            s = str(val)
            try:
                n = float(s)
                return str(int(n)) if n == int(n) else str(n)
            except (ValueError, OverflowError):
                return repr(s)
        expr = re.sub(r"\$\{([^}]+)\}", sub_var, expr)

        # XPath if(cond, yes, no) — rename before other substitutions
        expr = re.sub(r"\bif\s*\(", "__xlsif(", expr)

        # true() / false() XPath functions
        expr = re.sub(r"\btrue\s*\(\s*\)",  "True",  expr)
        expr = re.sub(r"\bfalse\s*\(\s*\)", "False", expr)

        # today() / now()
        expr = re.sub(r"\btoday\s*\(\s*\)", "__today", expr)
        expr = re.sub(r"\bnow\s*\(\s*\)",   "__now",   expr)

        # XPath functions → Python callables
        expr = re.sub(r"count-selected\s*\(",  "count_selected(",  expr)
        expr = re.sub(r"string-length\s*\(",   "string_length(",   expr)
        expr = re.sub(r"selected-at\s*\(",     "selected_at(",     expr)
        expr = re.sub(r"starts-with\s*\(",     "starts_with(",     expr)
        expr = re.sub(r"ends-with\s*\(",       "ends_with(",       expr)

        # not() — rename to avoid keyword conflict
        expr = re.sub(r"\bnot\s*\(", "__xlsnot(", expr)

        # Logical operators
        expr = re.sub(r"\band\b", " and ", expr)
        expr = re.sub(r"\bor\b",  " or ",  expr)

        # = → ==  (skip !=, <=, >=, ==)
        expr = re.sub(r"(?<![!<>=])=(?!=)", "==", expr)

        # Arithmetic
        expr = re.sub(r"\bdiv\b", "/", expr)
        expr = re.sub(r"\bmod\b", "%", expr)

        return expr

    @staticmethod
    def _selected(value_str: str, choice) -> bool:
        if not value_str:
            return False
        return str(choice).strip("'\"") in str(value_str).split()

    @staticmethod
    def _count_selected(value_str: str) -> int:
        return len(str(value_str).split()) if value_str else 0

    @staticmethod
    def _safe_int(v):
        try:    return int(float(str(v)))
        except: return 0

    @staticmethod
    def _safe_float(v):
        try:    return float(str(v))
        except: return 0.0


# ══════════════════════════════════════════════════════════════════════════════
# RESPONDENT PROFILE  — correlated demographics + response style
# ══════════════════════════════════════════════════════════════════════════════

class RespondentProfile:
    LOCATIONS  = ["rural", "urban", "peri-urban", "semi-urban"]
    EDUCATION  = ["none", "primary", "secondary", "tertiary"]
    MARITAL    = ["single", "married", "widowed", "divorced", "separated"]
    INCOME     = ["low", "medium", "high"]
    OCCUPATION = ["farmer", "trader", "teacher", "casual_laborer",
                  "unemployed", "student", "government_worker"]
    # Response style archetypes (psychological survey bias patterns)
    STYLES     = ["neutral", "agreeable", "cautious", "extreme"]

    def __init__(self, rng: random.Random):
        # Gender fixed to female — women's business study
        self.gender = "female"

        # Age: triangular distribution peaking at 32 (working-age women)
        self.age = int(rng.triangular(18, 65, 32))

        # Location: more rural than urban in developing-country household surveys
        self.location = rng.choices(
            self.LOCATIONS, weights=[45, 30, 15, 10]
        )[0]

        # Education: depends on location; younger cohorts slightly more educated
        self.education = self._draw_education(rng)

        # Marital status: depends on age
        self.marital_status = self._draw_marital(rng)

        # Household size: depends on marital status
        self.household_size = self._draw_hh_size(rng)

        # Children: depends on age, marital status, household size
        self.num_children = self._draw_children(rng)

        # Occupation: depends on education + location
        self.occupation = self._draw_occupation(rng)

        # Income: depends on education + occupation
        self.income_level = self._draw_income(rng)

        # Asset ownership: higher probability in urban / higher income
        urban = self.location == "urban"
        rich  = self.income_level == "high"
        self.has_phone       = rng.random() < (0.90 if urban else 0.60)
        self.has_electricity = rng.random() < (0.92 if urban else 0.48)
        self.has_land        = rng.random() < (0.20 if urban else 0.72)

        # Consent: ~85 % agree to participate
        self.consent = rng.random() < 0.85

        # Response style archetype drives Likert answer patterns
        self.response_style = rng.choices(
            self.STYLES, weights=[40, 30, 20, 10]
        )[0]

    # ── correlated draw helpers ───────────────────────────────────────────────

    def _draw_education(self, rng):
        if self.location == "urban":
            w = [5, 25, 45, 25]
        elif self.location in ("peri-urban", "semi-urban"):
            w = [10, 35, 40, 15]
        else:                        # rural
            w = [25, 45, 25, 5]
        if self.age < 30:            # younger cohorts: less "none", more tertiary
            w = [max(0, w[0] - 5), w[1], w[2], w[3] + 5]
        return rng.choices(self.EDUCATION, weights=w)[0]

    def _draw_marital(self, rng):
        a = self.age
        if   a < 22: w = [70, 20,  1,  5,  4]
        elif a < 30: w = [30, 55,  2,  8,  5]
        elif a < 45: w = [15, 60,  5, 12,  8]
        else:        w = [10, 45, 20, 15, 10]
        return rng.choices(self.MARITAL, weights=w)[0]

    def _draw_hh_size(self, rng):
        if self.marital_status == "married":
            return max(2, int(rng.triangular(2, 12, 5)))
        elif self.marital_status in ("divorced", "separated", "widowed"):
            return max(1, int(rng.triangular(1, 8, 3)))
        else:
            return max(1, int(rng.triangular(1, 6, 2)))

    def _draw_children(self, rng):
        if self.marital_status == "single" and self.age < 22:
            return 0
        max_k = max(0, min(self.household_size - 1, 8))
        if   self.age < 25: return rng.randint(0, min(2, max_k))
        elif self.age < 35: return rng.randint(0, min(4, max_k))
        else:               return rng.randint(min(1, max_k), max(1, max_k))

    def _draw_occupation(self, rng):
        if self.education == "tertiary":
            w = [3, 15, 30, 5, 5, 5, 37]
        elif self.location == "rural":
            w = [50, 20,  5, 15, 7, 2, 1]
        else:
            w = [5,  45, 10, 20, 12, 5, 3]
        return rng.choices(self.OCCUPATION, weights=w)[0]

    def _draw_income(self, rng):
        if self.education == "tertiary" or self.occupation in ("teacher", "government_worker"):
            w = [15, 50, 35]
        elif self.education == "none" or self.occupation in ("farmer", "unemployed"):
            w = [65, 30,  5]
        else:
            w = [35, 50, 15]
        return rng.choices(self.INCOME, weights=w)[0]

    def as_dict(self):
        # Exclude private attributes (prefixed _)
        return {k: v for k, v in self.__dict__.items() if not k.startswith("_")}


# ══════════════════════════════════════════════════════════════════════════════
# RESPONSE GENERATOR  — context-aware, profile-consistent answers
# ══════════════════════════════════════════════════════════════════════════════

class ResponseGenerator:
    _KEYWORD_MAP = {
        "age": "age", "sex": "gender", "gender": "gender",
        "hh_size": "household_size", "hhsize": "household_size",
        "household_size": "household_size", "num_hh": "household_size",
        "members": "household_size", "location": "location",
        "zone": "location", "area": "location", "region": "location",
        "education": "education", "edu": "education",
        "marital": "marital_status", "children": "num_children",
        "num_child": "num_children", "income": "income_level",
        "phone": "has_phone", "electricity": "has_electricity",
        "occupation": "occupation", "land": "has_land",
    }

    # Keywords that signal a Likert-type attitudinal question
    _LIKERT_KW = {
        "satisf","happy","stress","worry","feel","difficult","enjoy","sad",
        "nervous","positive","negative","content","depress","anxiety","afraid",
        "anger","calm","lonely","tired","hope","cope","support","conflict",
        "safe","trust","proud","agree","disagree","often","never","always",
        "wellbeing","burden","pressure","relax","pleasure","upset","bother",
    }

    # Keywords that signal a time/duration question
    _TIME_KW = {"hour","hrs","minute","min","duration","spend","time_use"}

    # Region-appropriate female names (Sub-Saharan Africa / South Asia)
    _FIRST = ["Amina","Fatima","Aisha","Zainab","Halima","Mariam","Hadiza",
              "Blessing","Chioma","Ngozi","Adaeze","Funke","Yetunde",
              "Abena","Akosua","Ama","Efua","Esi","Gifty","Adwoa",
              "Priya","Sunita","Kavita","Meera","Lakshmi"]
    _LAST  = ["Diallo","Mensah","Otieno","Yusuf","Nkosi","Osei","Adeyemi",
              "Ibrahim","Traoré","Ouédraogo","Koné","Coulibaly",
              "Bah","Camara","Touré","Keïta","Cissé","Patel","Sharma"]

    def __init__(self, profile: RespondentProfile, rng: random.Random):
        self.profile   = profile
        self.rng       = rng
        self._time_used = 0   # waking minutes already allocated across time-use Qs

    def generate(self, question: dict, responses: dict):
        q_type  = question["type"]
        q_name  = question["name"]
        choices = question.get("choices", [])
        base    = q_type.split()[0]
        n       = q_name.lower()

        if base in XLSFormParser.STRUCTURAL or q_type in ("note", "calculate"):
            return None

        # ── Consent: the single most important branching gate ─────────────────
        if "consent" in n:
            target = "1" if self.profile.consent else "0"
            if choices:
                # Match exact code first, then fall back to first/last choice
                for c in choices:
                    if c == target:
                        return c
                return choices[0] if self.profile.consent else choices[-1]
            return target

        # ── Profile keyword matching ──────────────────────────────────────────
        pv = self._from_profile(q_name, choices)
        if pv is not None:
            return pv

        # ── Type-based generation ─────────────────────────────────────────────
        if q_type == "integer":   return self._gen_int(n)
        if q_type == "decimal":   return self._gen_decimal(n)
        if q_type == "text":      return self._gen_text(n)

        if base == "select_one":
            return self._gen_select_one(n, choices)

        if base == "select_multiple":
            return self._gen_select_multiple(choices)

        if q_type == "date":
            # Dates in plausible recent past
            years_back = self.rng.randint(1, max(1, min(self.profile.age - 18, 15)))
            yr = 2024 - years_back
            return f"{yr}-{self.rng.randint(1,12):02d}-{self.rng.randint(1,28):02d}"

        if q_type == "time":
            # Realistic waking hours with quarter-hour precision
            hr  = self.rng.randint(6, 22)
            mn  = self.rng.choice([0, 15, 30, 45])
            return f"{hr:02d}:{mn:02d}"

        if q_type == "datetime":
            return (f"2024-{self.rng.randint(1,6):02d}-{self.rng.randint(1,28):02d}"
                    f"T{self.rng.randint(8,18):02d}:00")

        if q_type == "geopoint":
            # Coordinates centred on Sub-Saharan Africa
            return (f"{self.rng.uniform(-5, 15):.4f} "
                    f"{self.rng.uniform(-5, 40):.4f} 0 0")

        if q_type in ("image", "audio", "video", "file"):
            return "media_capture.jpg"
        if q_type == "barcode":     return str(self.rng.randint(100000, 999999))
        if q_type == "range":       return self._gen_range(question)
        if q_type == "acknowledge": return "OK"
        return "response"

    # ── select helpers ────────────────────────────────────────────────────────

    def _gen_select_one(self, name, choices):
        if not choices:
            return "other"
        if self._is_likert(name, choices):
            return self._likert_response(choices)
        return self.rng.choice(choices)

    def _gen_select_multiple(self, choices):
        if not choices:
            return ["other"]
        # Prefer fewer selections (1 most common, rarely 3+)
        max_n = min(len(choices), 4)
        k = self.rng.choices(
            range(1, max_n + 1),
            weights=[40, 30, 20, 10][:max_n]
        )[0]
        return self.rng.sample(choices, k)

    # ── Likert detection & weighted response ──────────────────────────────────

    def _is_likert(self, name, choices):
        # Name contains attitudinal keywords
        if any(kw in name for kw in self._LIKERT_KW):
            return True
        # Choices are a consecutive integer scale (e.g. 1-5, 0-10)
        if 3 <= len(choices) <= 10:
            try:
                nums = [int(c) for c in choices]
                if nums == list(range(min(nums), max(nums) + 1)):
                    return True
            except (ValueError, TypeError):
                pass
        return False

    def _likert_response(self, choices):
        """Weight choices according to respondent's response style archetype."""
        n   = len(choices)
        mid = n // 2
        style = self.profile.response_style

        if style == "neutral":
            # Bell curve around middle
            weights = [max(1, n - abs(i - mid) * 2) for i in range(n)]
        elif style == "agreeable":
            # Skew toward higher end (positive / agreement)
            weights = list(range(1, n + 1))
        elif style == "cautious":
            # Skew toward lower / middle (hesitant, underreporting)
            weights = list(range(n, 0, -1))
            weights[mid] = weights[mid] + 3
        else:  # extreme
            # Prefer endpoints strongly
            weights = [n if (i == 0 or i == n - 1) else 1 for i in range(n)]

        return self.rng.choices(choices, weights=weights)[0]

    # ── integer / decimal ─────────────────────────────────────────────────────

    def _gen_int(self, name):
        if "age"   in name: return self.profile.age
        if any(k in name for k in ("size", "member", "hh", "person")):
            return self.profile.household_size
        if any(k in name for k in ("child", "kid")):
            return self.profile.num_children
        if any(k in name for k in ("hour", "hrs")):
            return self._gen_hours(name)
        if "year"  in name: return self.rng.randint(2010, 2024)
        if "month" in name: return self.rng.randint(1, 12)
        if "day"   in name: return self.rng.randint(1, 30)
        if any(k in name for k in ("income", "earn", "salary", "wage")):
            return {"low":    self.rng.randint(20_000,  80_000),
                    "medium": self.rng.randint(80_000, 250_000),
                    "high":   self.rng.randint(250_000, 1_000_000),
                    }[self.profile.income_level]
        return self.rng.randint(0, 20)

    def _gen_decimal(self, name):
        if any(k in name for k in self._TIME_KW):
            return round(self._gen_hours(name) + self.rng.uniform(0, 0.9), 1)
        if any(k in name for k in ("income", "earn", "price", "cost", "amount")):
            return round(self.rng.uniform(500, 50_000), 0)
        return round(self.rng.uniform(0, 100), 1)

    def _gen_hours(self, name):
        """Generate realistic hours for time-use questions with a daily budget."""
        remaining = max(0, 16 - self._time_used)   # 16 waking hours total
        if   any(k in name for k in ("sleep", "wake")):
            hrs = self.rng.randint(6, 9)
        elif any(k in name for k in ("work", "business", "job", "farm")):
            hrs = self.rng.randint(4, min(10, max(4, remaining)))
        elif any(k in name for k in ("chore", "household", "cook", "clean", "wash")):
            hrs = self.rng.randint(1, min(5, max(1, remaining)))
        elif any(k in name for k in ("care", "child", "baby")):
            hrs = (self.rng.randint(1, 4) if self.profile.num_children > 0 else 0)
        elif any(k in name for k in ("leisure", "relax", "social", "friend")):
            hrs = self.rng.randint(0, min(3, max(0, remaining)))
        else:
            hrs = self.rng.randint(0, max(1, min(4, remaining)))
        self._time_used = min(16, self._time_used + hrs)
        return hrs

    def _gen_range(self, question):
        """Respect the range question's min/max when available."""
        # Attempt to read parameters; fall back to 1-10
        try:
            params     = str(question.get("parameters", "") or "")
            param_dict = dict(p.split("=") for p in params.split() if "=" in p)
            lo = int(param_dict.get("start", 1))
            hi = int(param_dict.get("end",  10))
        except Exception:
            lo, hi = 1, 10
        return self.rng.randint(lo, hi)

    # ── text ──────────────────────────────────────────────────────────────────

    def _gen_text(self, name):
        if "name" in name:
            return f"{self.rng.choice(self._FIRST)} {self.rng.choice(self._LAST)}"
        if any(k in name for k in ("phone", "tel", "mobile")):
            if self.profile.has_phone:
                return f"+2{self.rng.randint(10,99)}{self.rng.randint(10_000_000, 99_999_999)}"
            return ""
        if "other"   in name:
            return self.rng.choice(["Other reason", "Prefer not to say", "Not applicable", "Various"])
        if any(k in name for k in ("comment", "note", "remark")):
            return self.rng.choice(["None", "N/A", "No comments"])
        if "address" in name:
            return f"{self.rng.randint(1, 999)} Main Road"
        return f"Response_{self.rng.randint(100, 999)}"

    # ── profile keyword matching ──────────────────────────────────────────────

    def _from_profile(self, q_name, choices):
        n = q_name.lower()
        for kw, attr in self._KEYWORD_MAP.items():
            if kw in n:
                val = getattr(self.profile, attr, None)
                if val is None:
                    continue
                if isinstance(val, bool):
                    val = "yes" if val else "no"
                if choices:
                    val_s = str(val).lower()
                    for c in choices:
                        if val_s in c.lower() or c.lower() in val_s:
                            return c
                    return self.rng.choice(choices)
                return val
        return None


# ══════════════════════════════════════════════════════════════════════════════
# SIMULATION ENGINE
# ══════════════════════════════════════════════════════════════════════════════

class SimulationResult:
    __slots__ = ("sim_id","profile","path_taken","skipped",
                 "responses","runtime_issues","path_signature")

    def __init__(self, sim_id, profile, path_taken, skipped, responses, runtime_issues):
        self.sim_id         = sim_id
        self.profile        = profile
        self.path_taken     = path_taken
        self.skipped        = skipped
        self.responses      = responses
        self.runtime_issues = runtime_issues
        self.path_signature = tuple(path_taken)


class SimulationEngine:
    MAX_CONSECUTIVE_SKIPS = 10

    def __init__(self, parser: XLSFormParser):
        self.parser    = parser
        self.questions = parser.questions

    def run(self, sim_id: int, seed: int = None) -> SimulationResult:
        rng      = random.Random(seed)
        profile  = RespondentProfile(rng)
        gen      = ResponseGenerator(profile, rng)
        responses: dict      = {}
        path_taken: list     = []
        skipped: list        = []
        runtime_issues: list = []
        consec_skip = 0

        for q in self.questions:
            q_type = q["type"]
            q_name = q["name"]
            base   = q_type.split()[0]

            if base in XLSFormParser.STRUCTURAL:
                continue

            # Evaluate relevance
            try:
                is_relevant = ConditionEvaluator(responses).evaluate(q.get("relevant", ""))
            except Exception:
                is_relevant = True

            if not is_relevant:
                skipped.append(q_name)
                consec_skip += 1
                if consec_skip >= self.MAX_CONSECUTIVE_SKIPS:
                    runtime_issues.append({
                        "question": q_name,
                        "type":     "Deep Skip Chain Encountered",
                        "severity": "Medium",
                        "sim_id":   sim_id,
                    })
                continue

            consec_skip = 0
            path_taken.append(q_name)

            if q_type == "note":
                continue
            if q_type == "calculate":
                calc = q.get("calculation", "")
                if calc:
                    try:
                        responses[q_name] = str(ConditionEvaluator(responses).evaluate(calc))
                    except Exception:
                        responses[q_name] = ""
                continue

            response = gen.generate(q, responses)

            # Validate constraint
            constraint = q.get("constraint", "")
            if constraint and response is not None:
                val_str = (" ".join(str(v) for v in response)
                           if isinstance(response, list) else str(response))
                try:
                    valid = ConditionEvaluator(responses).evaluate(constraint, current_val=val_str)
                    if not valid:
                        response = self._try_satisfy(q, constraint, responses, rng) or response
                except Exception:
                    pass

            # Check required
            if str(q.get("required","")).lower() in ("yes","true","1") and response is None:
                runtime_issues.append({
                    "question": q_name, "type": "Required Question Unanswerable",
                    "severity": "High", "sim_id": sim_id,
                })

            if response is not None:
                responses[q_name] = (
                    " ".join(str(v) for v in response)
                    if isinstance(response, list) else str(response)
                )

        return SimulationResult(sim_id, profile.as_dict(), path_taken,
                                skipped, responses, runtime_issues)

    def _try_satisfy(self, q, constraint, responses, rng):
        choices = list(q.get("choices", []))
        if not choices: return None
        rng.shuffle(choices)
        for c in choices:
            try:
                if ConditionEvaluator(responses).evaluate(constraint, current_val=c):
                    return c
            except Exception:
                pass
        return None


# ══════════════════════════════════════════════════════════════════════════════
# ISSUE
# ══════════════════════════════════════════════════════════════════════════════

class Issue:
    __slots__ = ("question","issue_type","severity","explanation","suggestion","relevant","source")
    SEV_ORDER = {"Critical": 0, "High": 1, "Medium": 2, "Low": 3}

    def __init__(self, question, issue_type, severity, explanation, suggestion,
                 relevant="", source="static"):
        self.question    = question
        self.issue_type  = issue_type
        self.severity    = severity
        self.explanation = explanation
        self.suggestion  = suggestion
        self.relevant    = relevant
        self.source      = source   # "static" | "simulation"

    def to_dict(self):
        return {k: getattr(self, k) for k in self.__slots__}


# ══════════════════════════════════════════════════════════════════════════════
# STATIC ANALYZER  (World Bank ietestform + IPA ipacheckscto checks)
# ══════════════════════════════════════════════════════════════════════════════

class StaticAnalyzer:
    """
    Purely structural checks — no simulation required.
    Consolidates checks from:
      - World Bank DIME ietestform.ado
      - IPA ipacheckscto.ado
    """

    DATA_TYPES = {
        "select_one","select_multiple","date","geopoint","text",
        "integer","decimal","barcode","time","image","audio","video",
        "file","datetime","range","acknowledge",
    }

    def __init__(self, parser: XLSFormParser):
        self.parser   = parser
        self._issues: list = []
        self._seen:   set  = set()

    def analyze_all(self) -> list:
        self._check_metadata_fields()           # IPA §1
        self._check_missing_recommended_cols()  # WB §9
        self._check_leading_trailing_spaces()   # WB §2 / §10
        self._check_group_structure()           # WB §12-14 / IPA §8
        self._check_field_name_length()         # WB §15-17 / IPA §2
        self._check_invalid_name_chars()        # IPA §2
        self._check_note_required()             # WB §22-23 / IPA §4
        self._check_disabled_readonly()         # IPA §3
        self._check_non_required_fields()       # WB §21 / IPA §4
        self._check_duplicate_choice_codes()    # WB §4 / IPA §10
        self._check_missing_choice_labels()     # WB §6 / IPA §10
        self._check_duplicate_choice_labels()   # WB §7 / IPA §10
        self._check_unused_choice_lists()       # WB §24
        self._check_outdated_syntax()           # WB §11
        self._check_or_other_syntax()           # IPA §6
        self._check_integer_no_constraint()     # IPA §5
        self._check_constraint_no_message()     # IPA §5
        self._check_constraint_contradictions() # custom (fixed)
        return sorted(self._issues, key=lambda i: Issue.SEV_ORDER.get(i.severity, 3))

    # ── helpers ───────────────────────────────────────────────────────────────

    def _add(self, issue: Issue):
        key = (issue.question, issue.issue_type)
        if key not in self._seen:
            self._seen.add(key)
            self._issues.append(issue)

    # ── IPA §1 — Metadata fields ──────────────────────────────────────────────

    def _check_metadata_fields(self):
        types_present = {q["type"] for q in self.parser.questions}
        for meta_type, sev, desc in [
            ("start",    "Medium", "Tracks interview start time."),
            ("end",      "Medium", "Tracks interview end time."),
            ("deviceid", "Low",    "Identifies the data collection device."),
        ]:
            if meta_type not in types_present:
                self._add(Issue(
                    "(form)", f"Missing Metadata Field: {meta_type}", sev,
                    f'No field with type="{meta_type}" was found.',
                    f'Add a field with type="{meta_type}" to capture this metadata automatically.',
                ))

    # ── WB §9 — Missing recommended columns ───────────────────────────────────

    def _check_missing_recommended_cols(self):
        if self.parser.survey_df is None:
            return
        cols = self.parser.survey_df.columns.tolist()
        for col, sev, tip in [
            ("hint",               "Low",  "Hint text guides enumerators on how to ask each question."),
            ("constraint_message", "Low",  "Constraint messages explain to enumerators why validation failed."),
            ("required_message",   "Low",  "Required messages explain why a field cannot be left blank."),
        ]:
            present = any(c == col or c.startswith(col + ":") for c in cols)
            if not present:
                self._add(Issue(
                    "(form)", f"Missing Recommended Column: {col}", sev,
                    f'The "{col}" column is absent from the survey sheet.',
                    tip,
                ))

    # ── WB §2/§10 — Leading/trailing spaces ───────────────────────────────────

    def _check_leading_trailing_spaces(self):
        for q in self.parser.questions:
            for field in ("name", "type", "required"):
                raw = q.get(field, "")
                if raw and raw != raw.strip():
                    self._add(Issue(
                        q["name"], "Leading/Trailing Spaces in Survey Sheet", "High",
                        f'The "{field}" column for "{q["name"]}" has leading or trailing whitespace.',
                        f'Trim the "{field}" cell. Spaces cause silent parsing errors in SurveyCTO.',
                    ))

        if self.parser.choices_df is not None:
            for _, row in self.parser.choices_df.iterrows():
                for field in ("list_name", "name"):
                    raw = str(row.get(field, "") or "")
                    if raw and raw.lower() not in ("nan", "") and raw != raw.strip():
                        self._add(Issue(
                            f"choices/{raw.strip()}", "Leading/Trailing Spaces in Choices Sheet", "High",
                            f'Choice "{field}" value "{raw}" has surrounding whitespace.',
                            "Trim all cells in the choices sheet, especially list_name and name.",
                        ))

    # ── WB §12-14 / IPA §8 — Group structure ─────────────────────────────────

    def _check_group_structure(self):
        def norm(t):
            return t.lower().replace(" ", "_")   # → begin_group, end_repeat, etc.

        stack = []
        for q in self.parser.questions:
            nt = norm(q["type"])
            if nt in ("begin_group", "begin_repeat"):
                stack.append((nt, q["name"]))
            elif nt in ("end_group", "end_repeat"):
                if not q["name"]:
                    self._add(Issue(
                        "(unnamed)", "Missing End-Group Name", "High",
                        f'An {q["type"]} row has no name.',
                        "Add a name that matches the corresponding begin statement.",
                    ))
                    continue
                if stack:
                    begin_nt, begin_name = stack.pop()
                    begin_base = begin_nt.replace("begin_", "")
                    end_base   = nt.replace("end_", "")
                    if begin_base != end_base:
                        self._add(Issue(
                            q["name"], "Group Type Mismatch", "Critical",
                            f'"{begin_name}" opens as {begin_nt} but is closed by {nt}.',
                            f'Change to end_{begin_base} to match.',
                        ))
                    elif begin_name and q["name"] and begin_name != q["name"]:
                        self._add(Issue(
                            q["name"], "Group Name Mismatch", "Critical",
                            f'Group begins as "{begin_name}" but closes as "{q["name"]}".',
                            'Ensure begin/end names match exactly.',
                        ))
                else:
                    self._add(Issue(
                        q["name"], "Unmatched End Statement", "Critical",
                        f'"{q["name"]}" ({q["type"]}) has no matching begin statement.',
                        "Add a matching begin_group or begin_repeat before this row.",
                    ))

        for nt, name in stack:
            self._add(Issue(
                name, "Unclosed Group/Repeat", "Critical",
                f'"{name}" ({nt}) is never closed with a matching end statement.',
                "Add a matching end_group or end_repeat.",
            ))

    # ── WB §15-17 / IPA §2 — Field name length ───────────────────────────────

    def _check_field_name_length(self):
        for q in self.parser.questions:
            name = q["name"]
            if len(name) > 32:
                self._add(Issue(
                    name, "Field Name Too Long (Stata limit: 32)", "High",
                    f'"{name}" is {len(name)} characters. Stata enforces a 32-character variable name limit.',
                    "Shorten to ≤32 characters to ensure compatibility with Stata exports.",
                ))
            elif len(name) > 22:
                self._add(Issue(
                    name, "Field Name Too Long (SurveyCTO best practice: 22)", "Medium",
                    f'"{name}" is {len(name)} characters. SurveyCTO recommends ≤22 characters.',
                    "Shorten to ≤22 characters for better SurveyCTO compatibility.",
                ))

    # ── IPA §2 — Invalid characters in field names ────────────────────────────

    def _check_invalid_name_chars(self):
        for q in self.parser.questions:
            name = q["name"]
            bad  = [c for c in (".", "-") if c in name]
            if bad:
                chars = " and ".join(f'"{c}"' for c in bad)
                self._add(Issue(
                    name, "Invalid Characters in Field Name", "High",
                    f'"{name}" contains {chars}, which can break SurveyCTO and statistical software.',
                    "Use only letters, numbers, and underscores (_) in field names.",
                ))

    # ── WB §22-23 / IPA §4 — Note/label marked required ─────────────────────

    def _check_note_required(self):
        for q in self.parser.questions:
            req = q.get("required", "").lower()
            if req not in ("yes", "true", "1"):
                continue
            if q["type"] == "note":
                self._add(Issue(
                    q["name"], "Required Note Field", "High",
                    f'"{q["name"]}" is type=note but marked required. Notes cannot be required.',
                    'Remove "yes" from the required column for this note field.',
                ))
            if q.get("appearance", "").lower() == "label":
                self._add(Issue(
                    q["name"], "Required Label-Appearance Field", "High",
                    f'"{q["name"]}" has appearance=label but is marked required. Label fields cannot be required.',
                    'Remove "yes" from required, or change the appearance.',
                ))
            if q.get("read_only", "").lower() in ("yes", "true", "1"):
                self._add(Issue(
                    q["name"], "Required Read-Only Field", "Medium",
                    f'"{q["name"]}" is both required and read_only. Enumerators cannot enter a value.',
                    'Either remove required=yes, or remove read_only=yes.',
                ))

    # ── IPA §3 — Disabled / read-only ────────────────────────────────────────

    def _check_disabled_readonly(self):
        for q in self.parser.questions:
            if q.get("disabled", "").lower() in ("yes", "true", "1"):
                self._add(Issue(
                    q["name"], "Disabled Field", "Medium",
                    f'"{q["name"]}" is marked as disabled and will be hidden from enumerators.',
                    "Confirm this is intentional. Remove disabled=yes to make the field active.",
                ))

    # ── WB §21 / IPA §4 — Non-required data fields ────────────────────────────

    def _check_non_required_fields(self):
        for q in self.parser.questions:
            base = q["type"].split()[0]
            if base not in self.DATA_TYPES:
                continue
            if q.get("appearance", "").lower() == "label":
                continue
            req = q.get("required", "").lower()
            if req not in ("yes", "true", "1"):
                self._add(Issue(
                    q["name"], "Non-Required Data Field", "Low",
                    f'"{q["name"]}" ({q["type"]}) is not marked as required.',
                    'Add required=yes unless this field is intentionally optional.',
                ))

    # ── WB §4 / IPA §10 — Duplicate choice codes ─────────────────────────────

    def _check_duplicate_choice_codes(self):
        if self.parser.choices_df is None:
            return
        seen: dict = {}
        for _, row in self.parser.choices_df.iterrows():
            ln   = str(row.get("list_name", "") or "").strip()
            name = str(row.get("name", "") or "").strip()
            if not ln or not name or name.lower() == "nan":
                continue
            key = (ln, name)
            if key in seen:
                self._add(Issue(
                    f"{ln}", "Duplicate Choice Code", "Critical",
                    f'Code "{name}" appears more than once in choice list "{ln}".',
                    "Each choice code must be unique within a list. Remove or rename the duplicate.",
                ))
            seen[key] = True

    # ── WB §6 / IPA §10 — Missing choice labels ──────────────────────────────

    def _check_missing_choice_labels(self):
        if self.parser.choices_df is None:
            return
        label_cols = [c for c in self.parser.choices_df.columns
                      if c == "label" or c.startswith("label:")]
        if not label_cols:
            return
        for _, row in self.parser.choices_df.iterrows():
            ln   = str(row.get("list_name", "") or "").strip()
            name = str(row.get("name",      "") or "").strip()
            if not ln or not name or name.lower() == "nan":
                continue
            labels = [str(row.get(c, "") or "").strip() for c in label_cols]
            if not any(l and l.lower() != "nan" for l in labels):
                self._add(Issue(
                    f"{ln}/{name}", "Missing Choice Label", "High",
                    f'Choice "{name}" in list "{ln}" has no label in any language.',
                    "Add a label in the choices sheet for this choice.",
                ))

    # ── WB §7 / IPA §10 — Duplicate choice labels ────────────────────────────

    def _check_duplicate_choice_labels(self):
        if self.parser.choices_df is None:
            return
        label_cols = [c for c in self.parser.choices_df.columns
                      if c == "label" or c.startswith("label:")]
        for col in label_cols:
            seen: dict = {}
            for _, row in self.parser.choices_df.iterrows():
                ln    = str(row.get("list_name", "") or "").strip()
                label = str(row.get(col,         "") or "").strip()
                if not ln or not label or label.lower() == "nan":
                    continue
                key = (ln, label.lower())
                if key in seen:
                    self._add(Issue(
                        f"{ln}", "Duplicate Choice Label", "Medium",
                        f'Label "{label}" appears more than once in list "{ln}" ({col}).',
                        "Give each choice a distinct label to avoid respondent confusion.",
                    ))
                seen[key] = True

    # ── WB §24 — Unused choice lists ─────────────────────────────────────────

    def _check_unused_choice_lists(self):
        used = {q["list_name"] for q in self.parser.questions if q.get("list_name")}
        for ln in self.parser.choices_dict:
            if ln not in used:
                self._add(Issue(
                    f"choices/{ln}", "Unused Choice List", "Low",
                    f'Choice list "{ln}" is defined but not referenced by any survey question.',
                    "Remove unused choice lists to keep the form clean and reduce confusion.",
                ))

    # ── WB §11 — Outdated XLSForm syntax ─────────────────────────────────────

    def _check_outdated_syntax(self):
        PATTERNS = [
            (r"position\(\)",       "position()"),
            (r"jr:choice-name\(",   "jr:choice-name()"),
            (r"\bjr:itext\(",       "jr:itext()"),
        ]
        fields = ("relevant","constraint","calculation","default")
        for q in self.parser.questions:
            for field in fields:
                val = q.get(field, "")
                if not val:
                    continue
                for pattern, label in PATTERNS:
                    if re.search(pattern, val, re.IGNORECASE):
                        self._add(Issue(
                            q["name"], f"Outdated Syntax: {label}", "Low",
                            f'"{q["name"]}" uses deprecated "{label}" in the {field} column.',
                            "Replace with the current SurveyCTO/XLSForm equivalent syntax.",
                        ))

    # ── IPA §6 — or_other deprecated syntax ──────────────────────────────────

    def _check_or_other_syntax(self):
        for q in self.parser.questions:
            if "or_other" in q["type"].lower():
                self._add(Issue(
                    q["name"], "Deprecated or_other Syntax", "Medium",
                    f'"{q["name"]}" uses or_other, which is not supported in SurveyCTO.',
                    'Add an explicit "other" option to the choice list and a separate text '
                    'field gated by relevant=${field_name}="other".',
                ))

    # ── IPA §5 — Integer/decimal without constraint ───────────────────────────

    def _check_integer_no_constraint(self):
        for q in self.parser.questions:
            if q["type"] in ("integer", "decimal") and not q.get("constraint"):
                self._add(Issue(
                    q["name"], "Numeric Field Without Range Constraint", "Low",
                    f'"{q["name"]}" ({q["type"]}) has no constraint to validate entered values.',
                    'Add a constraint such as ". >= 0 and . <= 120" to catch impossible values early.',
                ))

    # ── IPA §5 — Constraint without message ──────────────────────────────────

    def _check_constraint_no_message(self):
        if self.parser.survey_df is None:
            return
        cols         = self.parser.survey_df.columns.tolist()
        has_cm_col   = any(c == "constraint_message" or c.startswith("constraint_message:") for c in cols)
        if not has_cm_col:
            return  # column missing entirely — caught by _check_missing_recommended_cols
        for q in self.parser.questions:
            if q.get("constraint") and not q.get("constraint_message"):
                self._add(Issue(
                    q["name"], "Constraint Without Message", "Low",
                    f'"{q["name"]}" has a constraint but no constraint_message.',
                    "Add a constraint_message so enumerators understand why their input was rejected.",
                ))

    # ── Constraint contradiction check (fixed not() evaluator) ────────────────

    def _check_constraint_contradictions(self):
        for q in self.parser.questions:
            constraint = q.get("constraint", "")
            choices    = q.get("choices", [])
            base       = q["type"].split()[0]
            if not constraint or not choices:
                continue
            if base not in ("select_one", "select_multiple"):
                continue

            # Test each choice individually — if any single choice satisfies the
            # constraint the question is answerable (fixed not() parser handles
            # patterns like not(selected(., '-99') and count-selected(.) > 1))
            satisfiable = False
            for c in choices:
                try:
                    if ConditionEvaluator({}).evaluate(constraint, current_val=c):
                        satisfiable = True
                        break
                except Exception:
                    satisfiable = True   # can't determine → assume ok
                    break

            if not satisfiable:
                sample = ", ".join(choices[:5]) + ("…" if len(choices) > 5 else "")
                self._add(Issue(
                    q["name"], "Constraint Contradiction", "High",
                    f'No available choice appears to satisfy the constraint `{constraint}` for "{q["name"]}".',
                    f"Fix the constraint expression. Available choices: {sample}",
                    q.get("relevant",""),
                ))


# ══════════════════════════════════════════════════════════════════════════════
# ISSUE DETECTOR  (simulation-informed checks)
# ══════════════════════════════════════════════════════════════════════════════

class IssueDetector:
    """Checks that require simulation results."""

    def __init__(self, parser: XLSFormParser, results: list):
        self.parser    = parser
        self.results   = results
        self.questions = parser.questions
        self._issues:  list = []
        self._seen:    set  = set()

    def detect_all(self) -> list:
        self._check_unreachable_questions()
        self._check_rarely_reached()
        self._check_circular_dependencies()
        self._check_deep_dependency_chains()
        self._collect_runtime_issues()
        return sorted(self._issues, key=lambda i: Issue.SEV_ORDER.get(i.severity, 3))

    def _add(self, issue: Issue):
        key = (issue.question, issue.issue_type)
        if key not in self._seen:
            self._seen.add(key)
            self._issues.append(issue)

    def _check_unreachable_questions(self):
        all_visited = set()
        for r in self.results:
            all_visited.update(r.path_taken)
        for q in self.questions:
            base = q["type"].split()[0]
            if base in XLSFormParser.STRUCTURAL or q["type"] == "note":
                continue
            if q["name"] not in all_visited and q.get("relevant"):
                self._add(Issue(
                    q["name"], "Unreachable Question", "High",
                    (f'"{q["name"]}" was never visited in {len(self.results)} simulation(s). '
                     f'Its relevance condition may always evaluate to false.'),
                    f'Review: `{q["relevant"]}`. Ensure referenced variables exist and '
                    f'are answered before this question.',
                    q.get("relevant",""), source="simulation",
                ))

    def _check_rarely_reached(self):
        visit  = Counter()
        skip   = Counter()
        total  = len(self.results)
        for r in self.results:
            for n in r.path_taken: visit[n] += 1
            for n in r.skipped:    skip[n]  += 1

        for q in self.questions:
            n    = q["name"]
            v, s = visit[n], skip[n]
            if (v + s) > 0 and s / (v + s) > 0.93 and q.get("relevant") and total >= 10:
                self._add(Issue(
                    n, "Rarely Reached Question", "Low",
                    f'"{n}" was skipped {s/(v+s)*100:.0f}% of the time across {total} simulations.',
                    f'Confirm the relevance condition `{q["relevant"]}` is intentional.',
                    q.get("relevant",""), source="simulation",
                ))

    def _check_circular_dependencies(self):
        deps = {}
        for q in self.questions:
            if q.get("relevant"):
                deps[q["name"]] = set(re.findall(r"\$\{([^}]+)\}", q["relevant"]))

        def has_cycle(start, current, stack):
            if current in stack: return current == start
            return any(has_cycle(start, nxt, stack | {current})
                       for nxt in deps.get(current, set()))

        for name in deps:
            if has_cycle(name, name, set()):
                self._add(Issue(
                    name, "Circular Skip Logic", "Critical",
                    (f'"{name}" participates in a circular relevance dependency. '
                     f'The form may behave unpredictably or loop infinitely.'),
                    "Break the cycle by restructuring skip logic. Use calculate fields for intermediate flags.",
                    source="simulation",
                ))

    def _check_deep_dependency_chains(self):
        deps = {}
        for q in self.questions:
            if q.get("relevant"):
                deps[q["name"]] = list(re.findall(r"\$\{([^}]+)\}", q["relevant"]))
        memo = {}

        def depth(name, path=frozenset()):
            if name in memo:     return memo[name]
            if name in path:     return 0
            children = deps.get(name, [])
            if not children:     return 0
            d = 1 + max((depth(c, path | {name}) for c in children), default=0)
            memo[name] = d
            return d

        for q in self.questions:
            d = depth(q["name"])
            if d > 5:
                self._add(Issue(
                    q["name"], "Excessively Deep Skip Chain", "Medium",
                    f'"{q["name"]}" sits at skip-logic dependency depth {d} (recommended max: 5).',
                    "Flatten the relevance chain or pre-compute values with calculate fields.",
                    q.get("relevant",""), source="simulation",
                ))

    def _collect_runtime_issues(self):
        for r in self.results:
            for ri in r.runtime_issues:
                self._add(Issue(
                    ri["question"], ri["type"], ri["severity"],
                    f'Detected during simulation #{ri["sim_id"]}.',
                    "Review the question requirements and surrounding skip logic.",
                    source="simulation",
                ))


# ══════════════════════════════════════════════════════════════════════════════
# DESIGN ADVISOR  (applied economics / social science best practices)
# ══════════════════════════════════════════════════════════════════════════════

class Suggestion:
    """A design improvement recommendation (not an error, but a best-practice hint)."""
    __slots__ = ("category", "title", "description", "action", "example", "vars", "priority")
    PRI_ORDER  = {"High": 0, "Medium": 1, "Low": 2}

    def __init__(self, category, title, description, action, vars=None,
                 priority="Medium", example=""):
        self.category    = category
        self.title       = title
        self.description = description
        self.action      = action
        self.example     = example
        self.vars        = vars or []
        self.priority    = priority


class DesignAdvisor:
    """
    Analyses parsed XLSForm questions and returns design improvement suggestions
    grounded in applied economics / social science survey methodology.
    Distinct from the StaticAnalyzer — these are not errors but opportunities
    to improve data quality, reduce non-response, and strengthen skip logic.
    """

    # Patterns that suggest a choice is meant to be exclusive
    _EXCLUSIVE_NAMES = re.compile(
        r"^(dk|don[t']?_?know|refuse[d]?|na|n_a|not_applicable|none|no_response"
        r"|prefer_not|no_answer|idk|unknown|refused|noanswer|no_opinion|not_sure)$",
        re.I,
    )
    _EXCLUSIVE_LABELS = re.compile(
        r"\b(don[''']?t know|do not know|refuse[d]?|not applicable|none of the above"
        r"|prefer not to answer|no response|no opinion|not sure|unsure)\b",
        re.I,
    )

    # Field-name keywords that suggest sensitive / numeric domains
    _AGE_KWORDS       = re.compile(r"\b(age|yrs?|years?_old|age_hh)\b", re.I)
    _YEAR_KWORDS      = re.compile(r"\b(year|yr|dob|birth_year|year_born)\b", re.I)
    _INCOME_KWORDS    = re.compile(r"\b(income|wage|salary|earn|revenue|profit|remit|transfer)\b", re.I)
    _EXPEND_KWORDS    = re.compile(r"\b(expend|spend|cost|payment|price|amount|value|asset)\b", re.I)
    _COUNT_KWORDS     = re.compile(r"\b(num_|n_|count_|total_|number_of_|hh_size|hhsize|members)\b", re.I)
    _SENSITIVE_KWORDS = re.compile(r"\b(income|wage|asset|wealth|religion|ethnic|politi|hiv|sex|violence)\b", re.I)
    _FOLLOWUP_SUFFIX  = re.compile(r"_(reason|specify|detail|explain|other_text|other_spec|comment|other)$", re.I)
    _LIKERT_LABELS    = re.compile(
        r"\b(strongly agree|agree|neutral|disagree|strongly disagree"
        r"|always|often|sometimes|rarely|never"
        r"|very (good|bad|satisfied|dissatisfied|likely|unlikely|important|concerned)"
        r"|extremely|very much|somewhat|not at all)\b",
        re.I,
    )

    def __init__(self, parser: XLSFormParser):
        self.parser      = parser
        self.questions   = parser.answerable_questions()
        self.choices     = parser.choices_dict
        self._sugg: list = []
        self._seen: set  = set()

    # ── public ────────────────────────────────────────────────────────────────

    def analyze_all(self) -> list:
        # ── Integrity (broken references / data corruption) ───────────────────
        self._check_broken_variable_references()
        self._check_select_multiple_space_names()
        # ── Choice logic ──────────────────────────────────────────────────────
        self._check_exclusive_choice_constraints()
        self._check_missing_dk_on_sensitive()
        # ── Validation bounds ─────────────────────────────────────────────────
        self._check_numeric_bounds()
        self._check_required_without_message()
        # ── Skip logic ────────────────────────────────────────────────────────
        self._check_implicit_followups()
        self._check_readonly_without_default()
        # ── Calculate fields ──────────────────────────────────────────────────
        self._check_calculate_once()
        self._check_pulldata_type_conversion()
        # ── Guidance & labels ─────────────────────────────────────────────────
        self._check_hint_coverage()
        self._check_label_quality()
        self._check_text_multiline_appearance()
        # ── Question design ───────────────────────────────────────────────────
        self._check_likert_balance()
        self._check_likert_appearance()
        # ── GPS / media ───────────────────────────────────────────────────────
        self._check_geopoint_accuracy()
        # ── Form metadata ─────────────────────────────────────────────────────
        self._check_settings_completeness()
        self._check_audit_trail()
        self._check_multilingual_completeness()
        # ── Survey flow ───────────────────────────────────────────────────────
        self._check_section_notes()
        # ── Performance & efficiency ──────────────────────────────────────────
        self._check_large_choice_lists()
        self._check_very_large_choice_csv()
        self._check_repeat_candidates()
        self._check_form_complexity()
        return sorted(self._sugg, key=lambda s: Suggestion.PRI_ORDER.get(s.priority, 1))

    # ── internal helpers ──────────────────────────────────────────────────────

    def _add(self, s: Suggestion):
        key = (s.category, s.title, tuple(sorted(s.vars)))
        if key not in self._seen:
            self._seen.add(key)
            self._sugg.append(s)

    # ── 1. Exclusive-choice constraints on select_multiple ────────────────────

    def _check_exclusive_choice_constraints(self):
        for q in self.questions:
            if not q["type"].startswith("select_multiple"):
                continue
            choices    = q["choices"]
            constraint = q.get("constraint", "")
            excl = [
                c for c in choices
                if self._EXCLUSIVE_NAMES.match(c)
                or any(self._EXCLUSIVE_LABELS.search(lbl)
                       for lbl in self._choice_labels(q["list_name"]))
            ]
            # Deduplicate: keep only choices whose name matches
            excl = [c for c in choices if self._EXCLUSIVE_NAMES.match(c)]
            if not excl:
                continue
            # Check whether any existing constraint already guards these options
            already_guarded = any(e in constraint for e in excl)
            if already_guarded:
                continue
            names_str = " / ".join(f"`{e}`" for e in excl)
            self._add(Suggestion(
                category="Choice Logic",
                title="Exclusive option selected alongside other responses",
                description=(
                    f"**{q['name']}** is a `select_multiple` question containing "
                    f"exclusive option(s) {names_str} (e.g. 'Don't know', 'None', 'Refuse'). "
                    "Without a constraint, respondents can select these alongside substantive "
                    "answers, producing contradictory data."
                ),
                action=(
                    "Add a `constraint` that prevents co-selection. "
                    f"Example for option `{excl[0]}`:"
                ),
                example=(
                    f"not(selected(., '{excl[0]}')) or count-selected(.) = 1"
                ),
                vars=[q["name"]],
                priority="High",
            ))

    # ── 2. Sensitive questions without refusal options ────────────────────────

    def _check_missing_dk_on_sensitive(self):
        for q in self.questions:
            base = q["type"].split()[0]
            if base not in ("select_one", "select_multiple"):
                continue
            if not self._SENSITIVE_KWORDS.search(q["name"]) and \
               not self._SENSITIVE_KWORDS.search(q.get("label", "")):
                continue
            choices = q["choices"]
            has_dk = any(self._EXCLUSIVE_NAMES.match(c) for c in choices)
            if has_dk:
                continue
            self._add(Suggestion(
                category="Respondent Experience",
                title="Sensitive question missing 'Prefer not to answer' option",
                description=(
                    f"**{q['name']}** appears to ask about a sensitive topic "
                    "(income, assets, religion, ethnicity, health, or violence) "
                    "but has no 'Don't know' or 'Prefer not to answer' choice. "
                    "Omitting this option forces a response or causes item non-response."
                ),
                action=(
                    "Add a choice (e.g. `prefer_not_to_answer`) to the choice list "
                    f"`{q['list_name']}`, labelled 'Prefer not to answer' or 'Refuse'."
                ),
                vars=[q["name"]],
                priority="Medium",
            ))

    # ── 3. Numeric fields without plausible bounds ────────────────────────────

    def _check_numeric_bounds(self):
        for q in self.questions:
            if q["type"].split()[0] not in ("integer", "decimal"):
                continue
            name       = q["name"]
            constraint = q.get("constraint", "")

            if self._AGE_KWORDS.search(name):
                if not constraint:
                    self._add(Suggestion(
                        category="Validation",
                        title="Age field without plausible bounds",
                        description=(
                            f"**{name}** appears to capture age but has no constraint. "
                            "Without bounds, ages of 0, 999, or negative values will pass "
                            "validation and require expensive cleaning later."
                        ),
                        action="Add a constraint to restrict implausible values:",
                        example=". >= 0 and . <= 120",
                        vars=[name],
                        priority="High",
                    ))

            elif self._YEAR_KWORDS.search(name):
                if not constraint:
                    self._add(Suggestion(
                        category="Validation",
                        title="Year field without range constraint",
                        description=(
                            f"**{name}** appears to capture a year but has no constraint. "
                            "Four-digit entry errors (e.g. 19, 20199) will be undetectable."
                        ),
                        action="Add a constraint with a plausible year range:",
                        example=". >= 1900 and . <= 2025",
                        vars=[name],
                        priority="Medium",
                    ))

            elif self._COUNT_KWORDS.search(name):
                if not constraint or ">=" not in constraint:
                    self._add(Suggestion(
                        category="Validation",
                        title="Count field without non-negative constraint",
                        description=(
                            f"**{name}** appears to count people or items but does not "
                            "prevent negative values."
                        ),
                        action="Add a non-negative constraint:",
                        example=". >= 0",
                        vars=[name],
                        priority="Medium",
                    ))

            elif self._INCOME_KWORDS.search(name) or self._EXPEND_KWORDS.search(name):
                if not constraint:
                    self._add(Suggestion(
                        category="Validation",
                        title="Income / expenditure field without bounds",
                        description=(
                            f"**{name}** captures a monetary amount but has no constraint. "
                            "Typos in large numbers (e.g. an extra zero) are a common data "
                            "entry error and are hard to detect post-collection."
                        ),
                        action=(
                            "Add a non-negative lower bound and consider a soft upper bound "
                            "appropriate for your population. Use a `constraint_message` to "
                            "prompt the enumerator to confirm unusually large values."
                        ),
                        example=". >= 0",
                        vars=[name],
                        priority="Medium",
                    ))

    # ── 4. Implicit follow-up questions without relevance conditions ──────────

    def _check_implicit_followups(self):
        names = {q["name"] for q in self.questions}
        for q in self.questions:
            if q.get("relevant"):
                continue
            if not self._FOLLOWUP_SUFFIX.search(q["name"]):
                continue
            # Try to find the likely parent question
            stem = self._FOLLOWUP_SUFFIX.sub("", q["name"])
            parent_exists = stem in names
            self._add(Suggestion(
                category="Skip Logic",
                title="Follow-up question without relevance condition",
                description=(
                    f"**{q['name']}** looks like a follow-up or 'specify other' question "
                    f"{'(likely parent: `' + stem + '`)' if parent_exists else ''} "
                    "but has no `relevant` condition. It will be displayed to every respondent "
                    "regardless of their prior answer."
                ),
                action=(
                    "Add a `relevant` condition so this question only appears when needed. "
                    + (f"Example (assuming parent `{stem}` has an 'other' option):" if parent_exists else "Example:")
                ),
                example=(f"selected(${{stem}}, 'other')" if parent_exists else "selected(${parent}, 'other')").replace("stem", stem),
                vars=[q["name"]],
                priority="High",
            ))

    # ── 5. Hint text coverage on complex questions ────────────────────────────

    def _check_hint_coverage(self):
        complex_types = {"integer", "decimal", "text", "date", "datetime"}
        for q in self.questions:
            base = q["type"].split()[0]
            if base not in complex_types:
                continue
            if q.get("hint"):
                continue
            has_constraint = bool(q.get("constraint"))
            is_sensitive   = bool(self._SENSITIVE_KWORDS.search(q["name"])
                                   or self._SENSITIVE_KWORDS.search(q.get("label", "")))
            label_long     = len(q.get("label", "")) > 120
            if not (has_constraint or is_sensitive or label_long):
                continue
            reason = (
                "has a validation constraint" if has_constraint
                else "appears sensitive" if is_sensitive
                else "has a long label"
            )
            self._add(Suggestion(
                category="Enumerator Guidance",
                title="Complex question without hint text",
                description=(
                    f"**{q['name']}** {reason} but provides no `hint` to the enumerator. "
                    "Hints appear below the question on tablets and are an effective way to "
                    "communicate valid ranges, units, or instructions without cluttering the label."
                ),
                action=(
                    "Add a `hint` column entry. For a constrained numeric field, specify the "
                    "expected unit and range (e.g. 'Enter amount in local currency, 0–99999'). "
                    "For sensitive questions, include a brief privacy assurance."
                ),
                vars=[q["name"]],
                priority="Low",
            ))

    # ── 6. Label quality ──────────────────────────────────────────────────────

    def _check_label_quality(self):
        for q in self.questions:
            label = q.get("label", "")
            name  = q["name"]
            if not label:
                continue
            # Very long label
            if len(label) > 180:
                self._add(Suggestion(
                    category="Question Design",
                    title="Excessively long question label",
                    description=(
                        f"**{name}** has a label of {len(label)} characters. "
                        "On a tablet screen this may wrap across several lines and slow the "
                        "interview, increasing enumerator fatigue and respondent drop-off."
                    ),
                    action=(
                        "Shorten the label to the core question (≤150 chars). Move "
                        "definitions, examples, and ranges into the `hint` field."
                    ),
                    vars=[name],
                    priority="Low",
                ))
            # Possible double-barreled question
            if " and " in label.lower() and label.strip().endswith("?"):
                self._add(Suggestion(
                    category="Question Design",
                    title="Possible double-barreled question",
                    description=(
                        f"**{name}** contains 'and' within a question ending in '?'. "
                        "Double-barreled questions ask about two things simultaneously, "
                        "making responses ambiguous and difficult to interpret."
                    ),
                    action=(
                        "Split into two separate questions, each asking about a single concept. "
                        "If the second concept is conditional, add a relevance condition."
                    ),
                    vars=[name],
                    priority="Medium",
                ))

    # ── 7. Likert scale balance ───────────────────────────────────────────────

    def _check_likert_balance(self):
        checked_lists = set()
        for q in self.questions:
            ln = q.get("list_name")
            if not ln or ln in checked_lists:
                continue
            choices = self.choices.get(ln, [])
            if len(choices) < 3:
                continue
            # Get labels for these choices
            labels_text = " | ".join(self._choice_labels(ln))
            if not self._LIKERT_LABELS.search(labels_text):
                continue
            checked_lists.add(ln)
            # Count positive vs negative poles
            positive = len(re.findall(
                r"\b(agree|always|good|satisfied|likely|important|positive|yes|often)\b",
                labels_text, re.I))
            negative = len(re.findall(
                r"\b(disagree|never|bad|dissatisfied|unlikely|unimportant|negative|no|rarely)\b",
                labels_text, re.I))
            if positive > 0 and negative == 0:
                self._add(Suggestion(
                    category="Question Design",
                    title="Likert scale may lack negative pole",
                    description=(
                        f"Choice list **{ln}** appears to be a Likert-type scale but may be "
                        "missing a negative pole (e.g. 'Disagree', 'Never', 'Dissatisfied'). "
                        "Unbalanced scales introduce acquiescence bias — respondents tend to "
                        "select positive options when negatives are unavailable."
                    ),
                    action=(
                        "Ensure the scale has symmetric positive and negative options around "
                        "a neutral midpoint (e.g. Strongly Agree / Agree / Neutral / "
                        "Disagree / Strongly Disagree)."
                    ),
                    vars=[q2["name"] for q2 in self.questions if q2.get("list_name") == ln],
                    priority="Medium",
                ))

    # ── 8. Audit / paradata fields ────────────────────────────────────────────

    def _check_audit_trail(self):
        all_names = [q["name"] for q in self.questions]
        all_types = [q["type"] for q in self.parser.questions]
        has_audit = any("audit" in t for t in all_types)
        if not has_audit:
            self._add(Suggestion(
                category="Data Quality",
                title="No audit field for interview timing",
                description=(
                    "The form does not include an `audit` field. SurveyCTO's audit log "
                    "records time spent on each question, GPS track of the interview, and "
                    "back-navigation events — all useful for detecting enumerator fabrication "
                    "and identifying slow/confusing questions."
                ),
                action=(
                    "Add a row in the survey sheet with `type = audit` and a name "
                    "(e.g. `audit`). No label is needed. This adds negligible burden "
                    "and significantly strengthens high-frequency checks."
                ),
                example="type: audit    name: audit",
                vars=[],
                priority="Low",
            ))

    # ── 9. Section context notes ──────────────────────────────────────────────

    def _check_section_notes(self):
        questions_only = [q for q in self.parser.questions
                          if q["type"].split()[0] not in XLSFormParser.STRUCTURAL]
        if len(questions_only) < 10:
            return
        # Find stretches of non-note questions longer than 15
        streak = 0
        long_blocks = []
        block_start = None
        for q in questions_only:
            if q["type"] == "note":
                streak = 0
                block_start = None
            else:
                if streak == 0:
                    block_start = q["name"]
                streak += 1
                if streak == 15:
                    long_blocks.append(block_start)
        if long_blocks:
            self._add(Suggestion(
                category="Survey Flow",
                title="Long question blocks without section notes",
                description=(
                    f"The form contains blocks of 15 or more consecutive questions with no "
                    f"`note` row to provide context or section transitions. "
                    f"First such block starts near **{long_blocks[0]}**. "
                    "Enumerators benefit from brief notes that signal topic changes, "
                    "provide module-level instructions, or re-establish rapport."
                ),
                action=(
                    "Insert `note` rows at module boundaries (e.g. 'Now I will ask you "
                    "about household income.') and before sensitive sections with a brief "
                    "privacy assurance."
                ),
                vars=long_blocks[:3],
                priority="Low",
            ))

    # ── 10. Oversized choice lists ────────────────────────────────────────────

    def _check_large_choice_lists(self):
        for q in self.questions:
            ln = q.get("list_name")
            if not ln:
                continue
            choices = self.choices.get(ln, [])
            if len(choices) < 20:
                continue
            appearance = q.get("appearance", "")
            if "search" in appearance or "autocomplete" in appearance:
                continue
            self._add(Suggestion(
                category="Efficiency",
                title="Large choice list without autocomplete appearance",
                description=(
                    f"**{q['name']}** uses choice list `{ln}` with {len(choices)} options. "
                    "Scrolling through a long list is slow and error-prone on tablets. "
                    "SurveyCTO supports a searchable dropdown that narrows options as "
                    "the enumerator types."
                ),
                action=(
                    "Set `appearance = search` (SurveyCTO) or `autocomplete` (ODK) in the "
                    "survey sheet for this question to enable a filtered dropdown."
                ),
                example=f"appearance: search",
                vars=[q["name"]],
                priority="Medium",
            ))

    # ── 11. Repeat-group candidates ───────────────────────────────────────────

    def _check_repeat_candidates(self):
        """Flag groups of similarly-named questions that suggest a roster pattern."""
        from collections import defaultdict
        stem_map = defaultdict(list)
        suffix_re = re.compile(r"_(\d+)$")
        for q in self.questions:
            m = suffix_re.search(q["name"])
            if m:
                stem = q["name"][:m.start()]
                stem_map[stem].append(q["name"])
        for stem, members in stem_map.items():
            if len(members) >= 4:
                self._add(Suggestion(
                    category="Efficiency",
                    title="Repeated numbered questions — consider a repeat group",
                    description=(
                        f"Found {len(members)} questions with the pattern `{stem}_N` "
                        f"({', '.join(members[:4])}{'…' if len(members) > 4 else ''}). "
                        "This pattern often indicates a roster that is hard-coded as individual "
                        "questions, which makes the form inflexible and hard to maintain."
                    ),
                    action=(
                        f"Replace `{stem}_1` … `{stem}_{len(members)}` with a single question "
                        f"inside a `begin repeat` / `end repeat` block. Set the repeat count "
                        "using a prior count question (e.g. household size)."
                    ),
                    vars=members[:6],
                    priority="Medium",
                ))

    # ── 12. Broken ${varname} references ─────────────────────────────────────

    def _check_broken_variable_references(self):
        """Detect ${varname} refs in relevance/constraint/calculation that don't exist."""
        defined = {q["name"] for q in self.parser.questions if q.get("name")}
        ref_re  = re.compile(r"\$\{([^}]+)\}")
        expr_fields = ("relevant", "constraint", "calculation", "default")
        broken_by_var: dict = {}   # undefined_name → [(field, col)]

        for q in self.parser.questions:
            for col in expr_fields:
                expr = q.get(col, "")
                if not expr:
                    continue
                for m in ref_re.finditer(expr):
                    ref = m.group(1).strip()
                    if ref not in defined:
                        broken_by_var.setdefault(ref, []).append((q["name"], col))

        for ref, usages in broken_by_var.items():
            affected = list({u[0] for u in usages})
            cols_used = list({u[1] for u in usages})
            self._add(Suggestion(
                category="Reference Integrity",
                title=f"Undefined variable reference: ${{{ref}}}",
                description=(
                    f"The expression `${{{ref}}}` appears in the `{'`, `'.join(cols_used)}` "
                    f"column(s) of {len(affected)} field(s) but `{ref}` is not defined anywhere "
                    "in the survey sheet. This is almost certainly a typo and will cause the "
                    "condition to silently fail — the question may always show or always hide."
                ),
                action=(
                    f"Check whether `{ref}` is a misspelling of an existing field name. "
                    "Correct the reference or add the missing field."
                ),
                vars=affected[:8],
                priority="High",
            ))

    # ── 13. select_multiple choice names with spaces ──────────────────────────

    def _check_select_multiple_space_names(self):
        """Choice names containing spaces break select_multiple response parsing."""
        seen_lists = set()
        for q in self.questions:
            if not q["type"].startswith("select_multiple"):
                continue
            ln = q.get("list_name")
            if not ln or ln in seen_lists:
                continue
            seen_lists.add(ln)
            bad = [c for c in self.choices.get(ln, []) if " " in c]
            if bad:
                self._add(Suggestion(
                    category="Reference Integrity",
                    title="Choice names with spaces in a select_multiple list",
                    description=(
                        f"Choice list **{ln}** (used by `select_multiple`) contains choice "
                        f"name(s) with spaces: {', '.join(f'`{b}`' for b in bad[:5])}. "
                        "XLSForm stores multi-select responses as space-separated values, so "
                        "a choice name with a space will be split into two tokens during analysis, "
                        "corrupting the data silently."
                    ),
                    action=(
                        "Replace spaces in choice names with underscores "
                        "(e.g. `crop type` → `crop_type`). Labels can still contain spaces."
                    ),
                    vars=[q2["name"] for q2 in self.questions if q2.get("list_name") == ln],
                    priority="High",
                ))

    # ── 14. Required fields without required_message ──────────────────────────

    def _check_required_without_message(self):
        for q in self.questions:
            if q.get("required", "").lower() not in ("yes", "true", "1"):
                continue
            if q.get("required_message"):
                continue
            # Only flag answerable non-structural types
            base = q["type"].split()[0]
            if base in XLSFormParser.STRUCTURAL or base in ("note", "calculate"):
                continue
            self._add(Suggestion(
                category="Enumerator Guidance",
                title="Required field without a required_message",
                description=(
                    f"**{q['name']}** is marked required but has no `required_message`. "
                    "When the enumerator tries to advance without answering, SurveyCTO shows "
                    "a generic 'This field is required' prompt that gives no context."
                ),
                action=(
                    "Add a `required_message` that explains why the field is mandatory "
                    "and what the enumerator should do (e.g. 'This question must be answered. "
                    "If the respondent refuses, select Prefer not to answer.')."
                ),
                vars=[q["name"]],
                priority="Low",
            ))

    # ── 15. Read-only fields without a default ────────────────────────────────

    def _check_readonly_without_default(self):
        for q in self.questions:
            if q.get("read_only", "").lower() not in ("yes", "true", "1"):
                continue
            if q.get("default") or q.get("calculation"):
                continue
            self._add(Suggestion(
                category="Skip Logic",
                title="Read-only field with no default or calculation",
                description=(
                    f"**{q['name']}** is marked `read_only` but has neither a `default` "
                    "value nor a `calculation`. The field will always be blank and "
                    "uneditable — it will display nothing to the enumerator."
                ),
                action=(
                    "Either add a `default` value or a `calculation` expression to populate "
                    "the field, or remove the `read_only` flag if editing is intended."
                ),
                vars=[q["name"]],
                priority="Medium",
            ))

    # ── 16. Calculate fields that should use once() ───────────────────────────

    def _check_calculate_once(self):
        """Static calculations (no variable references) should use once() to avoid
        re-evaluation every time any field changes."""
        ref_re = re.compile(r"\$\{[^}]+\}")
        for q in self.questions:
            if q["type"] != "calculate":
                continue
            calc = q.get("calculation", "")
            if not calc:
                continue
            # Skip if already uses once() or today()/now()
            if "once(" in calc or "today()" in calc or "now()" in calc or "random()" in calc:
                continue
            # If the expression contains no variable references, it's purely static
            if not ref_re.search(calc) and len(calc) > 3:
                self._add(Suggestion(
                    category="Performance",
                    title="Static calculate field not wrapped in once()",
                    description=(
                        f"**{q['name']}** is a `calculate` field whose expression "
                        f"`{calc[:80]}{'…' if len(calc) > 80 else ''}` contains no variable "
                        "references. SurveyCTO re-evaluates all calculations every time any "
                        "field changes, so static calculations add unnecessary overhead. "
                        "Wrapping in `once()` evaluates the expression only at first load."
                    ),
                    action="Wrap the expression with `once()` to prevent repeated evaluation:",
                    example=f"once({calc[:60]}{'…)' if len(calc) > 60 else ')'}",
                    vars=[q["name"]],
                    priority="Low",
                ))

    # ── 17. pulldata() without type conversion ────────────────────────────────

    def _check_pulldata_type_conversion(self):
        """pulldata() returns text strings; arithmetic on them requires int()/number()."""
        arith_re = re.compile(r"pulldata\s*\(", re.I)
        math_ops  = re.compile(r"[\+\-\*\/]|div\b|mod\b|>=|<=|>|<", re.I)
        for q in self.questions:
            calc = q.get("calculation", "") or q.get("constraint", "")
            if not calc:
                continue
            if not arith_re.search(calc):
                continue
            # Check if pulldata result is used in arithmetic without wrapping
            if math_ops.search(calc) and "int(" not in calc and "number(" not in calc:
                self._add(Suggestion(
                    category="Reference Integrity",
                    title="pulldata() result used in arithmetic without type conversion",
                    description=(
                        f"**{q['name']}** uses `pulldata()` alongside arithmetic operators "
                        "but does not convert the result with `int()` or `number()`. "
                        "`pulldata()` always returns a text string; using it directly in "
                        "arithmetic will silently produce empty or NaN results in SurveyCTO."
                    ),
                    action=(
                        "Wrap the `pulldata()` call with `int()` or `number()` before "
                        "performing arithmetic:"
                    ),
                    example="int(pulldata('dataset', 'col', 'key_col', ${keyfield}))",
                    vars=[q["name"]],
                    priority="High",
                ))

    # ── 18. Likert questions without randomized appearance ────────────────────

    def _check_likert_appearance(self):
        """Opinion/attitude Likert questions should randomize choice order to
        reduce primacy/recency and acquiescence bias."""
        checked_lists = set()
        for q in self.questions:
            ln = q.get("list_name")
            if not ln or ln in checked_lists:
                continue
            labels_text = " | ".join(self._choice_labels(ln))
            if not self._LIKERT_LABELS.search(labels_text):
                continue
            checked_lists.add(ln)
            appearance = q.get("appearance", "")
            if "randomized" in appearance or "likert" in appearance:
                continue
            affected = [q2["name"] for q2 in self.questions if q2.get("list_name") == ln]
            self._add(Suggestion(
                category="Question Design",
                title="Likert/attitude scale without randomized choice order",
                description=(
                    f"Questions using choice list **{ln}** appear to use a Likert or "
                    "attitude scale but have no `randomized` appearance. Presenting response "
                    "options in a fixed order introduces primacy bias (first option selected "
                    "more often) and recency bias, affecting cross-respondent comparability."
                ),
                action=(
                    "Add `randomized` to the `appearance` column for these questions. "
                    "If 'Other' or 'Don't know' must stay at the bottom, use "
                    "`randomized(0, 1)` to exclude the last choice from randomization:"
                ),
                example="randomized(0, 1)",
                vars=affected[:6],
                priority="Medium",
            ))

    # ── 19. GPS fields without accuracy parameters ────────────────────────────

    def _check_geopoint_accuracy(self):
        for q in self.questions:
            if q["type"].split()[0] != "geopoint":
                continue
            params = q.get("parameters", "")
            if "capture-accuracy" in params or "accuracy" in params:
                continue
            self._add(Suggestion(
                category="Data Quality",
                title="GPS field without accuracy threshold",
                description=(
                    f"**{q['name']}** collects GPS coordinates but does not set "
                    "`capture-accuracy` or `warning-accuracy` parameters. Without these, "
                    "the form will accept any GPS reading regardless of precision, "
                    "potentially recording locations accurate only to hundreds of metres."
                ),
                action=(
                    "Add accuracy parameters in the `parameters` column. "
                    "`capture-accuracy` sets the required precision before the point is "
                    "recorded; `warning-accuracy` shows a warning but allows submission:"
                ),
                example="capture-accuracy=10 warning-accuracy=25",
                vars=[q["name"]],
                priority="Medium",
            ))

    # ── 20. Settings sheet completeness ──────────────────────────────────────

    def _check_settings_completeness(self):
        settings = self.parser.settings
        missing  = []
        if not settings.get("form_id"):
            missing.append(("form_id", "Unique form identifier — required for server upload and version management."))
        if not settings.get("form_title"):
            missing.append(("form_title", "Human-readable title shown on the device and server."))
        if not settings.get("version"):
            missing.append(("version", "Version string — recommended format `yyyymmddrr` (e.g. `2024060101`)."))

        for field, desc in missing:
            self._add(Suggestion(
                category="Form Metadata",
                title=f"Missing `{field}` in settings sheet",
                description=(
                    f"The settings sheet does not define `{field}`. {desc} "
                    "Without a form_id, SurveyCTO may not correctly track or de-duplicate "
                    "form versions on the server."
                ),
                action=(
                    f"Add a `{field}` column to the settings sheet with an appropriate value. "
                    + ("The version should follow `yyyymmddrr` format (year-month-day-revision)."
                       if field == "version" else "")
                ),
                example={"form_id": "my_survey_v1", "form_title": "Household Baseline Survey",
                         "version": "2024060101"}.get(field, ""),
                vars=[],
                priority="Medium" if field == "version" else "High",
            ))

        # Check version format if present
        ver = settings.get("version", "")
        if ver and not re.match(r"^\d{8,10}$", ver.replace("-", "").replace("_", "")):
            self._add(Suggestion(
                category="Form Metadata",
                title="Form version not in recommended yyyymmddrr format",
                description=(
                    f"The form version is `{ver}`. SurveyCTO recommends the format "
                    "`yyyymmddrr` (year + month + day + 2-digit revision number, e.g. "
                    "`2024060101`). This format sorts correctly and makes the release date "
                    "immediately visible on the server."
                ),
                action="Update the `version` field in the settings sheet to follow `yyyymmddrr`:",
                example="2024060101",
                vars=[],
                priority="Low",
            ))

    # ── 21. Multilingual completeness ─────────────────────────────────────────

    def _check_multilingual_completeness(self):
        langs = self.parser.label_langs
        if len(langs) < 2:
            return
        incomplete_by_lang: dict = {}
        for q in self.questions:
            base = q["type"].split()[0]
            if base in XLSFormParser.STRUCTURAL or base == "calculate":
                continue
            label_cols = q.get("_label_cols", {})
            for col, val in label_cols.items():
                if col.startswith("label:"):
                    lang = col.split(":", 1)[1].strip()
                    if not val or val.lower() in ("nan", ""):
                        incomplete_by_lang.setdefault(lang, []).append(q["name"])

        for lang, missing_vars in incomplete_by_lang.items():
            if len(missing_vars) == 0:
                continue
            self._add(Suggestion(
                category="Multilingual",
                title=f"Incomplete translation: {len(missing_vars)} fields missing `{lang}` label",
                description=(
                    f"The form has multilingual labels ({', '.join(langs)}) but "
                    f"**{len(missing_vars)}** question(s) have no `{lang}` translation. "
                    "When a respondent or enumerator selects that language, these questions "
                    "will display blank labels, potentially halting the interview."
                ),
                action=(
                    f"Fill in the `label:{lang}` column for all questions. "
                    "Use the printable form feature in SurveyCTO to cross-check coverage."
                ),
                vars=missing_vars[:8],
                priority="High" if len(missing_vars) > 5 else "Medium",
            ))

    # ── 22. Long-answer text without multiline appearance ─────────────────────

    def _check_text_multiline_appearance(self):
        open_ended = re.compile(
            r"\b(describe|explain|comment|reason|feedback|opinion|suggest|detail|"
            r"specify|elaborate|note|other|additional)\b", re.I
        )
        for q in self.questions:
            if q["type"] != "text":
                continue
            if "multiline" in q.get("appearance", ""):
                continue
            label = q.get("label", "")
            name  = q["name"]
            if open_ended.search(label) or open_ended.search(name):
                self._add(Suggestion(
                    category="Enumerator Guidance",
                    title="Open-ended text field without multiline appearance",
                    description=(
                        f"**{name}** appears to invite a long or free-text response "
                        "(label/name suggests: describe, explain, specify, etc.) but uses "
                        "a single-line text input by default. On mobile devices this is "
                        "uncomfortable for long answers and may discourage complete responses."
                    ),
                    action=(
                        "Set `appearance = multiline` in the survey sheet to expand the "
                        "input box and allow comfortable multi-line entry."
                    ),
                    example="appearance: multiline",
                    vars=[name],
                    priority="Low",
                ))

    # ── 23. Very large choice lists → CSV ─────────────────────────────────────

    def _check_very_large_choice_csv(self):
        seen = set()
        for q in self.questions:
            ln = q.get("list_name")
            if not ln or ln in seen:
                continue
            seen.add(ln)
            count = len(self.choices.get(ln, []))
            if count < 200:
                continue
            self._add(Suggestion(
                category="Performance",
                title=f"Choice list with {count} options should be a CSV dataset",
                description=(
                    f"Choice list **{ln}** has {count} options stored on the choices sheet. "
                    "SurveyCTO loads all choices into device memory when the form opens. "
                    "Lists of this size significantly slow form loading and navigation, "
                    "and are a documented cause of app crashes on low-end Android devices."
                ),
                action=(
                    "Move this choice list to a pre-loaded CSV dataset and use "
                    "`select_one_from_file` or a `search()` expression with `pulldata()`. "
                    "Also set `appearance = search` to enable filtering as the enumerator types."
                ),
                example=f"type: select_one_from_file {ln}.csv",
                vars=[q2["name"] for q2 in self.questions if q2.get("list_name") == ln],
                priority="High",
            ))

    # ── 24. Form complexity / performance warnings ────────────────────────────

    def _check_form_complexity(self):
        all_qs = self.parser.questions
        answerable = self.questions

        # Very long form
        if len(answerable) > 300:
            self._add(Suggestion(
                category="Performance",
                title=f"Very long form ({len(answerable)} questions) — consider splitting",
                description=(
                    f"The form has {len(answerable)} answerable questions. "
                    "SurveyCTO holds all fields in memory simultaneously regardless of "
                    "skip logic, and forms exceeding ~300–400 fields can load slowly and "
                    "cause crashes on low-end devices, especially with nested repeats."
                ),
                action=(
                    "Split the form into topical modules linked by a shared unique ID "
                    "(barcode scan, household ID, or `caseid` field). Each module loads "
                    "independently and can be merged server-side."
                ),
                vars=[],
                priority="Medium",
            ))

        # Many calculate fields
        calc_count = sum(1 for q in all_qs if q["type"] == "calculate")
        if calc_count > 50:
            self._add(Suggestion(
                category="Performance",
                title=f"High calculate field count ({calc_count} fields)",
                description=(
                    f"The form contains {calc_count} `calculate` fields. "
                    "Every time any field is answered, SurveyCTO re-evaluates all "
                    "calculation expressions that reference it. A large number of "
                    "chained calculations can cause noticeable lag between questions."
                ),
                action=(
                    "Wrap static calculations (no variable references) in `once()`. "
                    "Consolidate related calculations into fewer fields using nested "
                    "expressions. Avoid chains where calculation A feeds B feeds C feeds D."
                ),
                vars=[],
                priority="Low",
            ))

        # Repeat groups with many fields
        in_repeat    = False
        repeat_stack = 0
        repeat_field_count = 0
        repeat_name  = None
        for q in all_qs:
            t = q["type"].split()[0]
            if t in ("begin_repeat", "begin repeat"):
                repeat_stack += 1
                if repeat_stack == 1:
                    in_repeat = True
                    repeat_name = q["name"]
                    repeat_field_count = 0
            elif t in ("end_repeat", "end repeat"):
                if repeat_stack == 1 and repeat_field_count > 30:
                    self._add(Suggestion(
                        category="Performance",
                        title=f"Large repeat group '{repeat_name}' ({repeat_field_count} fields)",
                        description=(
                            f"Repeat group **{repeat_name}** contains {repeat_field_count} fields. "
                            "In SurveyCTO, each repeat instance multiplies the in-memory field "
                            "count (e.g. 40 fields × 20 repeats = 800 virtual fields). "
                            "High repeat counts with many fields are the most common cause "
                            "of form slowness and crashes."
                        ),
                        action=(
                            "Reduce fields inside the repeat to the minimum required. "
                            "Move fields that only need to be asked once (e.g. household-level "
                            "information) outside the repeat group."
                        ),
                        vars=[repeat_name],
                        priority="Medium",
                    ))
                repeat_stack = max(0, repeat_stack - 1)
                if repeat_stack == 0:
                    in_repeat = False
            elif in_repeat and repeat_stack == 1:
                repeat_field_count += 1

    # ── choice-label lookup ───────────────────────────────────────────────────

    def _choice_labels(self, list_name: str) -> list:
        if self.parser.choices_df is None or not list_name:
            return []
        df = self.parser.choices_df
        mask = df.get("list_name", pd.Series(dtype=str)) == list_name
        labels = []
        for col in df.columns:
            if col == "label" or col.startswith("label:"):
                vals = df.loc[mask, col].dropna().astype(str).tolist()
                labels.extend(vals)
        return labels


# ── UI helper: render suggestion cards ───────────────────────────────────────

_PRI_COLOR = {
    "High":   ("#cc9470", "#fdf5ee"),
    "Medium": ("#ccb460", "#fdf8e8"),
    "Low":    ("#6aab90", "#e8f6f0"),
}

_CAT_ICON = {
    "Reference Integrity":  "🔗",
    "Choice Logic":         "🔘",
    "Respondent Experience":"🤝",
    "Validation":           "🔢",
    "Skip Logic":           "↪️",
    "Enumerator Guidance":  "💬",
    "Question Design":      "✏️",
    "Data Quality":         "📊",
    "Survey Flow":          "📋",
    "Efficiency":           "⚡",
    "Performance":          "🐢",
    "Form Metadata":        "🏷️",
    "Multilingual":         "🌐",
}


def _show_suggestions(suggestions: list):
    if not suggestions:
        st.success("✅ No design improvement suggestions — the form follows best practices in all checked areas.")
        return

    by_cat = {}
    for s in suggestions:
        by_cat.setdefault(s.category, []).append(s)

    pri_counts = Counter(s.priority for s in suggestions)
    sc1, sc2, sc3, sc4 = st.columns(4)
    sc1.metric("Total Suggestions", len(suggestions),
               help="Design improvement opportunities identified. These are not errors — the form will work — but addressing them can improve data quality, reduce non-response, and simplify data cleaning.")
    sc2.metric("⬆ High Priority",   pri_counts.get("High", 0),
               help="Suggestions with a strong potential impact on data quality or respondent experience.")
    sc3.metric("— Medium Priority", pri_counts.get("Medium", 0),
               help="Worthwhile improvements that carry moderate impact.")
    sc4.metric("⬇ Low Priority",    pri_counts.get("Low", 0),
               help="Minor refinements and nice-to-have enhancements.")

    st.markdown("")

    for cat, items in sorted(by_cat.items()):
        icon = _CAT_ICON.get(cat, "📌")
        with st.expander(f"{icon} **{cat}** — {len(items)} suggestion{'s' if len(items) != 1 else ''}", expanded=True):
            for s in items:
                fg, bg = _PRI_COLOR.get(s.priority, ("#64748b", "#f8fafc"))
                vars_html = ""
                if s.vars:
                    chips = "".join(f'<code style="background:#f1f5f9;padding:.1rem .35rem;border-radius:3px;font-size:.78rem;margin:1px 2px">{v}</code>' for v in s.vars[:8])
                    more  = f'<span style="color:#94a3b8;font-size:.78rem"> +{len(s.vars)-8} more</span>' if len(s.vars) > 8 else ""
                    vars_html = f'<div style="margin-top:.4rem">{chips}{more}</div>'
                example_html = ""
                if s.example:
                    example_html = (
                        f'<div style="margin-top:.5rem;background:#f1f5f9;padding:.4rem .7rem;'
                        f'border-radius:5px;font-family:monospace;font-size:.82rem;color:#1e293b">'
                        f'{html_module.escape(s.example)}</div>'
                    )
                st.markdown(
                    f'<div style="border:1px solid {bg};border-left:4px solid {fg};'
                    f'border-radius:6px;padding:.8rem 1rem;margin-bottom:.7rem;background:white">'
                    f'<div style="display:flex;justify-content:space-between;align-items:flex-start">'
                    f'<b style="font-size:.92rem">{s.title}</b>'
                    f'<span style="background:{bg};color:{fg};font-size:.68rem;font-weight:700;'
                    f'padding:.2rem .55rem;border-radius:12px;white-space:nowrap;margin-left:.5rem">'
                    f'{s.priority}</span></div>'
                    f'<div style="font-size:.87rem;color:#334155;margin-top:.35rem;line-height:1.55">'
                    f'{s.description}</div>'
                    f'<div style="font-size:.84rem;color:#64748b;margin-top:.4rem">'
                    f'<b>Suggested action:</b> {s.action}</div>'
                    f'{example_html}{vars_html}'
                    f'</div>',
                    unsafe_allow_html=True,
                )


# ══════════════════════════════════════════════════════════════════════════════
# HTML REPORT GENERATOR
# ══════════════════════════════════════════════════════════════════════════════

class ReportGenerator:

    _SEV_COLOR = {
        "Critical": ("#cc8080","#fdf0f0"),
        "High":     ("#cc9470","#fdf5ee"),
        "Medium":   ("#ccb460","#fdf8e8"),
        "Low":      ("#6aab90","#e8f6f0"),
    }

    def generate(self, parser: XLSFormParser, results: list,
                 static_issues: list, sim_issues: list, num_sims: int) -> str:
        ts          = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        all_issues  = static_issues + sim_issues
        sev_counts  = Counter(i.severity for i in all_issues)
        path_lens   = [len(r.path_taken) for r in results] if results else []
        unique_paths = len(set(r.path_signature for r in results)) if results else 0
        stats       = parser.form_stats()

        return (
            self._open(ts)
            + self._header(ts)
            + self._kpi(num_sims, unique_paths, len(all_issues), sev_counts, stats)
            + self._source_breakdown(static_issues, sim_issues)
            + self._severity_chart(sev_counts)
            + (self._path_chart(path_lens, num_sims, unique_paths) if path_lens else "")
            + self._issues_table(all_issues)
            + self._footer(ts)
            + "</div></body></html>"
        )

    def _open(self, ts):
        return (f'<!DOCTYPE html><html lang="en"><head>'
                f'<meta charset="UTF-8">'
                f'<meta name="viewport" content="width=device-width,initial-scale=1.0">'
                f'<title>XLSForm Quality Assurance Report — {ts}</title>'
                f'{self._css()}</head><body><div class="container">\n')

    def _header(self, ts):
        return (f'<div class="header"><div class="hicon">📋</div><div>'
                f'<h1>XLSForm Quality Assurance Report</h1>'
                f'<p>Generated {ts} &nbsp;·&nbsp; XLSForm Quality Reviewer &nbsp;·&nbsp; '
                f'Standards: World Bank DIME ietestform &amp; IPA ipacheckscto</p>'
                f'</div></div>\n')

    def _kpi(self, num_sims, unique_paths, total, sev_counts, stats):
        ic = "#cc8080" if total else "#6aab90"
        kpis = [
            (num_sims,                   "Simulations",    "#6a9cc8"),
            (unique_paths,               "Unique Paths",   "#6ab0c8"),
            (stats["answerable"],        "Questions",      "#9888c8"),
            (total,                      "Total Issues",   ic),
            (sev_counts.get("Critical",0),"Critical",      "#cc8080"),
            (sev_counts.get("High",0),   "High",           "#cc9470"),
            (sev_counts.get("Medium",0), "Medium",         "#ccb460"),
            (sev_counts.get("Low",0),    "Low",            "#6aab90"),
        ]
        cards = "".join(
            f'<div class="kpi"><div class="kv" style="color:{c}">{v}</div>'
            f'<div class="kl">{l}</div></div>' for v, l, c in kpis
        )
        return f'<div class="kpi-row">{cards}</div>\n'

    def _source_breakdown(self, static_issues, sim_issues):
        total = len(static_issues) + len(sim_issues)
        if total == 0: return ""
        sp = len(static_issues) / total * 100
        smp = len(sim_issues) / total * 100
        return (
            f'<div class="card"><h2>Issue Sources</h2>'
            f'<div class="sbar"><span class="sbl" style="color:#6a9cc8">Static</span>'
            f'<div class="strk"><div class="sfil" style="width:{sp:.0f}%;background:#6a9cc8"></div></div>'
            f'<span class="scnt" style="color:#6a9cc8">{len(static_issues)}</span></div>'
            f'<div class="sbar"><span class="sbl" style="color:#9888c8">Simulation</span>'
            f'<div class="strk"><div class="sfil" style="width:{smp:.0f}%;background:#9888c8"></div></div>'
            f'<span class="scnt" style="color:#9888c8">{len(sim_issues)}</span></div>'
            f'<p class="sub" style="margin-top:.5rem">Static checks: structural analysis on form design. '
            f'Simulation checks: logic issues found by simulating respondent paths.</p>'
            f'</div>\n'
        )

    def _severity_chart(self, sev_counts):
        max_v = max(sev_counts.values(), default=1)
        bars  = ""
        for sev in ("Critical","High","Medium","Low"):
            cnt = sev_counts.get(sev, 0)
            pct = cnt / max_v * 100 if max_v else 0
            fg, _ = self._SEV_COLOR.get(sev, ("#6b7280","#f3f4f6"))
            bars += (f'<div class="sbar"><span class="sbl" style="color:{fg}">{sev}</span>'
                     f'<div class="strk"><div class="sfil" style="width:{pct:.1f}%;background:{fg}"></div></div>'
                     f'<span class="scnt" style="color:{fg}">{cnt}</span></div>')
        return f'<div class="card"><h2>Issue Severity Breakdown</h2>{bars}</div>\n'

    def _path_chart(self, path_lens, num_sims, unique_paths):
        mn, mx = min(path_lens), max(path_lens)
        avg    = sum(path_lens) / len(path_lens)
        n_bkt  = min(12, mx - mn + 1) if mx > mn else 1
        bkt_sz = max(1, (mx - mn + 1) // n_bkt)
        bkts   = Counter((l - mn) // bkt_sz for l in path_lens)
        max_b  = max(bkts.values(), default=1)
        bars   = ""
        for i in range(n_bkt):
            cnt = bkts.get(i, 0)
            pct = cnt / max_b * 100
            lo  = mn + i * bkt_sz
            hi  = lo + bkt_sz - 1
            bars += (f'<div class="pbar"><span class="pbl">{lo}–{hi}q</span>'
                     f'<div class="ptrk"><div class="pfil" style="width:{pct:.1f}%"></div></div>'
                     f'<span class="pcnt">{cnt}</span></div>')
        div_pct = unique_paths / num_sims * 100
        return (f'<div class="card"><h2>Survey Path Length Distribution</h2>'
                f'<p class="sub">Min: <strong>{mn}</strong> &nbsp;|&nbsp; '
                f'Max: <strong>{mx}</strong> &nbsp;|&nbsp; Avg: <strong>{avg:.1f}</strong> '
                f'questions &nbsp;|&nbsp; Path diversity: <strong>{div_pct:.0f}%</strong></p>'
                f'{bars}</div>\n')

    def _issues_table(self, issues):
        if not issues:
            return ('<div class="card"><h2>Detected Issues</h2>'
                    '<div class="ok"><div class="ok-icon">✓</div>'
                    '<strong>No issues detected.</strong>'
                    '<p>The form passed all quality checks in this category.</p></div></div>\n')

        MAX_VARS = 6
        # Group by (issue_type, severity, source)
        groups: dict = {}
        for iss in issues:
            key = (iss.issue_type, iss.severity, iss.source)
            groups.setdefault(key, []).append(iss)

        sev_ord = {"Critical": 0, "High": 1, "Medium": 2, "Low": 3}
        sorted_groups = sorted(
            groups.items(),
            key=lambda x: (sev_ord.get(x[0][1], 3), x[0][0])
        )

        n_types = len(sorted_groups)
        rows = ""
        for i, ((issue_type, severity, source), grp) in enumerate(sorted_groups, 1):
            fg, bg   = self._SEV_COLOR.get(severity, ("#6b7280", "#f3f4f6"))
            example  = grp[0]
            esc_qs   = [html_module.escape(iss.question) for iss in grp]

            if len(esc_qs) <= MAX_VARS:
                vars_html = " ".join(f'<code class="qn">{q}</code>' for q in esc_qs)
            else:
                shown     = esc_qs[:MAX_VARS]
                remaining = len(esc_qs) - MAX_VARS
                vars_html = (
                    " ".join(f'<code class="qn">{q}</code>' for q in shown)
                    + f' <span style="color:#94a3b8;font-style:italic">+{remaining} more</span>'
                )

            src_badge = ('<span class="src-badge src-static">static</span>'
                         if source == "static"
                         else '<span class="src-badge src-sim">simulation</span>')

            rows += (
                f'<tr><td class="num">{i}</td>'
                f'<td><span class="badge" style="color:{fg};background:{bg}">{severity}</span></td>'
                f'<td><strong>{html_module.escape(issue_type)}</strong><br>{src_badge}</td>'
                f'<td style="text-align:center;font-weight:700;color:{fg}">{len(grp)}</td>'
                f'<td style="max-width:220px">{vars_html}</td>'
                f'<td><div>{html_module.escape(example.explanation)}</div>'
                f'<div class="fix">💡 {html_module.escape(example.suggestion)}</div></td></tr>'
            )

        return (
            f'<div class="card"><h2>Detected Issues ({len(issues)} total · {n_types} problem types)</h2>'
            f'<p class="sub">Issues are grouped by type — the <em># Vars</em> column shows how many '
            f'variables share the same problem. Fix one, then apply the same fix to all listed variables.</p>'
            f'<div class="tbl-wrap"><table>'
            f'<thead><tr><th>#</th><th>Severity</th><th>Issue Type</th>'
            f'<th style="text-align:center"># Vars</th>'
            f'<th>Affected Variables</th>'
            f'<th>What It Means &amp; How to Fix</th></tr></thead>'
            f'<tbody>{rows}</tbody></table></div></div>\n'
        )

    def _footer(self, ts):
        return (f'<div class="foot">XLSForm Quality Reviewer &nbsp;·&nbsp; Quality Assurance Report &nbsp;·&nbsp; {ts}'
                f'<br>Standards: World Bank DIME ietestform &amp; IPA ipacheckscto</div>\n')

    def _css(self):
        return """<style>
:root{--primary:#6a9cc8;--pl:#d8eaf8;--bg:#f5f8fc;--card:#fff;
  --bdr:#e2e8f0;--tx:#1e293b;--mu:#64748b}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;
  background:var(--bg);color:var(--tx);line-height:1.6}
.container{max-width:1100px;margin:0 auto;padding:2rem 1.5rem}
.header{display:flex;align-items:center;gap:1.1rem;
  background:linear-gradient(135deg,#6a9cc8 0%,#7aacdc 100%);
  color:#fff;padding:1.75rem 2rem;border-radius:12px;margin-bottom:1.75rem}
.hicon{font-size:2.4rem}
.header h1{font-size:1.6rem;font-weight:800}
.header p{opacity:.82;font-size:.84rem;margin-top:.15rem}
.kpi-row{display:grid;grid-template-columns:repeat(auto-fit,minmax(110px,1fr));
  gap:.9rem;margin-bottom:1.5rem}
.kpi{background:var(--card);border:1px solid var(--bdr);border-radius:10px;
  padding:1rem;text-align:center;box-shadow:0 1px 4px rgba(0,0,0,.05)}
.kv{font-size:1.9rem;font-weight:800}
.kl{font-size:.68rem;text-transform:uppercase;letter-spacing:.06em;
  color:var(--mu);margin-top:.2rem}
.card{background:var(--card);border:1px solid var(--bdr);border-radius:10px;
  padding:1.5rem;margin-bottom:1.5rem;box-shadow:0 1px 4px rgba(0,0,0,.05)}
.card h2{font-size:1.05rem;font-weight:700;color:var(--primary);
  margin-bottom:1rem;padding-bottom:.5rem;border-bottom:2px solid var(--pl)}
.sub{font-size:.84rem;color:var(--mu);margin-bottom:1rem}
.sbar,.pbar{display:flex;align-items:center;gap:.75rem;margin-bottom:.6rem}
.sbl,.pbl{width:80px;font-weight:600;font-size:.84rem;text-align:right;flex-shrink:0}
.strk,.ptrk{flex:1;background:#f1f5f9;border-radius:4px;height:20px;overflow:hidden}
.sfil,.pfil{height:100%;border-radius:4px;min-width:2px}
.pfil{background:#7aa4cc}
.scnt,.pcnt{width:40px;font-weight:700;font-size:.84rem;flex-shrink:0}
.tbl-wrap{overflow-x:auto}
table{width:100%;border-collapse:collapse;font-size:.84rem}
th{background:#f1f5f9;text-align:left;padding:.55rem .75rem;font-weight:600;
  color:var(--mu);text-transform:uppercase;letter-spacing:.04em;
  font-size:.7rem;border-bottom:2px solid var(--bdr)}
td{padding:.6rem .75rem;border-bottom:1px solid var(--bdr);vertical-align:top}
tr:last-child td{border-bottom:none}
tr:hover td{background:#f8fafc}
td.num{color:var(--mu);font-size:.76rem;text-align:center;width:28px}
.badge{display:inline-block;padding:.2rem .55rem;border-radius:20px;
  font-size:.7rem;font-weight:700;white-space:nowrap}
code.qn{font-family:"SF Mono","Fira Code",monospace;background:#f1f5f9;
  padding:.15rem .4rem;border-radius:4px;font-size:.79rem}
.fix{color:var(--mu);font-size:.81rem;margin-top:.35rem}
.src-badge{display:inline-block;font-size:.65rem;font-weight:600;padding:.1rem .4rem;
  border-radius:4px;margin-top:.25rem}
.src-static{background:#e8f2fc;color:#6a9cc8}
.src-sim{background:#f2eefc;color:#9888c8}
.ok{text-align:center;padding:3rem 1rem;color:var(--mu)}
.ok-icon{font-size:2.5rem;color:#6aab90}
.ok strong{color:#6aab90;font-size:1.1rem;display:block;margin:.5rem 0}
.foot{text-align:center;color:var(--mu);font-size:.76rem;
  margin-top:2rem;padding-top:1rem;border-top:1px solid var(--bdr)}
</style>"""


# ══════════════════════════════════════════════════════════════════════════════
# STREAMLIT APP
# ══════════════════════════════════════════════════════════════════════════════

def _sev_badge(sev: str) -> str:
    colours = {
        "Critical": "background:#fdf0f0;color:#cc8080;font-weight:700",
        "High":     "background:#fdf5ee;color:#cc9470;font-weight:700",
        "Medium":   "background:#fdf8e8;color:#ccb460;font-weight:700",
        "Low":      "background:#e8f6f0;color:#6aab90;font-weight:700",
    }
    return colours.get(sev, "")


def _show_issue_table(issues: list, key_suffix: str = ""):
    """Display issues grouped by type — one row per unique problem, listing all affected variables."""
    if not issues:
        st.success("✓ No issues found in this category.")
        return

    SEV_COLOR = {
        "Critical": ("#cc8080", "#fdf0f0"),
        "High":     ("#cc9470", "#fdf5ee"),
        "Medium":   ("#ccb460", "#fdf8e8"),
        "Low":      ("#6aab90", "#e8f6f0"),
    }
    SEV_ORD = {"Critical": 0, "High": 1, "Medium": 2, "Low": 3}
    MAX_VARS = 6   # variables shown inline before "+N more"

    # Group by (issue_type, severity, source) — same problem, many variables
    groups: dict = {}
    for iss in issues:
        key = (iss.issue_type, iss.severity, iss.source)
        groups.setdefault(key, []).append(iss)

    sorted_groups = sorted(
        groups.items(),
        key=lambda x: (SEV_ORD.get(x[0][1], 3), x[0][0])
    )

    n_types  = len(sorted_groups)
    n_total  = len(issues)
    saved    = n_total - n_types
    note     = (f"{n_total} issues across **{n_types} problem types**"
                + (f" — grouping saves {saved} duplicate rows" if saved > 0 else ""))
    st.caption(note)

    rows_html = ""
    for (issue_type, severity, source), grp in sorted_groups:
        fg, bg   = SEV_COLOR.get(severity, ("#6b7280", "#f3f4f6"))
        example  = grp[0]
        esc_qs   = [html_module.escape(iss.question) for iss in grp]

        if len(esc_qs) <= MAX_VARS:
            vars_html = " ".join(f"<code>{q}</code>" for q in esc_qs)
        else:
            shown     = esc_qs[:MAX_VARS]
            remaining = len(esc_qs) - MAX_VARS
            vars_html = (
                " ".join(f"<code>{q}</code>" for q in shown)
                + f' <span class="xit-more">+{remaining} more</span>'
            )

        src_cls   = "xit-src-s" if source == "static" else "xit-src-m"
        src_label = "static"    if source == "static" else "simulation"

        rows_html += f"""
        <tr>
          <td style="white-space:nowrap">
            <span class="xit-badge" style="color:{fg};background:{bg}">{severity}</span>
          </td>
          <td>
            <strong>{html_module.escape(issue_type)}</strong><br>
            <span class="xit-src {src_cls}">{src_label}</span>
          </td>
          <td class="xit-count" style="color:{fg}">{len(grp)}</td>
          <td class="xit-vars" style="max-width:280px">{vars_html}</td>
          <td>
            {html_module.escape(example.explanation)}
            <div class="xit-fix">💡 {html_module.escape(example.suggestion)}</div>
          </td>
        </tr>"""

    st.markdown(f"""
    <div style="overflow-x:auto">
      <table class="xit-table">
        <thead><tr>
          <th>Severity</th>
          <th>Issue Type</th>
          <th style="text-align:center" title="Number of variables affected"># Vars</th>
          <th>Affected Variables</th>
          <th>What It Means &amp; How to Fix</th>
        </tr></thead>
        <tbody>{rows_html}</tbody>
      </table>
    </div>
    """, unsafe_allow_html=True)


def _show_landing():
    st.markdown("""
<div style="background:linear-gradient(135deg,#6a9cc8,#7aacdc);color:white;
  padding:2rem;border-radius:12px;margin-bottom:1.5rem">
  <h2 style="margin:0 0 .5rem 0">Upload an XLSForm (.xlsx) to begin analysis</h2>
  <p style="opacity:.85;margin:0">The file must contain a <code>survey</code> sheet with
  <code>type</code> and <code>name</code> columns. A <code>choices</code> sheet is required
  for select-type questions.</p>
</div>
""", unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("**Structural Checks** *(World Bank DIME + IPA)*")
        st.markdown("""
- Group begin/end structure validation
- Duplicate and missing choice codes
- Field name length (Stata & SurveyCTO limits)
- Outdated syntax and or_other usage
- Missing metadata and recommended columns
- Disabled and read-only field flags
- Constraint contradiction detection
""")
    with c2:
        st.markdown("**Simulation-Based Checks**")
        st.markdown("""
- Unreachable questions
- Circular relevance dependencies
- Deep skip chains (exceeding threshold)
- Required questions that cannot be answered
- Questions rarely reached across simulations
""")
    with c3:
        st.markdown("**Exported Report Includes**")
        st.markdown("""
- Summary metrics dashboard
- Structural vs. simulation issue breakdown
- Severity distribution chart
- Response pathway diversity analysis
- Comprehensive issues table with remediation guidance
""")


def main():
    # ── Sidebar ───────────────────────────────────────────────────────────────
    with st.sidebar:
        st.markdown("## ⚙️ Configuration")
        num_sims = st.number_input(
            "Number of simulations", min_value=1, max_value=500, value=50, step=10,
            help="A higher number of simulations improves coverage of conditional logic branches. 50–100 simulations is recommended for most forms.",
        )
        st.markdown("---")
        st.markdown("**Severity Classification**")
        for sev, emoji, desc in [
            ("Critical", "🔴", "Errors that will cause the form to fail or block data collection entirely"),
            ("High",     "🟠", "Issues likely to result in data loss, collection errors, or platform rejection"),
            ("Medium",   "🟡", "Issues that reduce data quality or may cause enumerator confusion"),
            ("Low",      "🟢", "Deviations from best-practice recommendations"),
        ]:
            st.markdown(f"{emoji} **{sev}** — {desc}")
        st.markdown("---")
        st.markdown("**Standards & References**")
        st.caption("World Bank DIME · ietestform.ado")
        st.caption("IPA · ipacheckscto.ado")
        st.caption("SurveyCTO XLSForm specification")

    # ── Main ──────────────────────────────────────────────────────────────────
    st.title("📋 XLSForm Quality Reviewer")
    st.markdown("*Structural and simulation-based quality assurance for household survey instruments.*")

    uploaded = st.file_uploader(
        "Upload XLSForm (.xlsx)", type=["xlsx"],
        help="The file must contain a 'survey' sheet with 'type' and 'name' columns. A 'choices' sheet is required for select-type questions.",
    )

    if uploaded is None:
        _show_landing()
        return

    # ── Parse ─────────────────────────────────────────────────────────────────
    with st.spinner("Parsing XLSForm…"):
        try:
            parser = XLSFormParser(uploaded)
        except Exception as e:
            st.error(f"❌ Failed to parse XLSForm: {e}")
            return

    if not parser.questions:
        st.error("No questions found. Check that 'type' and 'name' columns exist in the survey sheet.")
        return

    stats = parser.form_stats()
    st.success(f"✓ Successfully parsed **{stats['answerable']}** answerable questions from `{uploaded.name}`")

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Questions",             stats["answerable"],
              help="Total answerable questions in the survey sheet. Excludes structural rows (begin/end group, begin/end repeat), notes, and calculate fields.")
    c2.metric("Choice Lists",          stats["choice_lists"],
              help="Number of unique choice lists defined in the Choices sheet. Each select_one and select_multiple question references one of these lists.")
    c3.metric("Relevance Conditions",  stats["skip_logic_count"],
              help="Questions containing a 'relevant' or 'relevance' condition — displayed only when prior responses meet specified criteria. A higher count indicates more complex conditional branching.")
    c4.metric("Validation Rules",      stats["constraint_count"],
              help="Questions with a 'constraint' expression that rejects responses outside defined parameters (e.g. age ≥ 0 and age ≤ 120). Validation rules reduce data entry errors during collection.")
    c5.metric("Required Fields",       stats["required_count"],
              help="Questions marked as required, preventing the enumerator from advancing without a response. A high proportion of required fields may increase enumerator burden.")

    with st.expander("📊 Question Type Breakdown", expanded=False):
        tdf = (pd.DataFrame(list(stats["types"].items()), columns=["Type","Count"])
               .sort_values("Count", ascending=False).reset_index(drop=True))
        st.dataframe(tdf, hide_index=True, use_container_width=True)

    # ── Static analysis (runs immediately, no button) ─────────────────────────
    st.divider()
    st.subheader("🔬 Static Analysis")
    st.markdown(
        "Inspects the **form structure** without execution, identifying definite problems "
        "such as broken group nesting, duplicate field names, missing labels, and outdated syntax. "
        "Checks are derived from **World Bank DIME ietestform** and **IPA ipacheckscto** specifications."
    )

    with st.spinner("Running structural checks…"):
        static_issues = StaticAnalyzer(parser).analyze_all()

    sev_static = Counter(i.severity for i in static_issues)
    sc1, sc2, sc3, sc4, sc5 = st.columns(5)
    sc1.metric("Total Issues", len(static_issues),
               help="Total structural issues identified. These represent definite problems in the form design, independent of how the form is completed by respondents.")
    sc2.metric("🔴 Critical",  sev_static.get("Critical", 0),
               help="Errors that will cause the form to fail: mismatched begin/end groups, duplicate field names, unmatched repeat blocks. All critical issues must be resolved prior to deployment.")
    sc3.metric("🟠 High",      sev_static.get("High", 0),
               help="Issues likely to result in data collection failures: whitespace in field names, illegal characters, missing required metadata. Resolve prior to field work.")
    sc4.metric("🟡 Medium",    sev_static.get("Medium", 0),
               help="Issues that reduce data quality or may cause enumerator confusion: missing constraint messages, unused choice lists, disabled fields. Resolve prior to enumerator training.")
    sc5.metric("🟢 Low",       sev_static.get("Low", 0),
               help="Deviations from best-practice recommendations: naming conventions, optional but recommended columns. Address as time permits.")

    _show_issue_table(static_issues, key_suffix="static")

    # ── Simulation ────────────────────────────────────────────────────────────
    st.divider()
    st.subheader("🎲 Respondent Path Simulation")
    st.markdown(
        "Executes simulated respondents through the form, following relevance conditions based on "
        "generated responses. Each simulated respondent is assigned a correlated demographic profile "
        "(age, education level, marital status, location, income) and a **response style** — neutral, "
        "agreeable, cautious, or extreme — which determines response distributions on attitudinal and "
        "Likert-scale questions. "
        "This method identifies logic issues not detectable through structural analysis alone: "
        "broken relevance chains, unreachable questions, and conditions that evaluate identically "
        "regardless of input variation."
    )

    run_clicked = st.button(
        f"▶ Run {num_sims} Simulations", type="primary", use_container_width=False,
    )

    sim_issues: list    = []
    results: list       = []

    if run_clicked:
        engine   = SimulationEngine(parser)
        progress = st.progress(0, text="Starting…")
        status   = st.empty()

        for i in range(num_sims):
            results.append(engine.run(sim_id=i + 1, seed=i * 13 + 7))
            pct = (i + 1) / num_sims
            progress.progress(pct, text=f"Simulation {i+1} / {num_sims}")
            if (i + 1) % max(1, num_sims // 10) == 0:
                status.text(f"Completed {i+1}/{num_sims}…")

        progress.progress(1.0, text="Simulations complete.")
        status.empty()

        with st.spinner("Analysing paths…"):
            sim_issues = IssueDetector(parser, results).detect_all()

        sev_sim      = Counter(i.severity for i in sim_issues)
        path_lens    = [len(r.path_taken) for r in results]
        unique_paths = len(set(r.path_signature for r in results))
        avg_path     = sum(path_lens) / len(path_lens) if path_lens else 0
        total_qs     = stats["answerable"]
        div_pct      = unique_paths / num_sims

        m1, m2, m3, m4, m5, m6 = st.columns(6)
        m1.metric("Simulation Issues",  len(sim_issues),
                  help="Total issues identified by running simulated respondents through the form. Simulation detects logic problems — such as broken skip chains or unreachable questions — that structural analysis cannot.")
        m2.metric("Unique Paths",       unique_paths,
                  help=f"Number of distinct question sequences observed across {num_sims} simulations. A higher count indicates that conditional logic is producing meaningful response pathway variation. A minimum of 10% path diversity relative to simulation count is recommended.")
        m3.metric("Mean Path Length",   f"{avg_path:.1f} questions" if path_lens else "—",
                  help=f"Average number of questions answered per simulated respondent, out of {total_qs} total answerable questions. A substantially lower value indicates that a significant proportion of the form is conditionally skipped.")
        m4.metric("🔴 Critical",  sev_sim.get("Critical", 0),
                  help="Unanswerable required questions or completely broken skip chains that would block data entry.")
        m5.metric("🟠 High",      sev_sim.get("High", 0),
                  help="Deep skip chains (>10 consecutive skips) that may trap or confuse respondents.")
        m6.metric("🟡 Medium",    sev_sim.get("Medium", 0),
                  help="Questions rarely or never reached by any simulated respondent — may indicate unreachable logic branches.")

        # ── Path diversity interpretation ─────────────────────────────────────
        skip_pct = 1 - (avg_path / total_qs) if total_qs else 0
        st.caption(
            f"Average respondent answered **{avg_path:.0f}** of **{total_qs}** questions "
            f"({skip_pct:.0%} skipped on average)."
        )
        if div_pct >= 0.5:
            st.success(
                f"✅ **High path diversity** — {unique_paths} distinct pathways observed across {num_sims} simulations "
                f"({div_pct:.0%}). Conditional logic is producing meaningful response pathway variation."
            )
        elif div_pct >= 0.2:
            st.info(
                f"ℹ️ **Moderate path diversity** — {unique_paths} distinct pathways observed across {num_sims} simulations "
                f"({div_pct:.0%}). The form exhibits meaningful branching across respondent profiles."
            )
        elif div_pct >= 0.05:
            st.warning(
                f"⚠️ **Low path diversity** — {unique_paths} distinct pathways observed across {num_sims} simulations "
                f"({div_pct:.0%}). The majority of respondents follow an identical question sequence. Review relevance conditions."
            )
        else:
            st.error(
                f"🔴 **Insufficient path diversity** — {unique_paths} distinct pathways observed across {num_sims} simulations "
                f"({div_pct:.0%}). Relevance conditions may not be evaluating as intended."
            )

        st.subheader("🔍 Simulation Issues")
        _show_issue_table(sim_issues, key_suffix="sim")

        with st.expander("📈 Path Distribution Details", expanded=False):
            sig_counts = Counter(r.path_signature for r in results)
            ca, cb = st.columns(2)
            with ca:
                st.metric("Distinct pathways",  unique_paths)
                st.metric("Path length range",  f"{min(path_lens)}–{max(path_lens)} questions")
            with cb:
                top = pd.DataFrame([
                    {"Rank": i+1, "Path Length": len(p), "Occurrences": c}
                    for i, (p, c) in enumerate(sig_counts.most_common(10))
                ])
                st.dataframe(top, hide_index=True, use_container_width=True)

    # ── Design Advisor ────────────────────────────────────────────────────────
    st.divider()
    st.subheader("💡 Design Improvement Suggestions")
    st.markdown(
        "Recommendations grounded in **applied economics and social science survey methodology**. "
        "These are not errors — the form will function as written — but addressing them can "
        "improve data quality, reduce item non-response, and simplify downstream cleaning. "
        "Checks cover exclusive-choice constraints, numeric validation bounds, implicit follow-up "
        "logic, hint coverage, label quality, Likert scale balance, and survey flow efficiency."
    )

    with st.spinner("Analysing design…"):
        suggestions = DesignAdvisor(parser).analyze_all()

    _show_suggestions(suggestions)

    # ── Report export ─────────────────────────────────────────────────────────
    st.divider()
    st.subheader("📥 Export Quality Report")

    all_issues = static_issues + sim_issues

    with st.spinner("Generating report…"):
        html_report = ReportGenerator().generate(
            parser, results, static_issues, sim_issues,
            num_sims if results else 0,
        )

    ts       = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"xlsform_quality_report_{ts}.html"

    col_dl, col_info = st.columns([1, 2])
    with col_dl:
        st.download_button(
            label="⬇️ Download Report (.html)",
            data=html_report.encode("utf-8"),
            file_name=filename,
            mime="text/html",
            use_container_width=True,
        )
    with col_info:
        st.caption(
            f"**{filename}** — {len(all_issues)} total issues "
            f"({len(static_issues)} structural, {len(sim_issues)} simulation-detected). "
            f"Open in any web browser to view the formatted report."
        )

    with st.expander("Report Preview", expanded=False):
        st.code(html_report[:3000] + "\n<!-- … truncated … -->", language="html")


if __name__ == "__main__":
    main()
