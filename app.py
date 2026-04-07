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
        self._parse(file_obj)

    def _parse(self, file_obj):
        xl = pd.ExcelFile(file_obj)
        sheet_map = {n.lower(): n for n in xl.sheet_names}

        if "survey" in sheet_map:
            df = xl.parse(sheet_map["survey"])
            df.columns = [c.strip().lower() for c in df.columns]
            self.survey_df = df

        if "choices" in sheet_map:
            df = xl.parse(sheet_map["choices"])
            df.columns = [c.strip().lower() for c in df.columns]
            self.choices_df = df
            self._build_choices()

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
                "read_only":          self._val(row, "read_only"),
                "choices":            [],
                "list_name":          None,
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
