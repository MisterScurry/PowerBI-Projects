import re

import pandas as pd

INPUT_FILE = "Survey.xlsx"
OUTPUT_FILE = "Survey_Clean.xlsx"

# Column names (after renaming)
ROLE_COL = "Which Title Best Fits your Current Role?"
INDUSTRY_COL = "What Industry do you work in?"
COUNTRY_COL = "Which Country do you live in?"
LANGUAGE_COL = "Favorite Programming Language"
PRIORITY_COL = "If you were to look for a new job today, what would be the most important thing to you?"
ETHNICITY_COL = "Ethnicity"

COLUMNS_TO_DROP = [
    "Time Taken (America/New_York)",
    "Browser",
    "OS",
    "City",
    "Country",
    "Referrer",
    "Time Spent",
]

INDUSTRY_MAP = {
    r"financ(e|al)": "Finance",
    r"manufact":     "Manufacturing",
    r"non\s*profit": "Nonprofit",
    r"consulti":     "Consulting",
}

# Order matters: the first matching pattern wins
ROLE_MAP = {
    r"data\s*engineer":                 "Data Engineer",
    r"software\s*(developer|engineer)": "Software Developer",
    r"\bai\b|machine learning|\bml\b":  "AI/ML Engineer",
    r"data\s*scien":                    "Data Scientist",
    r"data\s*analy":                    "Data Analyst",
    r"business\s*analy":                "Business Analyst",
    **INDUSTRY_MAP,
}

COUNTRY_MAP = {
    r"nigeri":       "Nigeria",
    r"argentin":     "Argentina",
    r"brazi":        "Brazil",
    r"^fin":         "Finland",
    r"^(ire|irel)$": "Ireland",
    r"^leba":        "Lebanon",
    r"perú":         "Peru",
    r"portug":       "Portugal",
    r"^uae$":        "United Arab Emirates",
    r"^uzb":         "Uzbekistan",
}


def load_survey(path=INPUT_FILE):
    """Read the survey and keep only the text after the last '-' in each column name."""
    df = pd.read_excel(path)
    df.columns = [col.split("-")[-1].strip() for col in df.columns]
    return df


def strip_other_prefix(val):
    """'Other (Please Specify): Data Engineer' -> 'data engineer'."""
    if pd.isna(val):
        return val
    return str(val).split(":")[-1].strip().lower()


def make_standardizer(mapping):
    """Return a function mapping free text to a standard label via regex,
    falling back to Title Case when nothing matches."""
    def standardize(val):
        if pd.isna(val) or val == "":
            return val
        for pattern, label in mapping.items():
            if re.search(pattern, val):
                return label
        return val.title()
    return standardize


def clean_and_standardize(series, mapping):
    return series.apply(strip_other_prefix).apply(make_standardizer(mapping))


def clean_language(val):
    if pd.isna(val):
        return val
    val = str(val).strip()
    if "sql" in val.lower():  # check SQL first so "PostgreSQL", "MySQL", etc. collapse to SQL
        return "SQL"
    return val.split(":")[0].strip().title()


def collapse_other(val):
    """Collapse any 'Other...' entry to 'Other'; leave other answers unchanged."""
    if pd.isna(val):
        return val
    val = str(val).strip()
    return "Other" if val.startswith("Other") else val


def clean_survey(df):
    df[ROLE_COL] = clean_and_standardize(df[ROLE_COL], ROLE_MAP)
    df[INDUSTRY_COL] = clean_and_standardize(df[INDUSTRY_COL], INDUSTRY_MAP)
    df[COUNTRY_COL] = clean_and_standardize(df[COUNTRY_COL], COUNTRY_MAP)
    df[LANGUAGE_COL] = df[LANGUAGE_COL].apply(clean_language)
    df[PRIORITY_COL] = df[PRIORITY_COL].apply(collapse_other)
    df[ETHNICITY_COL] = df[ETHNICITY_COL].apply(collapse_other)
    return df.drop(columns=COLUMNS_TO_DROP)


def main():
    df = clean_survey(load_survey())
    df.to_excel(OUTPUT_FILE, index=False)


if __name__ == "__main__":
    main()