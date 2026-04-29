import html
import pandas as pd
import re
import argparse

def remove_html_tags(text):
    """Decode HTML entities and remove HTML tags."""
    if not isinstance(text, str):
        return ""
    # Decode entities like &bull; &nbsp; &zwnj; &amp; etc.
    text = html.unescape(text)
    # Remove any remaining named/numeric entities that unescape didn't handle
    text = re.sub(r"&[a-zA-Z]+;|&#?\w+;", "", text)
    # Remove HTML tags
    text = re.sub(r"<[^>]+>", "", text)
    # Remove zero-width and invisible Unicode characters (zwnj, zwj, bom, etc.)
    text = re.sub(r"[\u200b\u200c\u200d\u200e\u200f\ufeff\u00ad]", "", text)
    # Remove bullet characters that may have been decoded from &bull;
    text = re.sub(r"[•·\-]{2,}\s*", "", text)  # repeated dashes/bullets used as separators
    text = re.sub(r"^[•·]\s*", "", text)  # leading bullet
    return text


def remove_option_prefixes(text):
    """Remove option prefixes like a) b) or ক) খ) গ) ঘ) ঙ) from the start of text."""
    if not isinstance(text, str):
        return ""
    # Remove latin prefixes: a) A) a. A. (a) etc.
    text = re.sub(r"^\s*[a-eA-E][).]\s*", "", text)
    # Remove Bengali prefixes: ক) খ) গ) ঘ) ঙ) and variants
    bengali_options = "কখগঘঙ"
    text = re.sub(rf"^\s*[{bengali_options}][).।]\s*", "", text)
    # Remove numeric prefixes: 1) 2) 1. 2. etc.
    text = re.sub(r"^\s*\d+[).]\s*", "", text)
    return text


def clean_text(text):
    """Remove HTML tags, option prefixes, and normalize whitespace."""
    text = remove_html_tags(text)
    text = remove_option_prefixes(text)
    # Collapse multiple spaces/newlines into a single space
    text = re.sub(r"\s+", " ", text)
    # Add a space after dari (।) if not already followed by a space or end of string
    text = re.sub(r"।(?!\s|$)", "। ", text)
    return text.strip()


def resolve_answer(row):
    """If answer is a number (1-5), return the option number in brackets followed by the option text."""
    raw = str(row.get("answer", "")).strip()
    if raw in {"1", "2", "3", "4", "5"}:
        option_text = clean_text(row.get(f"option{raw}", ""))
        #return f"({raw}) {option_text}" if option_text else raw
        return option_text if option_text else raw
    return clean_text(raw)


def format_record(row):
    """Format a single row into the required text block."""
    lines = []
    lines.append(f"question: {clean_text(row.get('question', ''))}")
    lines.append("")
    for i in range(1, 6):
        col = f"option{i}"
        value = clean_text(row.get(col, "")) if col in row else ""
        lines.append(f"option{i}: {value}")
        lines.append("")
    lines.append(f"answer: {resolve_answer(row)}")
    lines.append("")
    lines.append(f"explain: {clean_text(row.get('explain', ''))}")
    lines.append("")
    slug_val = str(row.get("slug", "")).strip() if "slug" in row else ""
    if slug_val:
        lines.append(f"slug: {slug_val}")
    return "\n".join(lines)


def convert_parquet_to_txt(input_parquet: str, output_txt: str):
    df = pd.read_parquet(input_parquet)

    # Ensure expected columns exist
    required = {"question", "answer", "explain"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"Missing columns in parquet file: {missing}")

    # Fill missing option columns with empty strings
    for i in range(1, 6):
        col = f"option{i}"
        if col not in df.columns:
            df[col] = ""

    records = []
    for _, row in df.iterrows():
        records.append(format_record(row))

    separator = "\n\n" +"\n"
    output = separator.join(records)

    with open(output_txt, "w", encoding="utf-8") as f:
        f.write(output)

    print(f"Saved {len(df)} records to {output_txt}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Convert a parquet QA dataset to a formatted txt file."
    )
    parser.add_argument("input", help="Path to the input .parquet file")
    parser.add_argument("output", help="Path to the output .txt file")
    args = parser.parse_args()

    convert_parquet_to_txt(args.input, args.output)

