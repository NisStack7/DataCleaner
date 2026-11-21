import logging
import os
import re
import tempfile
import time
from decimal import Decimal, InvalidOperation
from pathlib import Path

import pandas as pd
from watchdog.events import FileSystemEventHandler
from watchdog.observers import Observer

BASE = Path(__file__).resolve().parent
INPUT = BASE / "input"
OUTPUT = BASE / "output"

INPUT.mkdir(parents=True, exist_ok=True)
OUTPUT.mkdir(parents=True, exist_ok=True)

SUPPORTED = {".csv", ".xlsx"}
NUMBER_FORMAT = "dot"

MISSING = {
    "", "na", "n/a", "null", "none", "nan", "-",
    "not available"
}

STABILITY_SECONDS = 2
MAX_RETRIES = 3

# Protect CSV files against common spreadsheet formula injection.
SAFE_CSV_FOR_SPREADSHEETS = True
ALLOW_ZERO_UNIT_PRICE = True
ALLOW_NEGATIVE_SALARY = False
AMOUNT_TOLERANCE =Decimal("0.01")
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    datefmt="%H:%M:%S",
)


# ============================================================
# 2. COLUMN ALIASES
# ============================================================

ALIASES = {
    "Customer_ID": {
        "customer_id", "customerid", "customer id", "id"
    },
    "Full_Name": {
        "full_name", "name", "customer_name", "customer name"
    },
    "Email": {
        "email", "email_address", "e-mail"
    },
    "Phone": {
        "phone", "phone_number", "mobile",
        "mobile_number", "contact"
    },
    "Address": {
        "address", "street_address"
    },
    "State": {
        "state", "province"
    },
    "Postal_Code": {
        "postal_code", "postcode", "postal code",
        "zip", "zip_code"
    },
    "Join_Date": {
        "join_date", "joined", "registration_date",
        "date_joined"
    },
    "City": {
        "city", "town"
    },
    "Product": {
        "product", "item", "item_name"
    },
    "Quantity": {
        "quantity", "qty"
    },
    "Unit_Price": {
        "unit_price",
        "unit price",
        "price",
        "unit_cost",
        "unit cost",
        "unit_price_inr",
        "unit_price_usd",
        "unit_price_eur",
        "unit_price_gbp",
    },
    "Amount": {
        "amount",
        "total_amount",
        "order_amount",
    },
    "Salary": {
        "salary",
        "annual_salary",
        "monthly_salary",
        "income",
    },
    "Order_Status": {
        "order_status", "status"
    },
    "Order_Date": {
        "order_date", "date", "purchase_date"
    },
}


# ============================================================
# 3. BASIC CLEANING HELPERS
# ============================================================

def missing(value):
    """Return True for empty strings and common missing-value markers."""
    try:
        if value is None or pd.isna(value):
            return True
    except (TypeError, ValueError):
        pass

    return str(value).strip().casefold() in MISSING


def text(value):
    """Convert a value to trimmed text without erasing valid text tokens."""
    try:
        if value is None or pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass

    return str(value).strip()


def spaces(value):
    """Trim text and replace repeated whitespace with one space."""
    return re.sub(r"\s+", " ", text(value)).strip()


def clean_date(value):
    """Normalize unambiguous dates; preserve ambiguous dates for review."""
    if missing(value):
        return ""

    if isinstance(value, pd.Timestamp):
        return value.strftime("%Y-%m-%d")

    raw = str(value).strip()

    # Formats whose year-month-day order is explicit.
    explicit_formats = (
        "%Y-%m-%d",
        "%Y/%m/%d",
        "%Y-%m-%d %H:%M:%S",
    )

    for fmt in explicit_formats:
        try:
            return pd.to_datetime(
                raw, format=fmt, errors="raise"
            ).strftime("%Y-%m-%d")
        except (ValueError, TypeError):
            continue

    # Month names remove most day/month ambiguity.
    named_formats = (
        "%d %b %Y",
        "%d %B %Y",
        "%b %d %Y",
        "%B %d %Y",
    )

    for fmt in named_formats:
        try:
            return pd.to_datetime(
                raw, format=fmt, errors="raise"
            ).strftime("%Y-%m-%d")
        except (ValueError, TypeError):
            continue

    # For numeric dates, accept only when day/month order
    # cannot change the result.
    numeric_formats = (
        "%d/%m/%Y",
        "%m/%d/%Y",
        "%d-%m-%Y",
        "%m-%d-%Y",
        "%d.%m.%Y",
        "%m.%d.%Y",
    )

    candidates = set()

    for fmt in numeric_formats:
        try:
            candidates.add(
                pd.to_datetime(
                    raw, format=fmt, errors="raise"
                ).strftime("%Y-%m-%d")
            )
        except (ValueError, TypeError):
            continue

    if len(candidates) == 1:
        return candidates.pop()

    # Empty result lets validation report an invalid or ambiguous date.
    return ""

def clean_phone(value):
    """Remove common phone-number separators without guessing a country code."""
    return re.sub(r"[\s().-]+", "", text(value))

def clean_number(value, integer=False, style=NUMBER_FORMAT):
    """
    Normalize numbers using an explicitly selected numeric convention.
    dot:      1,234.56
    european: 1.234,56
    Invalid or ambiguous text is preserved for quality reporting.
    """
    value = text(value).strip()

    if not value:
        return ""

    if style not in {"dot", "european"}:
        raise ValueError(
            "NUMBER_FORMAT must be 'dot' or 'european'."
        )

    # Allow spaces as thousands separators, e.g. 1 234.56.
    value = re.sub(r"(?<=\d)\s+(?=\d)", "", value)

    if style == "dot":
        # Valid examples: 1234.56, 1,234.56, -25, 1,234
        pattern = r"[+-]?(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d+)?"

        if not re.fullmatch(pattern, value):
            return value

        value = value.replace(",", "")

    else:
        # Valid examples: 1234,56, 1.234,56, -25, 1.234
        pattern = r"[+-]?(?:\d+|\d{1,3}(?:\.\d{3})+)(?:,\d+)?"

        if not re.fullmatch(pattern, value):
            return value

        value = value.replace(".", "").replace(",", ".")

    try:
        number = Decimal(value)
    except InvalidOperation:
        return text(value)

    if not number.is_finite():
        return text(value)

    if integer:
        if number == number.to_integral_value():
            return str(int(number))
        return format(number.normalize(), "f")

    return format(number.normalize(), "f")
# 4. READ CSV AND EXCEL FILES
# ============================================================

def read_data(path):
    path = Path(path)

    if path.suffix.lower() == ".csv":
        try:
            return pd.read_csv(
                path,
                dtype=str,
                keep_default_na=False,
                encoding="utf-8-sig",
            )
        except UnicodeDecodeError:
            return pd.read_csv(
                path,
                dtype=str,
                keep_default_na=False,
                encoding="cp1252",
            )

    if path.suffix.lower() == ".xlsx":
        return pd.read_excel(
            path,
            dtype=str,
            keep_default_na=False,
            engine="openpyxl",
        )

    raise ValueError("Only CSV and XLSX files are supported.")


# ============================================================
# 5. STANDARDIZE COLUMN HEADERS
# ============================================================

def canonicalize(df):
    """Standardize known headers and reject ambiguous column mappings."""
    df = df.copy()

    normalized = [
        re.sub(r"\s+", "_", str(col).strip())
        for col in df.columns
    ]

    folded = [col.casefold() for col in normalized]

    if any(not col for col in normalized):
        raise ValueError(
            "One or more column headers are blank. "
            "Give every column a name."
        )

    if any(
        col == "source_row" or col.startswith("original_")
        for col in folded
    ):
        raise ValueError(
            "Input headers may not use reserved names "
            "Source_Row or Original_*. Rename those columns."
        )

    if len(set(folded)) != len(folded):
        duplicates = sorted({
            normalized[i]
            for i, col in enumerate(folded)
            if folded.count(col) > 1
        })

        raise ValueError(
            f"Duplicate column headers after normalization: {duplicates}"
        )

    df.columns = normalized

    lookup = {}

    for canonical, aliases in ALIASES.items():
        lookup[canonical.casefold()] = canonical

        for alias in aliases:
            key = re.sub(r"\s+", "_", alias).casefold()
            lookup[key] = canonical

    targets = {}

    for col in df.columns:
        key = re.sub(
            r"\s+", "_", str(col).strip()
        ).casefold()

        target = lookup.get(key)

        if target:
            targets.setdefault(target, []).append(col)

    collisions = {
        target: cols
        for target, cols in targets.items()
        if len(cols) > 1
    }

    if collisions:
        details = "; ".join(
            f"{target} <- {cols}"
            for target, cols in collisions.items()
        )

        raise ValueError(
            "Multiple input columns map to the same field: "
            f"{details}. Rename one of the source columns."
        )

    rename = {}

    for col in df.columns:
        key = re.sub(
            r"\s+", "_", str(col).strip()
        ).casefold()

        if key in lookup:
            rename[col] = lookup[key]

    return df.rename(columns=rename)


# ============================================================
# 6. STANDARDIZE FIELD VALUES
# ============================================================

def standardize_fields(df):
    """Clean recognized columns while retaining original values separately."""

    if "Customer_ID" in df:
        df["Customer_ID"] = df["Customer_ID"].map(spaces)

    if "Full_Name" in df:
        df["Full_Name"] = df["Full_Name"].map(spaces)

    if "Email" in df:
        df["Email"] = df["Email"].map(
            lambda value: spaces(value).casefold()
        )

    if "Phone" in df:
        df["Phone"] = df["Phone"].map(clean_phone)

    if "Address" in df:
        df["Address"] = df["Address"].map(spaces)

    if "City" in df:
        df["City"] = df["City"].map(spaces)

    if "State" in df:
        df["State"] = df["State"].map(spaces)

    if "Postal_Code" in df:
        def clean_postal(value):
            value = text(value)

            if re.fullmatch(r"\d+\.0+", value):
                return re.sub(r"\.0+$", "", value)

            return value

        df["Postal_Code"] = df["Postal_Code"].map(clean_postal)

    for col in ("Join_Date", "Order_Date"):
        if col in df:
            df[col] = df[col].map(clean_date)

    if "Quantity" in df:
        df["Quantity"] = df["Quantity"].map(
            lambda value: clean_number(value, integer=True)
        )

    for col in ("Unit_Price", "Amount", "Salary"):
        if col in df:
            df[col] = df[col].map(clean_number)

    if "Order_Status" in df:
        status_aliases = {
            "new": "New",
            "contacted": "Contacted",
            "pending": "Pending",
            "completed": "Completed",
            "cancelled": "Cancelled",
        }

        if "Order_Status" in df:
            status_aliases = {
                "new": "New",
                "contacted": "Contacted",
                "pending": "Pending",
                "completed": "Completed",
                "cancelled": "Cancelled",
        }

        df["Order_Status"] = df["Order_Status"].map(
            lambda value: status_aliases.get(
                spaces(value).casefold(),
                spaces(value),
            )
        )

    # Recalculate Amount = Quantity x Unit_Price
    if {"Quantity", "Unit_Price", "Amount"}.issubset(df.columns):

        def calculate_amount(row):
            qty = text(row["Quantity"])
            price = text(row["Unit_Price"])

            try:
                quantity = Decimal(qty)
                unit_price = Decimal(price)

                if (
                    not quantity.is_finite()
                    or not unit_price.is_finite()
                    or quantity <= 0
                    or quantity != quantity.to_integral_value()
                    or unit_price < 0
                ):
                    return row["Amount"]

                return format(quantity * unit_price, ".2f")

            except (InvalidOperation, ValueError):
                return row["Amount"]

        df["Amount"] = df.apply(calculate_amount, axis=1)

    return df


# ============================================================
# 7. DUPLICATE AND CONFLICT REVIEW
# ============================================================

def find_duplicates(df, compare_cols):
    """
    Remove only exact duplicates.

    Repeated IDs, emails, names, and name-plus-phone combinations
    are reported for review but do not cause automatic deletion.
    """
    duplicate_rows = []
    exact_duplicate_indices = []
    seen = {}

    for idx, row in df.iterrows():
        signature = tuple(
            text(row.get(col, ""))
            for col in compare_cols
        )
        if not any(signature):
            continue

        if signature in seen:
            exact_duplicate_indices.append(idx)

            duplicate_rows.append({
                "Issue": "Exact duplicate after cleaning",
                "Customer_ID": row.get("Customer_ID", ""),
                "Source_Row": row["Source_Row"],
                "Related_Source_Row": seen[signature],
                "Details": (
                    "Exact duplicate excluded from Cleaned_Data. "
                    "Original input file remains unchanged."
                ),
            })
        else:
            seen[signature] = row["Source_Row"]

    # Repeated customer IDs and email addresses.
    # Exclude exact duplicate rows already marked for removal.
    review_df = df.drop(index=exact_duplicate_indices)

    for key_col in ("Customer_ID", "Email"):
        if key_col not in review_df.columns:
            continue

        group_df = review_df[
            review_df[key_col].map(
                lambda value: bool(text(value))
            )
        ]

        for key, group in group_df.groupby(key_col, sort=False):
            if len(group) <= 1:
                continue

            source_rows = (
                group["Source_Row"].astype(str).tolist()
            )

            duplicate_rows.append({
                "Issue": f"Repeated {key_col}",
                "Customer_ID": group.iloc[-1].get(
                    "Customer_ID", ""
                ),
                "Source_Row": source_rows[-1],
                "Related_Source_Row": ", ".join(
                    source_rows[:-1]
                ),
                "Details": (
                    f"Repeated value {key!r}; "
                    "review before merging or deleting records."
                ),
            })
            # Repeated phone numbers: flag for manual review.
    # Do not automatically delete records based on phone alone.
    if "Phone" in review_df.columns:
        phone_df = review_df[
            review_df["Phone"].map(
                lambda value: bool(text(value))
            )
        ]

        for phone, group in phone_df.groupby(
            "Phone", sort=False
        ):
            if len(group) <= 1:
                continue

            source_rows = (
                group["Source_Row"].astype(str).tolist()
            )

            duplicate_rows.append({
                "Issue": "Repeated Phone",
                "Customer_ID": group.iloc[-1].get(
                    "Customer_ID", ""
                ),
                "Source_Row": source_rows[-1],
                "Related_Source_Row": ", ".join(
                    source_rows[:-1]
                ),
                "Details": (
                    f"Phone number {phone!r} appears in multiple "
                    "records. Verify whether these are the same "
                    "person or different people. No records deleted."
                ),
            })

    # Repeated name and phone.
    if {"Full_Name", "Phone"}.issubset(df.columns):
        subset = review_df[
            (review_df["Full_Name"] != "") &
            (review_df["Phone"] != "")
        ]

        for _, group in subset.groupby(
            ["Full_Name", "Phone"], sort=False
        ):
            if len(group) <= 1:
                continue

            source_rows = group["Source_Row"].astype(str).tolist()

            duplicate_rows.append({
                "Issue": "Repeated name + phone",
                "Customer_ID": group.iloc[-1].get("Customer_ID", ""),
                "Source_Row": source_rows[-1],
                "Related_Source_Row": ", ".join(source_rows[:-1]),
                "Details": (
                    "Possible same person or repeated record; "
                    "verify manually."
                ),
            })

    # Same name but differing IDs, emails, phones, or other details.
    if "Full_Name" in df.columns:
        subset = review_df[review_df["Full_Name"] != ""]

        conflict_cols = [
            col for col in compare_cols
            if col not in {"Full_Name"}
        ]

        for name, group in subset.groupby("Full_Name", sort=False):
            if len(group) <= 1:
                continue

            differing_fields = []

            for col in conflict_cols:
                unique_values = {
                    text(value)
                    for value in group[col].tolist()
                }

                if len(unique_values) > 1:
                    differing_fields.append(col)

            if not differing_fields:
                continue

            source_rows = group["Source_Row"].astype(str).tolist()

            duplicate_rows.append({
                "Issue": "Same name with differing details",
                "Customer_ID": group.iloc[-1].get("Customer_ID", ""),
                "Source_Row": source_rows[-1],
                "Related_Source_Row": ", ".join(source_rows[:-1]),
                "Details": (
                    f"Name {name!r} has differing values in: "
                    f"{', '.join(differing_fields)}. "
                    "Possible different people or separate transactions; "
                    "review manually. No records deleted for this finding."
                ),
            })

    return duplicate_rows, exact_duplicate_indices

# ============================================================
# 8. DATA QUALITY VALIDATION
# ============================================================

def validate_dataframe(df):
    quality = []

    check_fields = [
        "Customer_ID", "Full_Name", "Email", "Phone",
        "Address", "State", "Postal_Code", "Join_Date",
        "City", "Product", "Quantity", "Unit_Price",
        "Amount", "Order_Status", "Order_Date", "Salary",
    ]

    def add_issue(row, field, issue, original, action):
        quality.append({
            "SOURCE_ROW": row["Source_Row"],
            "CUSTOMER_ID": row.get("Customer_ID", ""),
            "FIELD_TO_FIX": field,
            "PROBLEM_FOUND": issue.replace("_", " ").title(),
            "ORIGINAL_VALUE": original,
            "WHERE_TO_CHECK": (
                f"Open the original file, go to row "
                f"{row['Source_Row']}, and find the "
                f"'{field}' column."
            ),
            "ACTION_REQUIRED": action,
        })

    for _, row in df.iterrows():

        # 1. Missing values and invalid dates
        for col in check_fields:
            if col not in df.columns:
                continue

            value = text(row.get(col, ""))
            original = text(
                row.get("Original_" + col, row.get(col, ""))
            )

            if missing(value):
                if col in ("Join_Date", "Order_Date") and original:
                    issue = "INVALID DATE - PLEASE VERIFY"
                else:
                    issue = "MISSING VALUE"

                add_issue(
                    row,
                    col,
                    issue,
                    original,
                    "Check the original file and enter the correct value.",
                )


        # 2. Email validation
        email = text(row.get("Email", "")).strip().lower()
        original_email = text(
            row.get("Original_Email", row.get("Email", ""))
        ).strip()

        # Validate common and less-common email providers.
        email_pattern = (
            r"[a-z0-9!#$%&'*+/=?^_`{|}~-]+"
            r"(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*"
            r"@[a-z0-9](?:[a-z0-9-]*[a-z0-9])?"
            r"(?:\.[a-z0-9](?:[a-z0-9-]*[a-z0-9])?)+"
        )

        if email:
            if not re.fullmatch(email_pattern, email):
                add_issue(
                    row,
                    "Email",
                    "INVALID EMAIL FORMAT",
                    original_email,
                    "Verify the email address in the original file.",
                )

        # Save lowercase and trimmed email.
        if email != text(row.get("Email", "")).strip():
            df.at[_, "Email"] = email

        # 3. Country-neutral phone format validation
        phone = text(row.get("Phone", ""))

        if phone:
            digits = phone[1:] if phone.startswith("+") else phone

            invalid_phone = (
                not digits.isdigit()
                or not 7 <= len(digits) <= 15
                or phone.count("+") > 1
                or ("+" in phone and not phone.startswith("+"))
            )

            if invalid_phone:
                add_issue(
                    row,
                    "Phone",
                    "INVALID PHONE FORMAT",
                    text(row.get("Original_Phone", phone)),
                    (
                        "Check the number, including its country code. "
                        "This basic check does not confirm country-specific validity."
                    ),
                )

        # 4. Country-neutral postal-code sanity check
        postal = text(row.get("Postal_Code", ""))

        if postal:
            invalid_postal = not re.fullmatch(
                r"[A-Za-z0-9](?:[A-Za-z0-9 -]{0,18}[A-Za-z0-9])?",
                postal,
            )

            if invalid_postal:
                add_issue(
                    row,
                    "Postal_Code",
                    "INVALID POSTAL CODE FORMAT",
                    text(
                        row.get("Original_Postal_Code", postal)
                    ),
                    (
                        "Check the postal code against the destination "
                        "country's official postal-code rules."
                    ),
                )
        # 5. Quantity validation
        qty = text(row.get("Quantity", ""))

        if qty:
            if not re.fullmatch(r"-?\d+", qty):
                issue = "QUANTITY IS NOT A WHOLE NUMBER"
            elif int(qty) <= 0:
                issue = "QUANTITY IS ZERO OR NEGATIVE"
            else:
                issue = ""

            if issue:
                add_issue(
                    row,
                    "Quantity",
                    issue,
                    text(row.get("Original_Quantity", qty)),
                    "Verify the quantity against the original source.",
                )

        # 6. Unit price validation
        price = text(row.get("Unit_Price", ""))

        if price:
            try:
                number = Decimal(price)

                if not number.is_finite():
                    issue = "PRICE IS NOT FINITE"
                elif number < 0:
                    issue = "PRICE IS NEGATIVE"
                else:
                    issue = ""

            except (InvalidOperation, ValueError):
                issue = "PRICE IS NOT NUMERIC"

            if issue:
                add_issue(
                    row,
                    "Unit_Price",
                    issue,
                    text(row.get("Original_Unit_Price", price)),
                    "Verify the price against the original source.",
                )

        # 7. Amount validation
        amount = text(row.get("Amount", ""))

        if amount:
            try:
                number = Decimal(amount)

                if not number.is_finite():
                    issue = "AMOUNT IS NOT FINITE"
                elif number < 0:
                    issue = "AMOUNT IS NEGATIVE"
                else:
                    issue = ""

            except (InvalidOperation, ValueError):
                issue = "AMOUNT IS NOT NUMERIC"

            if issue:
                add_issue(
                    row,
                    "Amount",
                    issue,
                    text(row.get("Original_Amount", amount)),
                    "Verify the amount against the original source.",
                )
        # 7B. Check whether Amount matches Quantity x Unit_Price
        qty = text(row.get("Quantity", ""))
        price = text(row.get("Unit_Price", ""))
        amount = text(row.get("Amount", ""))

        if qty and price and amount:
            try:
                qty_number = Decimal(qty)
                price_number = Decimal(price)
                amount_number = Decimal(amount)

                values_are_valid = (
                    qty_number.is_finite()
                    and price_number.is_finite()
                    and amount_number.is_finite()
                    and qty_number > 0
                    and qty_number == qty_number.to_integral_value()
                    and price_number >= 0
                    and amount_number >= 0
                )

                if values_are_valid:
                    expected_amount = qty_number * price_number

                    if abs(amount_number - expected_amount) > AMOUNT_TOLERANCE:
                        add_issue(
                            row,
                            "Amount",
                            "AMOUNT DOES NOT MATCH QUANTITY X UNIT PRICE",
                            text(row.get("Original_Amount", amount)),
                            (
                                f"Expected amount: {expected_amount:.2f}. "
                                f"Recorded amount: {amount_number:.2f}. "
                                "Check the quantity, unit price, and amount."
                            ),
                        )

            except (InvalidOperation, ValueError):
                # Existing validation handles non-numeric values.
                pass

        # 8. Salary validation
        salary = text(row.get("Salary", ""))

        if salary:
            try:
                number = Decimal(salary)

                if not number.is_finite():
                    issue = "SALARY IS NOT FINITE"
                elif number < 0:
                    issue = "SALARY IS NEGATIVE"
                else:
                    issue = ""

            except (InvalidOperation, ValueError):
                issue = "SALARY IS NOT NUMERIC"

            if issue:
                add_issue(
                    row,
                    "Salary",
                    issue,
                    text(row.get("Original_Salary", salary)),
                    "Verify the salary against the original source.",
                )

    return pd.DataFrame(
        quality,
        columns=[
            "SOURCE_ROW",
            "CUSTOMER_ID",
            "FIELD_TO_FIX",
            "PROBLEM_FOUND",
            "ORIGINAL_VALUE",
            "WHERE_TO_CHECK",
            "ACTION_REQUIRED",
        ],
    ).rename(columns={
        "SOURCE_ROW": "ROW NUMBER",
        "CUSTOMER_ID": "CUSTOMER ID",
        "FIELD_TO_FIX": "FIELD TO CHECK",
        "PROBLEM_FOUND": "WHAT IS WRONG?",
        "ORIGINAL_VALUE": "ORIGINAL VALUE",
        "WHERE_TO_CHECK": "WHERE TO FIND IT",
        "ACTION_REQUIRED": "HOW TO FIX IT",
    })

# ============================================================
# 9. MAIN CLEANING PIPELINE
# ============================================================

def clean_dataframe(raw):
    if len(raw.columns) == 0:
        raise ValueError("No column headers found.")

    df = canonicalize(raw)

    # For ordinary CSV/XLSX files, the first data row is Excel row 2.
    df.insert(0, "Source_Row", range(2, len(df) + 2))

    fields = [col for col in ALIASES if col in df.columns]

    # Preserve original values before normalization.
    for col in fields:
        df["Original_" + col] = df[col].map(
            lambda value: (
                "" if value is None or pd.isna(value)
                else str(value)
            )
        )

    # Clean whitespace in data columns, but not the audit columns.
    for col in df.columns:
        if col != "Source_Row" and not col.startswith("Original_"):
            df[col] = df[col].map(spaces)

    df = standardize_fields(df)

    compare_cols = [
        col for col in df.columns
        if col != "Source_Row" and not col.startswith("Original_")
    ]
    duplicate_rows, exact_duplicate_indices = find_duplicates(
        df, compare_cols
    )

    quality_df = validate_dataframe(df)

    duplicates_df = pd.DataFrame(
        duplicate_rows,
        columns=[
            "Issue", "Customer_ID", "Source_Row",
            "Related_Source_Row", "Details",
        ],
    )

    # Only exact duplicates are removed.
    # Final output contains clean data only.
    cleaned = df.drop(
        index=exact_duplicate_indices
    ).drop(
        columns=[
            col for col in df.columns
            if col.startswith("Original_")
        ],
        errors="ignore",
    ).copy()

    return cleaned, duplicates_df, quality_df, len(exact_duplicate_indices)


# ============================================================
# 10. CSV SPREADSHEET FORMULA PROTECTION
# ============================================================

def csv_safe_frame(frame):
    """Protect CSV exports from common formula-like text values."""
    if not SAFE_CSV_FOR_SPREADSHEETS:
        return frame

    out = frame.copy()

    dangerous = re.compile(r"^[\s\x00-\x1f]*[=+@]")
    formula_minus = re.compile(
        r"^[\s\x00-\x1f]*-(?!\d+(?:\.\d+)?$)"
    )

    def protect(value):
        if not isinstance(value, str):
            return value

        if dangerous.search(value) or formula_minus.search(value):
            return "'" + value

        return value

    for col in out.columns:
        out[col] = out[col].map(protect)

    return out
# ============================================================
# 11. SAFE CSV WRITING
# ============================================================

def atomic_to_csv(frame, destination):
    """Write a valid CSV safely, then replace the destination."""
    destination = Path(destination)

    fd, temp_name = tempfile.mkstemp(
        prefix=destination.stem + "_",
        suffix=".tmp",
        dir=destination.parent,
    )
    os.close(fd)

    try:
        csv_safe_frame(frame).to_csv(
            temp_name,
            index=False,
            encoding="utf-8-sig",
            lineterminator="\n",
        )

        os.replace(temp_name, destination)

    except Exception:
        try:
            os.unlink(temp_name)
        except OSError:
            pass
        raise

# ============================================================
# 12. EXCEL REPORT WRITING
# ============================================================

def atomic_to_excel(cleaned, duplicates, quality, destination):
    """Create a readable Excel report with clear headings and highlights."""

    from openpyxl.styles import (
        Font, PatternFill, Alignment, Border, Side
    )
    from openpyxl.utils import get_column_letter

    destination = Path(destination)

    fd, temp_name = tempfile.mkstemp(
        prefix=destination.stem + "_",
        suffix=".tmp.xlsx",
        dir=destination.parent,
    )
    os.close(fd)

    try:
        with pd.ExcelWriter(temp_name, engine="openpyxl") as writer:
            cleaned.to_excel(
                writer, index=False, sheet_name="Cleaned_Data"
            )
            duplicates.to_excel(
                writer, index=False, sheet_name="Duplicate_Review"
            )
            quality.to_excel(
                writer, index=False, sheet_name="Data_Quality_Report"
            )

            # ------------------------------------------------
            # 1. COLORS AND BORDERS
            # ------------------------------------------------
            blue_fill = PatternFill(
                fill_type="solid", fgColor="365F91"
            )
            title_fill = PatternFill(
                fill_type="solid", fgColor="203864"
            )
            yellow_fill = PatternFill(
                fill_type="solid", fgColor="FFF2CC"
            )

            red_fill = PatternFill(
                fill_type="solid", fgColor="FFC7CE"
            )

            red_font = Font(
                color="9C0006",
                bold=True,
            )
            yellow_font = Font(
                color="9C6500",
                bold=True,
            )

            alternate_fill = PatternFill(
                fill_type="solid", fgColor="F3F6FA"
            )

            thin_gray = Side(
                style="thin", color="D9E2F3"
            )
            cell_border = Border(
                bottom=thin_gray,
                left=thin_gray,
                right=thin_gray,
            )

            # ------------------------------------------------
            # 2. FORMAT ALL SHEETS
            # ------------------------------------------------
            for ws in writer.book.worksheets:
                ws.freeze_panes = "A2"

                if ws.max_row and ws.max_column:
                    ws.auto_filter.ref = ws.dimensions

                for cell in ws[1]:
                    cell.font = Font(
                        bold=True, color="FFFFFF", size=11
                    )
                    cell.fill = blue_fill
                    cell.alignment = Alignment(
                        horizontal="center",
                        vertical="center",
                        wrap_text=True,
                    )
                    cell.border = cell_border

                ws.row_dimensions[1].height = 32

                for row in ws.iter_rows(min_row=2):
                    for cell in row:
                        cell.alignment = Alignment(
                            vertical="top",
                            wrap_text=True,
                        )
                        cell.border = cell_border

                        # Prevent text from being interpreted
                        # as an Excel formula.
                        if isinstance(cell.value, str):
                            cell.data_type = "s"

                # Set sensible default widths.
                for col_cells in ws.columns:
                    letter = get_column_letter(
                        col_cells[0].column
                    )
                    width = max(
                        (
                            len(str(cell.value or ""))
                            for cell in list(col_cells)[:500]
                        ),
                        default=10,
                    )
                    ws.column_dimensions[letter].width = min(
                        max(width + 2, 12), 35
                    )

            # ------------------------------------------------
            # 3. HIGHLIGHT EMPTY CELLS IN CLEANED DATA
            # ------------------------------------------------
            cleaned_ws = writer.book["Cleaned_Data"]

            for row_num in range(2, cleaned_ws.max_row + 1):
                for col_num in range(
                    1, cleaned_ws.max_column + 1
                ):
                    cell = cleaned_ws.cell(
                        row_num, col_num
                    )

                    if missing(cell.value):
                        cell.fill = yellow_fill
                        cell.font = yellow_font

            # ------------------------------------------------
            # 4. MAKE THE QUALITY REPORT EASY TO READ
            # ------------------------------------------------
            quality_ws = writer.book["Data_Quality_Report"]

            # Give the report simple, reader-friendly headings.
            friendly_headers = {
                "SOURCE_ROW": "ROW NUMBER",
                "CUSTOMER_ID": "CUSTOMER ID",
                "FIELD_TO_FIX": "FIELD TO CHECK",
                "PROBLEM_FOUND": "WHAT IS WRONG?",
                "ORIGINAL_VALUE": "ORIGINAL VALUE",
                "WHERE_TO_CHECK": "WHERE TO FIND IT",
                "ACTION_REQUIRED": "HOW TO FIX IT",
            }

            # Add a title row above the column headings.
            quality_ws.insert_rows(1)
            quality_ws.merge_cells(
                start_row=1,
                start_column=1,
                end_row=1,
                end_column=max(quality_ws.max_column, 1),
            )

            title_cell = quality_ws.cell(1, 1)
            title_cell.value = "DATA QUALITY REPORT"
            title_cell.font = Font(
                bold=True, color="FFFFFF", size=16
            )
            title_cell.fill = title_fill
            title_cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
            )
            quality_ws.row_dimensions[1].height = 36

            # Row 2 now contains the original column headings.
            headers = {
                str(cell.value).strip().upper(): cell.column
                for cell in quality_ws[2]
                if cell.value is not None
            }

            for old_name, new_name in friendly_headers.items():
                col_num = headers.get(old_name)

                if col_num is not None:
                    quality_ws.cell(
                        row=2, column=col_num
                    ).value = new_name

            # Style the friendly column headings.
            for cell in quality_ws[2]:
                cell.font = Font(
                    bold=True, color="FFFFFF", size=11
                )
                cell.fill = blue_fill
                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center",
                    wrap_text=True,
                )
                cell.border = cell_border

            quality_ws.row_dimensions[2].height = 34
            quality_ws.freeze_panes = "A3"

            if quality_ws.max_row >= 2:
                quality_ws.auto_filter.ref = (
                    f"A2:{get_column_letter(quality_ws.max_column)}"
                    f"{quality_ws.max_row}"
                )

            # Give long instructions enough room.
            friendly_widths = {
                "ROW NUMBER": 13,
                "CUSTOMER ID": 16,
                "FIELD TO CHECK": 18,
                "WHAT IS WRONG?": 28,
                "ORIGINAL VALUE": 26,
                "WHERE TO FIND IT": 38,
                "HOW TO FIX IT": 42,
            }

            for cell in quality_ws[2]:
                heading = str(cell.value or "")
                quality_ws.column_dimensions[
                    get_column_letter(cell.column)
                ].width = friendly_widths.get(heading, 20)

            # ------------------------------------------------
            # ------------------------------------------------
            # 5. HIGHLIGHT ALL MANUAL-REVIEW FINDINGS IN RED
            # ------------------------------------------------

            problem_col = None
            value_col = None
            row_number_col = None

            for cell in quality_ws[2]:
                heading = str(cell.value or "").strip().upper()

                if heading == "WHAT IS WRONG?":
                    problem_col = cell.column
                elif heading == "ORIGINAL VALUE":
                    value_col = cell.column
                elif heading == "ROW NUMBER":
                    row_number_col = cell.column

            # Row 1 = title, row 2 = headings, row 3 onward = findings.
            for row_num in range(3, quality_ws.max_row + 1):

                problem = ""

                if problem_col is not None:
                    problem = str(
                        quality_ws.cell(
                            row=row_num,
                            column=problem_col,
                        ).value or ""
                    ).strip().upper()

                # Missing values are yellow; invalid values remain red.
                if "MISSING VALUE" in problem:
                    fill = yellow_fill
                    font = yellow_font
                else:
                    fill = red_fill
                    font = red_font

                # Highlight the problem description.
                if problem_col is not None:
                    problem_cell = quality_ws.cell(
                        row=row_num,
                        column=problem_col,
                    )
                    problem_cell.fill = fill
                    problem_cell.font = font

                # Highlight the original value that needs checking.
                if value_col is not None:
                    value_cell = quality_ws.cell(
                        row=row_num,
                        column=value_col,
                    )
                    value_cell.fill = fill
                    value_cell.font = font

                # Highlight the row number so the finding is easy to locate.
                if row_number_col is not None:
                    row_number_cell = quality_ws.cell(
                        row=row_num,
                        column=row_number_col,
                    )
                    row_number_cell.fill = fill
                    row_number_cell.font = font
    
            # ------------------------------------------------
            # 6. HIGHLIGHT INVALID VALUES IN CLEANED DATA
            # ------------------------------------------------
            cleaned_ws = writer.book["Cleaned_Data"]
            quality_ws = writer.book["Data_Quality_Report"]

            # Find the internal source-row tracking column.
            source_col = None

            for cell in cleaned_ws[1]:
                if str(cell.value or "").strip() == "Source_Row":
                    source_col = cell.column
                    break

            # Find the quality-report columns by their headings.
            report_cols = {
                str(cell.value or "").strip().upper(): cell.column
                for cell in quality_ws[2]
                if cell.value is not None
            }

            field_col = report_cols.get("FIELD TO CHECK")
            row_col = report_cols.get("ROW NUMBER")

            # Map cleaned-data column names to Excel column numbers.
            data_cols = {
                str(cell.value).strip(): cell.column
                for cell in cleaned_ws[1]
                if cell.value is not None
            }

            # Match each quality finding to its cleaned-data cell.
            if source_col and field_col and row_col:
                source_to_excel_row = {}

                for excel_row in range(2, cleaned_ws.max_row + 1):
                    source_value = cleaned_ws.cell(
                        row=excel_row,
                        column=source_col,
                    ).value

                    if source_value is not None:
                        source_to_excel_row[str(source_value)] = excel_row

                for report_row in range(3, quality_ws.max_row + 1):
                    source_value = quality_ws.cell(
                        row=report_row,
                        column=row_col,
                    ).value

                    field_value = quality_ws.cell(
                        row=report_row,
                        column=field_col,
                    ).value

                    if source_value is None or field_value is None:
                        continue

                    target_row = source_to_excel_row.get(
                        str(source_value)
                    )
                    target_col = data_cols.get(
                        str(field_value).strip()
                    )

                    if target_row is None or target_col is None:
                        continue

                    target_cell = cleaned_ws.cell(
                        row=target_row,
                        column=target_col,
                    )

                    problem_value = str(
                        quality_ws.cell(
                            row=report_row,
                            column=report_cols["WHAT IS WRONG?"],
                        ).value or ""
                    ).strip().upper()

                    # Empty cells stay yellow; invalid values become red.
                    if (
                        missing(target_cell.value)
                        or "MISSING VALUE" in problem_value
                    ):
                        target_cell.fill = yellow_fill
                        target_cell.font = yellow_font
                    else:
                        target_cell.fill = red_fill
                        target_cell.font = red_font


            # Hide the internal tracking column in Excel.
            if source_col:
                cleaned_ws.column_dimensions[
                    get_column_letter(source_col)
                ].hidden = True

            # Hide gridlines on every sheet.
            for ws in writer.book.worksheets:
                ws.sheet_view.showGridLines = False

        os.replace(temp_name, destination)

    except Exception:
        try:
            os.unlink(temp_name)
        except OSError:
            pass
        raise

# 13. WRITE ALL OUTPUTS

def write_outputs(
    path, cleaned, duplicates, quality,
    input_count, exact_duplicate_count,
):
    stem = f"{path.stem}_{path.suffix.lower().lstrip('.')}"

    csv_path = OUTPUT / f"cleaned_{stem}.csv"
    xlsx_path = OUTPUT / f"cleaned_{stem}.xlsx"
    dup_path = OUTPUT / f"duplicate_review_{stem}.csv"
    quality_path = OUTPUT / f"data_quality_report_{stem}.csv"

    atomic_to_csv(
        cleaned.drop(columns=["Source_Row"], errors="ignore"),
        csv_path,
    )
    atomic_to_csv(duplicates, dup_path)
    atomic_to_csv(quality, quality_path)

    atomic_to_excel(
        cleaned,
        duplicates,
        quality,
        xlsx_path,
    )

    logging.info(
        "Finished %s | input rows=%d | output rows=%d | "
        "exact duplicates removed=%d",
        path.name,
        input_count,
        len(cleaned),
        exact_duplicate_count,
    )

    logging.info(
        "Duplicate findings=%d | quality findings=%d",
        len(duplicates),
        len(quality),
    )

    logging.info(
        "Created: %s | %s | %s | %s",
        csv_path,
        xlsx_path,
        dup_path,
        quality_path,
    )


# ============================================================
# 14. PROCESS ONE INPUT FILE
# ============================================================

def process_file(path):
    path = Path(path)

    if (
        not path.is_file()
        or path.suffix.lower() not in SUPPORTED
        or path.name.startswith("~$")
    ):
        return

    if OUTPUT in path.parents:
        return

    logging.info("Processing %s", path.name)

    raw = read_data(path)

    cleaned, duplicates, quality, exact_duplicate_count = (
        clean_dataframe(raw)
    )

    write_outputs(
        path,
        cleaned,
        duplicates,
        quality,
        len(raw),
        exact_duplicate_count,
    )

# 15. WATCH INPUT FOLDER

class Handler(FileSystemEventHandler):
    def __init__(self):
        super().__init__()
        self.events = set()

    def on_created(self, event):
        if not event.is_directory:
            self.events.add(event.src_path)

    def on_modified(self, event):
        if not event.is_directory:
            self.events.add(event.src_path)

    def on_moved(self, event):
        if not event.is_directory:
            self.events.add(event.dest_path)


# ============================================================
# 16. MAIN LOOP
# ============================================================

def main():
    handler = Handler()
    observer = Observer()

    observer.schedule(
        handler,
        str(INPUT),
        recursive=False,
    )

    observer.start()

    waiting = {}
    processed = {}

    # Process files already present when the script starts.
    for path in INPUT.iterdir():
        if path.is_file():
            handler.events.add(str(path))

    logging.info(
        "Advanced Data Cleaner started. Input=%s Output=%s",
        INPUT,
        OUTPUT,
    )
    logging.info(
        "Drop CSV/XLSX files into input/. Press Ctrl+C to stop."
    )

    try:
        while True:
            now = time.monotonic()

            # Register new or modified file events.
            for path in list(handler.events):
                handler.events.discard(path)

                waiting[path] = {
                    "signature": None,
                    "since": now,
                    "retries": 0,
                }

            # Wait for each file to stop changing before processing.
            for path, item in list(waiting.items()):
                file_path = Path(path)

                if (
                    not file_path.exists()
                    or not file_path.is_file()
                    or file_path.suffix.lower() not in SUPPORTED
                    or file_path.name.startswith("~$")
                    or OUTPUT in file_path.parents
                ):
                    waiting.pop(path, None)
                    continue

                try:
                    stat = file_path.stat()
                    signature = (
                        stat.st_size,
                        stat.st_mtime_ns,
                    )
                except OSError:
                    continue

                if signature != item["signature"]:
                    item["signature"] = signature
                    item["since"] = now
                    continue

                if now - item["since"] < STABILITY_SECONDS:
                    continue

                if processed.get(path) == signature:
                    waiting.pop(path, None)
                    continue

                try:
                    process_file(file_path)
                    processed[path] = signature
                    waiting.pop(path, None)

                except Exception as exc:
                    item["retries"] += 1

                    logging.warning(
                        "Attempt %d/%d failed for %s: %s",
                        item["retries"],
                        MAX_RETRIES,
                        file_path.name,
                        exc,
                    )

                    if item["retries"] >= MAX_RETRIES:
                        logging.exception(
                            "Giving up on %s",
                            file_path.name,
                        )
                        waiting.pop(path, None)
                    else:
                        item["since"] = now

            time.sleep(0.5)

    except KeyboardInterrupt:
        logging.info("Stopping cleaner...")

    finally:
        observer.stop()
        observer.join()
        logging.info("Cleaner stopped.")


if __name__ == "__main__":
    main()