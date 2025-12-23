import csv
import datetime
import os
import shutil
from dataclasses import dataclass
from typing import List, Optional, Tuple
from string import Template

# =========================
#  COMPANY / PATH SETTINGS
# =========================

COMPANY_NAME = "Cienega General Contractor Inc."
COMPANY_CITY_LINE = "Redwood City, CA 94063"
COMPANY_PHONE = "(650) 346-8373"
COMPANY_EMAIL = "cienegageneralcontractor@gmail.com"
COMPANY_LICENSE = "CL#: 1085341"

# Put your logo PNG/JPG file in the SAME folder as this script
COMPANY_LOGO_FILENAME = "CienegaLogo.png"   # or "" to hide logo

# Root folder where all jobs live; default is Desktop/CienegaJobs
ROOT_BASE_DIR = os.path.join(os.path.expanduser("~"), "Desktop", "CienegaJobs")

# =========================
#  CONSOLE COLORS
# =========================

COLOR_HEADER = "\033[95m"   # magenta-ish for big headers
COLOR_SECTION = "\033[94m"  # blue for section titles
COLOR_PROMPT = "\033[92m"   # green for user prompts
COLOR_RESET = "\033[0m"     # reset


def colored_prompt(msg: str) -> str:
    """Wrap an input prompt message in green."""
    return f"{COLOR_PROMPT}{msg}{COLOR_RESET}"


def colored_header(msg: str) -> str:
    return f"{COLOR_HEADER}{msg}{COLOR_RESET}"


def colored_section(msg: str) -> str:
    return f"{COLOR_SECTION}{msg}{COLOR_RESET}"


# =========================
#  DATE HELPERS (MM-DD-YYYY)
# =========================

DATE_FORMAT = "%m-%d-%Y"


def format_date(date_obj: datetime.date) -> str:
    """Return MM-DD-YYYY string."""
    return date_obj.strftime(DATE_FORMAT)


def parse_mmddyyyy(date_str: str) -> datetime.date:
    """Parse MM-DD-YYYY into a date object."""
    return datetime.datetime.strptime(date_str, DATE_FORMAT).date()


# =========================
#  DATA MODELS
# =========================

@dataclass
class EstimateLine:
    """
    For both quotes & invoices.

    For QUOTES:
      - Standard pricing:
          labor lines lump-sum (hours=1, rate=total) or materials w/ qty & unit price.
      - Fixed pricing:
          scope-only lines with blank numeric columns (hours=0, rate=0).
      - 'worker_type' is just the label for the Worker/Type column
        ("Labor", "Material", or "Labor and Material").

    For INVOICES:
      - Labor lines:
          worker_type = Contractor / Carpenter / Laborer / Custom
          hours = hours worked
          rate = hourly rate
      - Material lines same as above.
    """
    section: str = ""       # subheading like "Interior Paint", "Framing & Demo"
    description: str = ""
    worker_type: str = ""   # label shown in Worker / Type column
    hours: float = 0.0
    rate: float = 0.0
    detail: str = ""

    @property
    def total(self) -> float:
        return self.hours * self.rate


@dataclass
class OwnerAllowance:
    """
    Items the client will purchase directly (appliances, finishes, etc.)
    Shown separately and NOT included in contractor totals.
    """
    description: str
    quantity: float
    unit_cost: float

    @property
    def total(self) -> float:
        return self.quantity * self.unit_cost


# =========================
#  FILESYSTEM / JOB HELPERS
# =========================

def ensure_root_dir() -> None:
    os.makedirs(ROOT_BASE_DIR, exist_ok=True)


def _job_counter_path() -> str:
    return os.path.join(ROOT_BASE_DIR, "job_counter.txt")


def get_next_job_id() -> str:
    """
    Finds the next available Job ID.

    - Scans existing job folders in ROOT_BASE_DIR whose names start with "J####"
    - Reads job_counter.txt if present
    - Uses the highest number from both, then returns the next one (e.g. max 4 -> J0005)
    """
    ensure_root_dir()
    max_n = 0

    # Look at existing job folders
    for name in os.listdir(ROOT_BASE_DIR):
        path = os.path.join(ROOT_BASE_DIR, name)
        if not os.path.isdir(path):
            continue
        prefix = name[:5]
        if len(prefix) == 5 and prefix[0] == "J" and prefix[1:].isdigit():
            value = int(prefix[1:])
            if value > max_n:
                max_n = value

    # Look at job_counter.txt
    counter_path = _job_counter_path()
    if os.path.exists(counter_path):
        with open(counter_path, "r", encoding="utf-8") as f:
            txt = f.read().strip()
            if txt.isdigit():
                value = int(txt)
                if value > max_n:
                    max_n = value

    next_n = max_n + 1

    with open(counter_path, "w", encoding="utf-8") as f:
        f.write(str(next_n))

    return f"J{next_n:04d}"


def sanitize_name(name: str) -> str:
    safe = "".join(c for c in name if c.isalnum() or c in (" ", "_", "-", "."))
    safe = safe.strip().replace(" ", "_")
    return safe or "Job"


def get_job_folder(job_id: str, project_name: Optional[str] = None) -> str:
    """
    Returns the folder path for a job and creates it if needed.
    Folder name: "J0001 - Bathroom_Remodel"
    """
    ensure_root_dir()
    if project_name:
        folder_name = f"{job_id} - {sanitize_name(project_name)}"
    else:
        folder_name = job_id
    job_dir = os.path.join(ROOT_BASE_DIR, folder_name)
    os.makedirs(job_dir, exist_ok=True)
    return job_dir


def find_existing_job_folders(job_id: str) -> List[str]:
    """
    Returns a list of full paths to job folders that start with the given job_id.
    Example matches for job_id 'J0005':
      'J0005'
      'J0005 - Bathroom_Remodel'
    """
    ensure_root_dir()
    matches: List[str] = []
    for name in os.listdir(ROOT_BASE_DIR):
        path = os.path.join(ROOT_BASE_DIR, name)
        if os.path.isdir(path) and name.startswith(job_id):
            matches.append(path)
    matches.sort()
    return matches


def show_job_folder_contents(job_dir: str) -> None:
    """
    Prints out what files exist in a specific job folder.
    """
    print(colored_section(f"\nExisting files in job folder: {os.path.basename(job_dir)}"))

    all_files = sorted(
        f for f in os.listdir(job_dir)
        if os.path.isfile(os.path.join(job_dir, f))
    )

    if not all_files:
        print("  (No files in this job folder yet.)\n")
        return

    quotes_csv: List[str] = []
    quotes_html: List[str] = []
    invoices_csv: List[str] = []
    invoices_html: List[str] = []
    receipts: List[str] = []
    others: List[str] = []

    for fname in all_files:
        lower = fname.lower()
        if "_quote_" in lower and lower.endswith(".csv"):
            quotes_csv.append(fname)
        elif "_quote_" in lower and lower.endswith(".html"):
            quotes_html.append(fname)
        elif "_invoice_" in lower and lower.endswith(".csv"):
            invoices_csv.append(fname)
        elif "_invoice_" in lower and lower.endswith(".html"):
            invoices_html.append(fname)
        elif "receipts" in lower and lower.endswith(".csv"):
            receipts.append(fname)
        else:
            others.append(fname)

    total_quotes = len(quotes_csv) + len(quotes_html)
    total_invoices = len(invoices_csv) + len(invoices_html)
    total_receipts = len(receipts)
    total_others = len(others)

    print(
        f"  Summary: {total_quotes} quote file(s), "
        f"{total_invoices} invoice file(s), "
        f"{total_receipts} receipts file(s), "
        f"{total_others} other file(s)."
    )

    def _print_group(label: str, items: List[str]) -> None:
        if not items:
            return
        print(f"  {label}:")
        for x in items:
            print(f"    - {x}")

    _print_group("Quotes (CSV)", quotes_csv)
    _print_group("Quotes (HTML)", quotes_html)
    _print_group("Invoices (CSV)", invoices_csv)
    _print_group("Invoices (HTML)", invoices_html)
    _print_group("Receipts", receipts)
    _print_group("Other files", others)

    print("")


# =========================
#  RECEIPT FILE HELPERS
# =========================

def list_receipt_files(job_dir: str, job_id: str) -> List[str]:
    files: List[str] = []
    for fname in os.listdir(job_dir):
        if not fname.lower().endswith(".csv"):
            continue
        lower = fname.lower()
        if "receipts" in lower and fname.startswith(job_id):
            files.append(fname)
    files.sort()
    return files


def generate_new_receipts_path(job_dir: str, job_id: str) -> str:
    base = os.path.join(job_dir, f"{job_id}_receipts.csv")
    if not os.path.exists(base):
        return base

    i = 2
    while True:
        path = os.path.join(job_dir, f"{job_id}_receipts_{i}.csv")
        if not os.path.exists(path):
            return path
        i += 1


def choose_receipt_file_for_logging(job_dir: str, job_id: str) -> str:
    existing = list_receipt_files(job_dir, job_id)

    if not existing:
        path = generate_new_receipts_path(job_dir, job_id)
        print(colored_section(f"\nNo receipts files found for {job_id}."))
        print(f"Creating new receipts file: {os.path.basename(path)}")
        return path

    print(colored_section(f"\nExisting receipts files for {job_id}:"))
    for idx, fname in enumerate(existing, start=1):
        print(f"  {idx}. {fname}")
    print("  0. Create a NEW receipts file")

    while True:
        choice = input(colored_prompt(f"Select receipts file (0–{len(existing)}): ")).strip()
        try:
            n = int(choice)
        except ValueError:
            print("Please enter a valid number.")
            continue
        if 0 <= n <= len(existing):
            break
        print("Out of range, try again.")

    if n == 0:
        path = generate_new_receipts_path(job_dir, job_id)
        print(f"\nCreating new receipts file: {os.path.basename(path)}")
        return path
    else:
        path = os.path.join(job_dir, existing[n - 1])
        print(f"\nUsing existing receipts file: {os.path.basename(path)}")
        return path


def choose_receipt_file_for_invoice(job_dir: str, job_id: str) -> Optional[str]:
    existing = list_receipt_files(job_dir, job_id)

    if not existing:
        print(colored_section(f"\nNo receipts files found for {job_id}."))
        return None

    if len(existing) == 1:
        path = os.path.join(job_dir, existing[0])
        print(colored_section(f"\nUsing receipts file: {os.path.basename(path)}"))
        return path

    print(colored_section(f"\nMultiple receipts files found for {job_id}:"))
    for idx, fname in enumerate(existing, start=1):
        print(f"  {idx}. {fname}")
    print("  0. Cancel (do not include receipts)")

    while True:
        choice = input(colored_prompt(f"Select receipts file (0–{len(existing)}): ")).strip()
        try:
            n = int(choice)
        except ValueError:
            print("Please enter a valid number.")
            continue
        if 0 <= n <= len(existing):
            break
        print("Out of range, try again.")

    if n == 0:
        print("Skipping receipts for this invoice.")
        return None

    path = os.path.join(job_dir, existing[n - 1])
    print(colored_section(f"\nUsing receipts file: {os.path.basename(path)}"))
    return path


def reorganize_receipts_file(path: str) -> None:
    """
    Reorganize receipts by date then item.
    """
    if not os.path.exists(path):
        return

    try:
        with open(path, "r", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            rows = list(reader)
            fieldnames = reader.fieldnames
    except Exception as e:
        print(f"⚠️ Could not reorganize receipts file: {e}")
        return

    if not rows or not fieldnames:
        return

    def sort_key(row: dict):
        date_str = (row.get("Date") or "").strip()
        try:
            d = parse_mmddyyyy(date_str)
        except ValueError:
            d = datetime.date(9999, 12, 31)
        item = (row.get("Item") or "").strip()
        return (d, item)

    rows.sort(key=sort_key)

    try:
        with open(path, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(rows)
    except Exception as e:
        print(f"⚠️ Could not write reorganized receipts file: {e}")


# =========================
#  INPUT HELPERS
# =========================

def prompt_float(message: str, default: Optional[float] = None) -> float:
    while True:
        if default is not None:
            raw = input(colored_prompt(f"{message} [default {default}]: ")).strip()
            if raw == "":
                return default
        else:
            raw = input(colored_prompt(f"{message}: ")).strip()
        try:
            return float(raw)
        except ValueError:
            print("Please enter a number (e.g. 10 or 10.5).")


def prompt_yes_no(message: str, default: bool = True) -> bool:
    default_str = "Y/n" if default else "y/N"
    while True:
        raw = input(colored_prompt(f"{message} [{default_str}]: ")).strip().lower()
        if raw == "" and default is not None:
            return default
        if raw in ("y", "yes"):
            return True
        if raw in ("n", "no"):
            return False
        print("Please answer y or n.")


def choose_worker_type() -> str:
    print("\nSelect worker type:")
    print("  1. Contractor ($80/hr default)")
    print("  2. Carpenter ($70/hr default)")
    print("  3. Laborer  ($55/hr default)")
    print("  4. Custom")

    while True:
        choice = input(colored_prompt("Choice (1–4): ")).strip()
        if choice in ("1", "2", "3", "4"):
            mapping = {
                "1": "Contractor",
                "2": "Carpenter",
                "3": "Laborer",
                "4": "Custom",
            }
            return mapping[choice]
        print("Invalid choice, please enter 1, 2, 3, or 4.")


def get_default_rate(worker_type: str) -> float:
    if worker_type == "Contractor":
        return 80.0
    if worker_type == "Carpenter":
        return 70.0
    if worker_type == "Laborer":
        return 55.0
    return 0.0


def choose_rate(worker_type: str) -> float:
    default_rate = get_default_rate(worker_type)
    if worker_type == "Custom":
        return prompt_float("Enter custom hourly rate ($)")
    use_default = prompt_yes_no(
        f"Use default rate for {worker_type} (${default_rate:.2f}/hr)?", default=True
    )
    if use_default:
        return default_rate
    return prompt_float("Enter custom hourly rate ($)")


def choose_quote_worker_label() -> str:
    """
    For QUOTES (both standard and fixed):
    Let the user choose how the line is labeled in the Worker / Type column.
    """
    print("  How should this line be labeled in the 'Worker / Type' column?")
    print("    1. Labor")
    print("    2. Material")
    print("    3. Labor and Material")
    while True:
        choice = input(colored_prompt("  Choice (1–3, default 1): ")).strip()
        if choice == "" or choice == "1":
            return "Labor"
        if choice == "2":
            return "Material"
        if choice == "3":
            return "Labor and Material"
        print("Please enter 1, 2, or 3.")


def prompt_float_keep_current(label: str, current: float) -> float:
    """
    Ask for a number; ENTER keeps current value.
    """
    while True:
        raw = input(colored_prompt(f"{label} (current {current:,.2f}, ENTER to keep): ")).strip()
        if raw == "":
            return current
        try:
            return float(raw)
        except ValueError:
            print("Please enter a valid number, or press ENTER to keep current value.")


def edit_quote_headers_interactive(meta: dict, totals: dict) -> Tuple[dict, dict]:
    """
    Allows editing of the main quote header fields and totals.
    meta = {
        'project_name', 'client_name', 'project_address',
        'doc_date' (datetime.date), 'receipts_total'
    }
    totals = {
        'subtotal', 'receipts_total', 'grand_total'
    }
    """
    print(colored_section("\n=== Edit Quote Header Fields ==="))

    # PROJECT NAME
    print(f"Current project name: {meta['project_name']}")
    new_val = input(colored_prompt("New project name (ENTER to keep): ")).strip()
    if new_val:
        meta["project_name"] = new_val

    # CLIENT NAME
    print(f"Current client name: {meta['client_name']}")
    new_val = input(colored_prompt("New client name (ENTER to keep): ")).strip()
    if new_val:
        meta["client_name"] = new_val

    # PROJECT ADDRESS
    print(f"Current project address: {meta['project_address']}")
    new_val = input(colored_prompt("New address (ENTER to keep): ")).strip()
    if new_val:
        meta["project_address"] = new_val

    # DATE
    print(f"Current document date: {format_date(meta['doc_date'])}")
    while True:
        new_val = input(colored_prompt("New date (MM-DD-YYYY, ENTER to keep): ")).strip()
        if not new_val:
            break
        try:
            meta["doc_date"] = parse_mmddyyyy(new_val)
            break
        except ValueError:
            print("Invalid date format. Use MM-DD-YYYY.")

    # RECEIPTS TOTAL
    print(f"Current receipts total: {totals.get('receipts_total', 0.0):,.2f}")
    raw = input(colored_prompt("New receipts total (ENTER to keep): ")).strip()
    if raw:
        try:
            totals["receipts_total"] = float(raw)
        except ValueError:
            print("Invalid number — keeping current value.")

    # SUBTOTAL
    print(f"Current subtotal: {totals.get('subtotal', 0.0):,.2f}")
    raw = input(colored_prompt("New subtotal (ENTER to keep): ")).strip()
    if raw:
        try:
            totals["subtotal"] = float(raw)
        except ValueError:
            print("Invalid number — keeping current value.")

    # GRAND TOTAL
    print(f"Current grand total: {totals.get('grand_total', 0.0):,.2f}")
    print("Press ENTER to auto-calc (subtotal + receipts).")
    raw = input(colored_prompt("New grand total (ENTER for auto): ")).strip()

    if raw:
        try:
            totals["grand_total"] = float(raw)
        except ValueError:
            print("Invalid number — recalculating.")
            totals["grand_total"] = totals["subtotal"] + totals["receipts_total"]
    else:
        totals["grand_total"] = totals["subtotal"] + totals["receipts_total"]

    return meta, totals


# =========================
#  TOTALS (NO MARKUP)
# =========================

def compute_totals(
    lines: List[EstimateLine],
    receipts_total: float = 0.0
) -> dict:
    contractor_subtotal = sum(l.total for l in lines)
    grand_total = contractor_subtotal + receipts_total
    return {
        "subtotal": contractor_subtotal,
        "receipts_total": receipts_total,
        "grand_total": grand_total,
    }


# =========================
#  CSV EXPORT
# =========================

def save_document_csv(
    job_dir: str,
    job_id: str,
    doc_type: str,
    project_name: str,
    project_address: str,
    client_name: str,
    lines: List[EstimateLine],
    totals: dict,
    allowances: List[OwnerAllowance],
    doc_date: datetime.date,
) -> str:
    date_str = format_date(doc_date)
    filename = f"{job_id}_{doc_type}_{date_str}.csv"
    path = os.path.join(job_dir, filename)

    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["DocType", doc_type])
        writer.writerow(["JobID", job_id])
        writer.writerow(["Project", project_name])
        writer.writerow(["Client", client_name])
        writer.writerow(["Address", project_address])
        writer.writerow(["Date", date_str])
        writer.writerow(["ReceiptsTotal", totals.get("receipts_total", 0.0)])
        writer.writerow([])
        writer.writerow([
            "Section",
            "Description",
            "FullDescription",
            "WorkerType",
            "Hours/Qty",
            "Rate/UnitPrice",
            "LineTotal",
        ])

        for line in lines:
            writer.writerow(
                [
                    line.section,
                    line.description,
                    line.detail,
                    line.worker_type,
                    line.hours,
                    line.rate,
                    line.total,
                ]
            )

        writer.writerow([])
        writer.writerow(["Subtotal", "", "", "", "", "", totals["subtotal"]])
        writer.writerow(["GrandTotal", "", "", "", "", "", totals["grand_total"]])

        if allowances:
            writer.writerow([])
            writer.writerow(["Owner Allowances (client purchases directly)"])
            writer.writerow(["Description", "Quantity", "EstUnitCost", "EstTotal"])
            allowance_total = 0.0
            for a in allowances:
                writer.writerow([a.description, a.quantity, a.unit_cost, a.total])
                allowance_total += a.total
            writer.writerow(["OwnerAllowancesTotal", "", "", "", allowance_total])

    return path


def load_quote_from_csv(path: str) -> Tuple[dict, List[EstimateLine], dict, List[OwnerAllowance]]:
    """
    Load an existing QUOTE CSV that was created by save_document_csv.
    Returns:
      meta: {
          'doc_type', 'job_id', 'project_name', 'client_name',
          'project_address', 'doc_date'
      }
      lines: List[EstimateLine]
      totals: {
          'subtotal', 'receipts_total', 'grand_total'
      }
      allowances: List[OwnerAllowance]
    """
    if not os.path.exists(path):
        print(f"⚠️ Quote file not found: {path}")
        return {}, [], {}, []

    with open(path, "r", encoding="utf-8") as f:
        reader = csv.reader(f)
        rows = list(reader)

    meta = {
        "doc_type": "quote",
        "job_id": "",
        "project_name": "",
        "client_name": "",
        "project_address": "",
        "doc_date": datetime.date.today(),
    }
    totals = {
        "subtotal": 0.0,
        "receipts_total": 0.0,
        "grand_total": 0.0,
    }
    lines: List[EstimateLine] = []
    allowances: List[OwnerAllowance] = []

    i = 0
    while i < len(rows):
        row = rows[i]
        if not row:
            i += 1
            continue
        key = row[0].strip()

        if key == "DocType" and len(row) > 1:
            meta["doc_type"] = row[1].strip()
        elif key == "JobID" and len(row) > 1:
            meta["job_id"] = row[1].strip()
        elif key == "Project" and len(row) > 1:
            meta["project_name"] = row[1].strip()
        elif key == "Client" and len(row) > 1:
            meta["client_name"] = row[1].strip()
        elif key == "Address" and len(row) > 1:
            meta["project_address"] = row[1].strip()
        elif key == "Date" and len(row) > 1:
            try:
                meta["doc_date"] = parse_mmddyyyy(row[1].strip())
            except ValueError:
                pass
        elif key == "ReceiptsTotal" and len(row) > 1:
            try:
                totals["receipts_total"] = float(row[1])
            except (ValueError, TypeError):
                totals["receipts_total"] = 0.0

        elif key == "Section":
            # Start of line items table
            i += 1
            while i < len(rows):
                r = rows[i]
                if (not r) or (r[0].strip() in (
                    "Subtotal",
                    "GrandTotal",
                    "Owner Allowances (client purchases directly)",
                )):
                    break

                if all(not cell.strip() for cell in r):
                    i += 1
                    continue

                section = r[0].strip() if len(r) > 0 else ""
                desc    = r[1].strip() if len(r) > 1 else ""
                detail  = r[2].strip() if len(r) > 2 else ""
                worker  = r[3].strip() if len(r) > 3 else ""
                try:
                    hours = float(r[4]) if len(r) > 4 and r[4].strip() != "" else 0.0
                except ValueError:
                    hours = 0.0
                try:
                    rate = float(r[5]) if len(r) > 5 and r[5].strip() != "" else 0.0
                except ValueError:
                    rate = 0.0

                lines.append(
                    EstimateLine(
                        section=section,
                        description=desc,
                        worker_type=worker,
                        hours=hours,
                        rate=rate,
                        detail=detail,
                    )
                )
                i += 1
            continue

        elif key == "Subtotal":
            try:
                totals["subtotal"] = float(row[-1])
            except (ValueError, TypeError):
                totals["subtotal"] = 0.0
        elif key == "GrandTotal":
            try:
                totals["grand_total"] = float(row[-1])
            except (ValueError, TypeError):
                totals["grand_total"] = 0.0
        elif key == "Owner Allowances (client purchases directly)":
            i += 2
            while i < len(rows):
                r = rows[i]
                if not r:
                    i += 1
                    continue
                first = r[0].strip()
                if first == "OwnerAllowancesTotal":
                    break
                desc = r[0].strip() if len(r) > 0 else ""
                try:
                    qty = float(r[1]) if len(r) > 1 and r[1].strip() != "" else 0.0
                except ValueError:
                    qty = 0.0
                try:
                    unit = float(r[2]) if len(r) > 2 and r[2].strip() != "" else 0.0
                except ValueError:
                    unit = 0.0
                allowances.append(OwnerAllowance(desc, qty, unit))
                i += 1

        i += 1

    return meta, lines, totals, allowances


def backup_original_quote_files(csv_path: str) -> None:
    """
    Create backup copies of the original quote CSV and any matching HTML
    BEFORE writing revised versions.
    """
    if not os.path.exists(csv_path):
        return

    job_dir = os.path.dirname(csv_path)
    base_name = os.path.basename(csv_path)
    stem, ext = os.path.splitext(base_name)

    # Backup CSV
    backup_csv = os.path.join(job_dir, stem + "_original" + ext)
    if not os.path.exists(backup_csv):
        try:
            shutil.copy2(csv_path, backup_csv)
            print(colored_section(f"Backup created: {os.path.basename(backup_csv)}"))
        except Exception as e:
            print(f"⚠️ Could not create CSV backup: {e}")

    # Backup matching HTML (same stem prefix)
    for fname in os.listdir(job_dir):
        if not fname.lower().endswith(".html"):
            continue
        if not fname.startswith(stem):
            continue
        html_path = os.path.join(job_dir, fname)
        h_stem, h_ext = os.path.splitext(fname)
        backup_html = os.path.join(job_dir, h_stem + "_original" + h_ext)
        if os.path.exists(backup_html):
            continue
        try:
            shutil.copy2(html_path, backup_html)
            print(colored_section(f"Backup created: {os.path.basename(backup_html)}"))
        except Exception as e:
            print(f"⚠️ Could not create HTML backup for {fname}: {e}")


def append_receipt_csv(
    path: str,
    job_id: str,
    item_name: str,
    date_str: str,
    cost: float,
) -> str:
    file_exists = os.path.exists(path)
    with open(path, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        if not file_exists:
            writer.writerow(["JobID", "Item", "Date", "Cost"])
        writer.writerow([job_id, item_name, date_str, cost])
    return path


# =========================
#  RECEIPT IMPORT / FILTERING
# =========================

def load_receipts_data_from_csv(path: str) -> List[dict]:
    rows: List[dict] = []
    if not os.path.exists(path):
        print(f"⚠️ Receipts file not found: {path}")
        return rows

    try:
        with open(path, "r", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            for row in reader:
                rows.append(row)
    except Exception as e:
        print(f"⚠️ Could not read receipts file: {e}")
        return []

    return rows


def _sum_receipt_cost(rows: List[dict]) -> float:
    total = 0.0
    for row in rows:
        raw_cost = row.get("Cost")
        if raw_cost is None and row:
            raw_cost = list(row.values())[-1]
        try:
            total += float(str(raw_cost).strip())
        except (TypeError, ValueError):
            continue
    return total


def _filter_receipts_by_date(rows: List[dict]) -> List[dict]:
    if not rows:
        return []

    print("\nFilter receipts by date range.")
    print("Dates must be in MM-DD-YYYY format.")
    while True:
        start_raw = input(colored_prompt("Start date (MM-DD-YYYY): ")).strip()
        end_raw = input(colored_prompt("End date (MM-DD-YYYY): ")).strip()
        try:
            start_date = parse_mmddyyyy(start_raw)
            end_date = parse_mmddyyyy(end_raw)
        except ValueError:
            print("Invalid date format. Please try again.\n")
            continue
        if end_date < start_date:
            print("End date is before start date. Please try again.\n")
            continue
        break

    filtered: List[dict] = []
    for row in rows:
        date_str = (row.get("Date") or "").strip()
        try:
            d = parse_mmddyyyy(date_str)
        except ValueError:
            continue
        if start_date <= d <= end_date:
            filtered.append(row)
    return filtered


def _filter_receipts_by_lines(rows: List[dict]) -> List[dict]:
    if not rows:
        return []

    print("\nCurrent receipts in file:")
    for idx, row in enumerate(rows, start=1):
        jobid = row.get("JobID", "")
        item = row.get("Item", "")
        date = row.get("Date", "")
        cost = row.get("Cost", "")
        print(f"  {idx}. [{date}] {item} - ${cost} (Job {jobid})")

    print("\nEnter line numbers to include (e.g. 1,3,5-7).")
    selection = input(colored_prompt("Selection: ")).strip()
    if not selection:
        return []

    indices: set[int] = set()
    for part in selection.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            try:
                start_s, end_s = part.split("-", 1)
                start_i = int(start_s)
                end_i = int(end_s)
                if start_i <= end_i:
                    for i in range(start_i, end_i + 1):
                        indices.add(i)
            except ValueError:
                continue
        else:
            try:
                i = int(part)
                indices.add(i)
            except ValueError:
                continue

    filtered: List[dict] = []
    for i in sorted(indices):
        if 1 <= i <= len(rows):
            filtered.append(rows[i - 1])
    return filtered


def choose_receipts_for_invoice(job_dir: str, job_id: str) -> Tuple[List[dict], float]:
    use_receipts = prompt_yes_no(
        "Include receipts from CSV in this invoice?", default=False
    )
    if not use_receipts:
        return [], 0.0

    receipts_path = choose_receipt_file_for_invoice(job_dir, job_id)
    if not receipts_path:
        return [], 0.0

    print(f"\nReading receipts from: {receipts_path}")
    rows = load_receipts_data_from_csv(receipts_path)
    if not rows:
        print("⚠️ No receipt data found in that file. Continuing without receipts.")
        return [], 0.0

    print("\nHow would you like to include receipts?")
    print("  1. All receipts in file")
    print("  2. Filter by date range (MM-DD-YYYY)")
    print("  3. Select specific line numbers")
    while True:
        choice = input(colored_prompt("Choice (1–3): ")).strip()
        if choice in ("1", "2", "3"):
            break
        print("Invalid choice. Please enter 1, 2, or 3.")

    if choice == "2":
        rows_filtered = _filter_receipts_by_date(rows)
    elif choice == "3":
        rows_filtered = _filter_receipts_by_lines(rows)
    else:
        rows_filtered = rows

    if not rows_filtered:
        print("⚠️ No receipts selected. Continuing without receipts.")
        return [], 0.0

    total = _sum_receipt_cost(rows_filtered)
    if total <= 0.0:
        print("⚠️ Selected receipts total is 0. Continuing without receipts.")
        return [], 0.0

    print(f"Receipts total to include in invoice: ${total:,.2f}")
    print("This amount will be added to the grand total and shown in a receipts section.\n")
    return rows_filtered, total


# =========================
#  HTML (QUOTE / INVOICE)
# =========================

HTML_TEMPLATE = Template(r"""
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>$title_text - $project_name</title>
    <style>
        body {
            font-family: system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
            margin: 40px;
            color: #222;
            line-height: 1.5;
        }
        .header {
            display: flex;
            justify-content: space-between;
            align-items: flex-start;
            border-bottom: 2px solid #444;
            padding-bottom: 16px;
            margin-bottom: 24px;
        }
        .logo {
            max-height: 60px;
            margin-bottom: 6px;
        }
        .company-info {
            font-size: 13px;
        }
        .company-name {
            font-size: 22px;
            font-weight: 700;
        }
        .summary-box {
            font-size: 13px;
            min-width: 220px;
        }
        .summary-label {
            text-transform: uppercase;
            font-size: 11px;
            letter-spacing: 0.12em;
            color: #777;
        }
        .summary-value {
            font-size: 18px;
            font-weight: 700;
        }
        .meta-row {
            display: flex;
            gap: 24px;
            font-size: 13px;
            margin-bottom: 10px;
        }
        .meta-col-label {
            font-size: 11px;
            text-transform: uppercase;
            letter-spacing: 0.12em;
            color: #777;
        }
        .meta-col-value {
            margin-top: 2px;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 14px;
            font-size: 13px;
        }
        th, td {
            padding: 8px 6px;
            border-bottom: 1px solid #ddd;
            vertical-align: top;
        }
        th {
            text-align: left;
            font-size: 11px;
            text-transform: uppercase;
            letter-spacing: 0.08em;
            color: #555;
            border-bottom: 2px solid #333;
        }
        td.num {
            text-align: right;
            white-space: nowrap;
        }
        .line-detail {
            font-size: 11px;
            color: #555;
            margin-top: 2px;
        }
        .section-row td {
            background-color: #f4f4f4;
            font-weight: 600;
            border-top: 2px solid #333;
            text-transform: uppercase;
            letter-spacing: 0.08em;
            font-size: 11px;
        }
        .totals {
            margin-top: 18px;
            max-width: 320px;
            margin-left: auto;
            font-size: 13px;
        }
        .totals-row {
            display: flex;
            justify-content: space-between;
            padding: 4px 0;
        }
        .totals-row.grand {
            border-top: 2px solid #333;
            margin-top: 6px;
            padding-top: 8px;
            font-size: 14px;
            font-weight: 700;
        }
        .notes {
            margin-top: 24px;
            font-size: 12px;
        }
        .notes-title {
            font-size: 11px;
            text-transform: uppercase;
            letter-spacing: 0.14em;
            font-weight: 700;
            margin-bottom: 6px;
        }
        .notes ul {
            margin: 6px 0 0 18px;
            padding: 0;
        }
        .notes li {
            margin-bottom: 2px;
        }
        .allowances {
            margin-top: 28px;
            font-size: 13px;
        }
        .allowances-title {
            font-size: 11px;
            text-transform: uppercase;
            letter-spacing: 0.14em;
            font-weight: 700;
            margin-bottom: 6px;
        }
        .allowances-table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 8px;
            font-size: 13px;
        }
        .allowances-table th, .allowances-table td {
            padding: 6px 6px;
            border-bottom: 1px solid #ddd;
            vertical-align: top;
        }
        .footer {
            margin-top: 36px;
            font-size: 10px;
            color: #777;
            border-top: 1px solid #ddd;
            padding-top: 6px;
        }
        .receipt-section {
            margin-top: 24px;
            font-size: 13px;
        }
        .receipt-title {
            font-size: 11px;
            text-transform: uppercase;
            letter-spacing: 0.14em;
            font-weight: 700;
            margin-bottom: 6px;
        }
        .receipt-table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 8px;
            font-size: 12px;
        }
        .receipt-table th, .receipt-table td {
            padding: 6px 6px;
            border-bottom: 1px solid #ddd;
            vertical-align: top;
        }
        .receipt-table th {
            text-align: left;
            font-size: 11px;
            text-transform: uppercase;
            letter-spacing: 0.08em;
            color: #555;
            border-bottom: 2px solid #333;
        }
    </style>
</head>
<body>

<div class="header">
    <div>
        $logo_html
        <div class="company-name">$company_name</div>
        <div class="company-info">
            $company_city_line<br>
            Phone: $company_phone<br>
            Email: $company_email
        </div>
    </div>
    <div class="summary-box">
        <div class="summary-label">$summary_label</div>
        <div class="summary-value">$grand_total_fmt</div>
        <div style="margin-top:6px;">Date: $date_long</div>
        <div>Job ID: $job_id</div>
    </div>
</div>

<div class="meta-row">
    <div>
        <div class="meta-col-label">Project</div>
        <div class="meta-col-value">$project_name</div>
    </div>
    <div>
        <div class="meta-col-label">Client</div>
        <div class="meta-col-value">$client_name</div>
    </div>
    <div>
        <div class="meta-col-label">Address</div>
        <div class="meta-col-value">$project_address</div>
    </div>
    <div>
        <div class="meta-col-label">Contractor Info</div>
        <div class="meta-col-value">
            <strong>$company_license</strong><br>
            $company_city_line<br>
            Phone: $company_phone<br>
            Email: $company_email
        </div>
    </div>
</div>

<table>
    <thead>
        <tr>
            <th style="width: 40%;">Scope Item</th>
            <th style="width: 20%;">Worker / Type</th>
            <th style="width: 15%;">Hours / Qty</th>
            <th style="width: 12%;">Rate / Unit</th>
            <th style="width: 13%;">Line Total</th>
        </tr>
    </thead>
    <tbody>
        $table_rows
    </tbody>
</table>

$receipt_section_html

<div class="totals">
    <div class="totals-row">
        <div>Subtotal</div>
        <div>$subtotal_fmt</div>
    </div>
    $receipts_row_html
    <div class="totals-row grand">
        <div>Grand Total</div>
        <div>$grand_total_fmt</div>
    </div>
</div>

$owner_allowance_html

<div class="notes">
    <div class="notes-title">Notes & Conditions</div>
    <ul>
        <li>Proposal/Invoice is valid for 30 days unless otherwise noted.</li>
        <li>Any additional work or materials not included will be charged via additional quote/invoice.</li>
        <li>All work to be performed in a workmanlike manner according to standard practices.</li>
    </ul>
</div>

<div class="footer">
    This document is provided by $company_name. Pricing is based on current material and labor rates and may be subject to change.
</div>

</body>
</html>
""")


def render_estimate_html(
    job_id: str,
    doc_type: str,
    project_name: str,
    project_address: str,
    client_name: str,
    lines: List[EstimateLine],
    totals: dict,
    allowances: List[OwnerAllowance],
    doc_date: datetime.date,
    receipt_rows: Optional[List[dict]] = None,
    hide_line_prices: bool = False,
) -> str:
    date_long = format_date(doc_date)
    if COMPANY_LOGO_FILENAME and os.path.exists(COMPANY_LOGO_FILENAME):
        logo_html = f'<img class="logo" src="{COMPANY_LOGO_FILENAME}" alt="Logo">'
    else:
        logo_html = ""

    if doc_type == "quote":
        title_text = "Quote"
        summary_label = "Quote Total"
    else:
        title_text = "Invoice"
        summary_label = "Amount Due"

    row_html_list = []
    last_section = None
    for line in lines:
        if line.section and line.section != last_section:
            row_html_list.append(
                f"""
        <tr class="section-row">
            <td colspan="5">{line.section}</td>
        </tr>
        """
            )
            last_section = line.section

        if hide_line_prices:
            hours_str = ""
            rate_str = ""
            total_str = ""
        else:
            if doc_type == "quote" and line.worker_type.lower() == "labor":
                hours_str = ""
                rate_str = ""
            else:
                hours_str = f"{line.hours:.2f}" if line.hours != 0 else ""
                rate_str = f"${line.rate:,.2f}" if line.rate != 0 else ""
            total_str = f"${line.total:,.2f}" if line.total != 0 else ""

        if doc_type == "quote" and line.detail.strip():
            desc_html = f'{line.description}<div class="line-detail">{line.detail}</div>'
        else:
            desc_html = line.description

        row_html_list.append(
            f"""
        <tr>
            <td>{desc_html}</td>
            <td>{line.worker_type}</td>
            <td class="num">{hours_str}</td>
            <td class="num">{rate_str}</td>
            <td class="num">{total_str}</td>
        </tr>
        """
        )
    table_rows = "\n".join(row_html_list)

    receipt_rows = receipt_rows or []
    if doc_type == "invoice" and receipt_rows:
        headers = list(receipt_rows[0].keys())
        header_cells = "".join(f"<th>{h}</th>" for h in headers)
        body_rows = []
        for r in receipt_rows:
            cells = "".join(f"<td>{r.get(h, '')}</td>" for h in headers)
            body_rows.append(f"<tr>{cells}</tr>")
        body_html = "\n".join(body_rows)
        receipt_section_html = f"""
<div class="receipt-section">
    <div class="receipt-title">Receipts</div>
    <table class="receipt-table">
        <thead>
            <tr>{header_cells}</tr>
        </thead>
        <tbody>
            {body_html}
        </tbody>
    </table>
</div>
"""
    else:
        receipt_section_html = ""

    receipts_total = totals.get("receipts_total", 0.0)
    if receipts_total and receipts_total > 0.0:
        receipts_row_html = f"""
    <div class="totals-row">
        <div>Receipts Total</div>
        <div>${receipts_total:,.2f}</div>
    </div>
"""
    else:
        receipts_row_html = ""

    if allowances:
        allowance_rows = []
        total_allowances = 0.0
        for a in allowances:
            allowance_rows.append(
                f"""
            <tr>
                <td>{a.description}</td>
                <td class="num">{a.quantity:.2f}</td>
                <td class="num">${a.unit_cost:,.2f}</td>
                <td class="num">${a.total:,.2f}</td>
            </tr>
            """
            )
            total_allowances += a.total

        allowance_rows_html = "\n".join(allowance_rows)
        owner_allowance_html = f"""
<div class="allowances">
    <div class="allowances-title">Owner Allowances (items client pays for directly)</div>
    <table class="allowances-table">
        <thead>
            <tr>
                <th style="width: 40%;">Description</th>
                <th style="width: 15%;">Qty</th>
                <th style="width: 20%;">Est. Unit Cost</th>
                <th style="width: 25%;">Est. Total</th>
            </tr>
        </thead>
        <tbody>
            {allowance_rows_html}
            <tr>
                <td colspan="3" style="text-align:right; font-weight:700;">Total Owner Allowances (client responsibility)</td>
                <td class="num" style="font-weight:700;">${total_allowances:,.2f}</td>
            </tr>
        </tbody>
    </table>
</div>
"""
    else:
        owner_allowance_html = ""

    html = HTML_TEMPLATE.substitute(
        title_text=title_text,
        project_name=project_name,
        project_address=project_address,
        client_name=client_name,
        job_id=job_id,
        logo_html=logo_html,
        company_name=COMPANY_NAME,
        company_city_line=COMPANY_CITY_LINE,
        company_phone=COMPANY_PHONE,
        company_email=COMPANY_EMAIL,
        company_license=COMPANY_LICENSE,
        summary_label=summary_label,
        date_long=date_long,
        subtotal_fmt=f"${totals['subtotal']:,.2f}",
        grand_total_fmt=f"${totals['grand_total']:,.2f}",
        table_rows=table_rows,
        owner_allowance_html=owner_allowance_html,
        receipts_row_html=receipts_row_html,
        receipt_section_html=receipt_section_html,
    )
    return html


def save_estimate_html(
    job_dir: str,
    job_id: str,
    doc_type: str,
    project_name: str,
    html: str,
    doc_date: datetime.date,
) -> str:
    date_str = format_date(doc_date)
    filename = f"{job_id}_{doc_type}_{sanitize_name(project_name)}_{date_str}.html"
    path = os.path.join(job_dir, filename)

    if COMPANY_LOGO_FILENAME:
        logo_source = os.path.abspath(COMPANY_LOGO_FILENAME)
        if os.path.exists(logo_source):
            logo_dest = os.path.join(job_dir, os.path.basename(COMPANY_LOGO_FILENAME))
            if not os.path.exists(logo_dest):
                try:
                    shutil.copy2(logo_source, logo_dest)
                except Exception as e:
                    print(f"⚠️ Could not copy logo to job folder: {e}")

    with open(path, "w", encoding="utf-8") as f:
        f.write(html)
    return path


# =========================
#  SECTION OVERVIEW & REORDER
# =========================

def show_sections_overview(lines: List[EstimateLine]) -> None:
    """
    Print a grouped overview by section in current order.
    """
    if not lines:
        print("No line items.")
        return

    section_order: List[str] = []
    for l in lines:
        key = l.section or ""
        if key not in section_order:
            section_order.append(key)

    print(colored_section("\nCurrent section overview (in order):"))
    for idx, sec in enumerate(section_order, start=1):
        label = sec if sec else "[No Section]"
        print(f"\n  {idx}. {label}")
        for line in lines:
            if (line.section or "") == sec:
                print(f"     - {line.description} ({line.worker_type})")


def reorder_sections_interactive(lines: List[EstimateLine]) -> List[EstimateLine]:
    """
    Let the user see section overview and reorder them.
    Returns a new lines list with sections regrouped in new order.
    """
    if not lines:
        return lines

    section_order: List[str] = []
    for l in lines:
        key = l.section or ""
        if key not in section_order:
            section_order.append(key)

    if len(section_order) <= 1:
        return lines

    show_sections_overview(lines)

    if not prompt_yes_no("Would you like to reorder these sections?", default=False):
        return lines

    print(colored_section("\nSections (current order):"))
    for idx, sec in enumerate(section_order, start=1):
        label = sec if sec else "[No Section]"
        print(f"  {idx}. {label}")

    print("\nEnter a new order for sections using their numbers.")
    print("Example: if you see 1. Interior, 2. Exterior, 3. Framing")
    print("and you want Framing first, then Interior, then Exterior,")
    print("you would enter: 3,1,2\n")

    while True:
        raw = input(colored_prompt("New section order (comma-separated): ")).strip()
        parts = [p.strip() for p in raw.split(",") if p.strip()]
        try:
            indices = [int(p) for p in parts]
        except ValueError:
            print("Please enter only numbers separated by commas.")
            continue

        if len(indices) != len(section_order):
            print(f"Please provide exactly {len(section_order)} numbers.")
            continue
        if not all(1 <= i <= len(section_order) for i in indices):
            print("One or more numbers are out of range.")
            continue
        if len(set(indices)) != len(indices):
            print("Each section number should be used exactly once.")
            continue

        break

    new_section_order = [section_order[i - 1] for i in indices]

    by_section = {sec: [] for sec in section_order}
    for l in lines:
        key = l.section or ""
        by_section[key].append(l)

    new_lines: List[EstimateLine] = []
    for sec in new_section_order:
        new_lines.extend(by_section.get(sec, []))

    print(colored_section("\nSections have been reordered.\n"))
    show_sections_overview(new_lines)

    return new_lines


# =========================
#  BUILD LINE ITEMS
# =========================

def build_quote_lines() -> List[EstimateLine]:
    """
    QUOTES (standard pricing)
    """
    lines: List[EstimateLine] = []
    current_section = ""

    print(colored_section("\nAdd Quote Line Items (Standard Pricing)."))
    print("You can create section headings (e.g. 'Interior Paint', 'Framing & Demo').")

    while True:
        print(colored_section("\nMenu:"))
        print("  1. Set / Change Section Heading")
        print("  2. Add Labor Line (fixed amount)")
        print("  3. Add Material Line")
        print("  4. Done adding quote items")
        choice = input(colored_prompt("Choice (1–4): ")).strip()

        if choice == "1":
            section_name = input(colored_prompt("Section name (e.g. Interior Painting): ")).strip()
            current_section = section_name
            print(f"Current section set to: {current_section or '(none)'}")

        elif choice == "2":
            print("\nLabor (fixed amount for quote)")
            desc = input(colored_prompt("Scope item / short label: ")).strip()
            if not desc:
                print("Description required.")
                continue
            full_desc = input(
                colored_prompt("Detailed description for quote (shows under the line item): ")
            ).strip()
            total_amount = prompt_float("Total labor amount for this item")

            worker_label = choose_quote_worker_label()

            line = EstimateLine(
                section=current_section,
                description=desc,
                worker_type=worker_label,
                hours=1.0,
                rate=total_amount,
                detail=full_desc,
            )
            lines.append(line)
            print(f"  Added labor line: {desc} | Label = {worker_label} | Amount = ${line.total:,.2f}\n")

        elif choice == "3":
            print("\nMaterial item (quote)")
            desc = input(colored_prompt("Scope item / short label (e.g. tile, lumber, fixtures): ")).strip()
            if not desc:
                print("Description required.")
                continue
            full_desc = input(
                colored_prompt("Detailed description for this material line (optional): ")
            ).strip()
            qty = prompt_float("Quantity")
            price = prompt_float("Unit price ($)")

            worker_label = choose_quote_worker_label()

            line = EstimateLine(
                section=current_section,
                description=desc,
                worker_type=worker_label,
                hours=qty,
                rate=price,
                detail=full_desc,
            )
            lines.append(line)
            print(f"  Added material line: {desc} | Label = {worker_label} | {qty} @ ${price:.2f} = ${line.total:,.2f}\n")

        elif choice == "4":
            break
        else:
            print("Invalid choice, please enter 1, 2, 3, or 4.")

    lines = reorder_sections_interactive(lines)
    return lines


def build_invoice_lines() -> List[EstimateLine]:
    """
    INVOICES (Time & Materials)
    """
    lines: List[EstimateLine] = []
    current_section = ""

    print(colored_section("\nAdd Invoice Line Items (Time & Materials)."))
    print("You can create section headings (e.g. 'Interior Painting', 'Framing & Demo').")

    while True:
        print(colored_section("\nMenu:"))
        print("  1. Set / Change Section Heading")
        print("  2. Add Labor Line (with rate)")
        print("  3. Add Material Line")
        print("  4. Done adding invoice items")
        choice = input(colored_prompt("Choice (1–4): ")).strip()

        if choice == "1":
            section_name = input(colored_prompt("Section name (e.g. Framing & Demo): ")).strip()
            current_section = section_name
            print(f"Current section set to: {current_section or '(none)'}")

        elif choice == "2":
            print("\nLabor (invoice)")
            desc = input(colored_prompt("Scope / description (e.g. Demo, Framing): ")).strip()
            if not desc:
                print("Description required.")
                continue
            worker_type = choose_worker_type()
            rate = choose_rate(worker_type)
            hours = prompt_float("Hours for this line item")
            line = EstimateLine(
                section=current_section,
                description=desc,
                worker_type=worker_type,
                hours=hours,
                rate=rate,
            )
            lines.append(line)
            print(
                f"  Added labor: {desc} | {worker_type} | "
                f"{hours:.2f} hrs @ ${rate:.2f}/hr = ${line.total:,.2f}\n"
            )

        elif choice == "3":
            print("\nMaterial item (invoice)")
            desc = input(colored_prompt("Scope / description (e.g. lumber, drywall, tile): ")).strip()
            if not desc:
                print("Description required.")
                continue
            qty = prompt_float("Quantity")
            price = prompt_float("Unit price ($)")
            line = EstimateLine(
                section=current_section,
                description=desc,
                worker_type="Material",
                hours=qty,
                rate=price,
            )
            lines.append(line)
            print(f"  Added material: {desc} | {qty} @ ${price:.2f} = ${line.total:,.2f}\n")

        elif choice == "4":
            break
        else:
            print("Invalid choice, please enter 1, 2, 3, or 4.")

    lines = reorder_sections_interactive(lines)
    return lines


def build_scope_only_lines(context_label: str) -> List[EstimateLine]:
    """
    For fixed pricing (quote or invoice)
    """
    lines: List[EstimateLine] = []
    current_section = ""

    print(colored_section(f"\nAdding scope-only line items for fixed-price {context_label}."))
    print("You can set section headings like 'Interior Paint', 'Exterior Paint', 'Framing & Demo'.")
    print("These lines will show scope only, with blank pricing columns.\n")

    while True:
        print(colored_section("Menu:"))
        print("  1. Set / Change Section Heading")
        print("  2. Add Scope Line")
        print("  3. Done adding scope items")
        choice = input(colored_prompt("Choice (1–3): ")).strip()

        if choice == "1":
            section_name = input(colored_prompt("Section name (e.g. Interior Painting): ")).strip()
            current_section = section_name
            print(f"Current section set to: {current_section or '(none)'}")

        elif choice == "2":
            desc = input(colored_prompt("Scope item / short label (or ENTER to finish): ")).strip()
            if not desc:
                continue

            worker_label = choose_quote_worker_label()

            if context_label.lower() == "quote":
                detail = input(colored_prompt("  Detailed description for quote (optional): ")).strip()
            else:
                detail = ""

            line = EstimateLine(
                section=current_section,
                description=desc,
                worker_type=worker_label,
                hours=0.0,
                rate=0.0,
                detail=detail,
            )
            lines.append(line)
            print(f"  Added scope-only line: {desc} ({worker_label})\n")

        elif choice == "3":
            break
        else:
            print("Invalid choice. Please enter 1, 2, or 3.")

    lines = reorder_sections_interactive(lines)
    return lines


def build_owner_allowances() -> List[OwnerAllowance]:
    allowances: List[OwnerAllowance] = []

    print(colored_section("\nOwner Allowances (items client will purchase directly)."))
    print("Press ENTER on description to finish.\n")

    while True:
        desc = input(colored_prompt("Allowance description (or ENTER to finish): ")).strip()
        if not desc:
            break
        qty = prompt_float("Quantity")
        price = prompt_float("Estimated unit cost ($)")
        a = OwnerAllowance(description=desc, quantity=qty, unit_cost=price)
        allowances.append(a)
        print(f"  Added allowance: {desc} | {qty} @ ${price:.2f} = ${a.total:,.2f}\n")

    return allowances


# =========================
#  HIGH-LEVEL FLOWS
# =========================

def create_quote_or_invoice(doc_type: str) -> None:
    assert doc_type in ("quote", "invoice")

    print(colored_header(f"\n=== Create {doc_type.capitalize()} ==="))
    existing_job = input(colored_prompt("Enter existing Job ID (e.g. J0001) or press Enter for new: ")).strip().upper()

    project_name = input(colored_prompt("Project name: ")).strip()
    client_name = input(colored_prompt("Client name: ")).strip()
    project_address = input(colored_prompt("Project address (street, city, etc.): ")).strip()

    if existing_job:
        job_id = existing_job
        matching_dirs = find_existing_job_folders(job_id)

        if matching_dirs:
            if len(matching_dirs) == 1:
                job_dir = matching_dirs[0]
                print(colored_section(f"\nUsing existing job folder: {os.path.basename(job_dir)}"))
            else:
                print(colored_section("\nMultiple folders found for this Job ID:"))
                for idx, d in enumerate(matching_dirs, start=1):
                    print(f"  {idx}. {os.path.basename(d)}")
                print("  0. Create a new folder for this Job ID")

                while True:
                    choice = input(colored_prompt("Select folder (0–{n}): ".format(n=len(matching_dirs)))).strip()
                    try:
                        c = int(choice)
                    except ValueError:
                        print("Please enter a valid number.")
                        continue
                    if 0 <= c <= len(matching_dirs):
                        break
                    print("Out of range, try again.")

                if c == 0:
                    job_dir = get_job_folder(job_id, project_name)
                    print(colored_section(f"\nCreated new job folder: {os.path.basename(job_dir)}"))
                else:
                    job_dir = matching_dirs[c - 1]
                    print(colored_section(f"\nUsing existing job folder: {os.path.basename(job_dir)}"))
        else:
            job_dir = get_job_folder(job_id, project_name)
            print(colored_section(f"\nNo existing folder found for {job_id}. Created new folder: {os.path.basename(job_dir)}"))

        show_job_folder_contents(job_dir)

    else:
        job_id = get_next_job_id()
        job_dir = get_job_folder(job_id, project_name)

        if os.listdir(job_dir):
            print(colored_section(f"\n⚠️ Warning: job folder {os.path.basename(job_dir)} already existed and is not empty."))
            use_existing = prompt_yes_no("Use this existing folder for the new document?", default=True)
            if not use_existing:
                while True:
                    job_id = get_next_job_id()
                    job_dir = get_job_folder(job_id, project_name)
                    if not os.listdir(job_dir):
                        break
                print(colored_section(f"Using new empty job folder: {os.path.basename(job_dir)}"))
        else:
            print(colored_section(f"\nAssigned new Job ID: {job_id} and created folder {os.path.basename(job_dir)}"))

        show_job_folder_contents(job_dir)

    pricing_mode_quote = None
    billing_mode_invoice = None

    if doc_type == "quote":
        print(colored_section("\nPricing mode for this quote:"))
        print("  1. Standard (line totals roll up to grand total)")
        print("  2. Fixed price (manual grand total, blank line pricing)")
        while True:
            pm_choice = input(colored_prompt("Choice (1–2): ")).strip()
            if pm_choice in ("1", "2"):
                pricing_mode_quote = "standard" if pm_choice == "1" else "fixed"
                break
            print("Invalid choice. Please enter 1 or 2.")
    else:
        print(colored_section("\nBilling method for this invoice:"))
        print("  1. Time & Materials (labor with rates, materials)")
        print("  2. Fixed price (scope-only lines, blank prices)")
        while True:
            bm_choice = input(colored_prompt("Choice (1–2): ")).strip()
            if bm_choice in ("1", "2"):
                billing_mode_invoice = "tm" if bm_choice == "1" else "fixed"
                break
            print("Invalid choice. Please enter 1 or 2.")

    hide_line_prices = False

    if doc_type == "quote":
        if pricing_mode_quote == "standard":
            lines = build_quote_lines()
        else:
            lines = build_scope_only_lines("quote")
            hide_line_prices = True
    else:
        if billing_mode_invoice == "tm":
            lines = build_invoice_lines()
        else:
            lines = build_scope_only_lines("invoice")
            hide_line_prices = True

    if not lines:
        print("No line items added. Cancelling.")
        return


    # After labor/line items are entered for an invoice, optionally log receipts in a loop.
    # This runs before selecting receipts to include, so the new entries are available immediately.
    if doc_type == "invoice":
        log_receipts_loop_for_job(job_dir, job_id)

    receipt_rows: List[dict] = []
    receipts_total = 0.0

    if doc_type == "invoice":
        receipt_rows, receipts_total = choose_receipts_for_invoice(job_dir, job_id)

    allowances = build_owner_allowances()
    doc_date = datetime.date.today()

    if doc_type == "quote" and pricing_mode_quote == "fixed":
        fixed_total = prompt_float("Enter fixed grand total for this quote ($)")
        totals = {
            "subtotal": fixed_total,
            "receipts_total": 0.0,
            "grand_total": fixed_total,
        }
    elif doc_type == "invoice" and billing_mode_invoice == "fixed":
        manual_subtotal = prompt_float(
            "Enter fixed subtotal for this invoice (before receipts) ($)"
        )
        grand_total = manual_subtotal + receipts_total
        totals = {
            "subtotal": manual_subtotal,
            "receipts_total": receipts_total,
            "grand_total": grand_total,
        }
    else:
        totals = compute_totals(lines, receipts_total=receipts_total)

    csv_path = save_document_csv(
        job_dir=job_dir,
        job_id=job_id,
        doc_type=doc_type,
        project_name=project_name,
        project_address=project_address,
        client_name=client_name,
        lines=lines,
        totals=totals,
        allowances=allowances,
        doc_date=doc_date,
    )

    html_str = render_estimate_html(
        job_id=job_id,
        doc_type=doc_type,
        project_name=project_name,
        project_address=project_address,
        client_name=client_name,
        lines=lines,
        totals=totals,
        allowances=allowances,
        doc_date=doc_date,
        receipt_rows=receipt_rows,
        hide_line_prices=hide_line_prices,
    )
    html_path = save_estimate_html(
        job_dir=job_dir,
        job_id=job_id,
        doc_type=doc_type,
        project_name=project_name,
        html=html_str,
        doc_date=doc_date,
    )

    print(f"\nSaved {doc_type} CSV:   {csv_path}")
    print(f"Saved {doc_type} HTML:  {html_path}")
    print("Open the HTML file in your browser and use 'Print → Save as PDF' for client copies.\n")


def edit_existing_quote() -> None:
    print(colored_header("\n=== Edit Existing Quote ==="))
    job_id = input(colored_prompt("Enter Job ID for the quote (e.g. J0001): ")).strip().upper()
    if not job_id:
        print("Job ID is required.")
        return

    matching_dirs = find_existing_job_folders(job_id)
    if not matching_dirs:
        print(f"No job folder found for {job_id}.")
        return

    if len(matching_dirs) == 1:
        job_dir = matching_dirs[0]
        print(colored_section(f"Using job folder: {os.path.basename(job_dir)}"))
    else:
        print(colored_section("\nMultiple folders found for this Job ID:"))
        for idx, d in enumerate(matching_dirs, start=1):
            print(f"  {idx}. {os.path.basename(d)}")
        while True:
            raw = input(colored_prompt(f"Select folder (1–{len(matching_dirs)}): ")).strip()
            try:
                i = int(raw)
            except ValueError:
                print("Please enter a number.")
                continue
            if 1 <= i <= len(matching_dirs):
                job_dir = matching_dirs[i - 1]
                break
            print("Out of range, try again.")

    quote_files = [
        f for f in os.listdir(job_dir)
        if os.path.isfile(os.path.join(job_dir, f))
        and f.lower().endswith(".csv")
        and "_quote_" in f.lower()
    ]

    if not quote_files:
        print("No quote CSV files found in this job folder.")
        return

    quote_files = sorted(quote_files)
    print(colored_section("\nQuote files:"))
    for idx, fname in enumerate(quote_files, start=1):
        print(f"  {idx}. {fname}")

    while True:
        raw = input(colored_prompt(f"Select quote to edit (1–{len(quote_files)}): ")).strip()
        try:
            i = int(raw)
        except ValueError:
            print("Please enter a number.")
            continue
        if 1 <= i <= len(quote_files):
            csv_name = quote_files[i - 1]
            break
        print("Out of range, try again.")

    csv_path = os.path.join(job_dir, csv_name)
    print(colored_section(f"\nLoading quote: {csv_name}"))

    # NEW: backup original CSV + matching HTML before modifying anything
    backup_original_quote_files(csv_path)

    meta, lines, totals, allowances = load_quote_from_csv(csv_path)
    if not meta:
        print("Could not load quote file.")
        return

    print(colored_section("\nCurrent header values:"))
    print(f"  Project:  {meta['project_name']}")
    print(f"  Client:   {meta['client_name']}")
    print(f"  Address:  {meta['project_address']}")
    print(f"  Date:     {format_date(meta['doc_date'])}")
    print(f"  Subtotal: ${totals.get('subtotal', 0.0):,.2f}")
    print(f"  Receipts: ${totals.get('receipts_total', 0.0):,.2f}")
    print(f"  Grand:    ${totals.get('grand_total', 0.0):,.2f}")

    if not prompt_yes_no("Edit these values?", default=True):
        print("No changes made.")
        return

    meta, totals = edit_quote_headers_interactive(meta, totals)

    hide_line_prices = all((l.hours == 0 and l.rate == 0) for l in lines)

    job_id = meta.get("job_id") or job_id
    project_name = meta["project_name"]
    project_address = meta["project_address"]
    client_name = meta["client_name"]
    doc_date = meta["doc_date"]

    new_csv_path = save_document_csv(
        job_dir=job_dir,
        job_id=job_id,
        doc_type="quote",
        project_name=project_name,
        project_address=project_address,
        client_name=client_name,
        lines=lines,
        totals=totals,
        allowances=allowances,
        doc_date=doc_date,
    )

    html_str = render_estimate_html(
        job_id=job_id,
        doc_type="quote",
        project_name=project_name,
        project_address=project_address,
        client_name=client_name,
        lines=lines,
        totals=totals,
        allowances=allowances,
        doc_date=doc_date,
        receipt_rows=None,
        hide_line_prices=hide_line_prices,
    )
    new_html_path = save_estimate_html(
        job_dir=job_dir,
        job_id=job_id,
        doc_type="quote",
        project_name=project_name,
        html=html_str,
        doc_date=doc_date,
    )

    print(colored_section("\nHow would you like to name the updated files?"))
    print("  1. Mark as REVISED (adds _revised before extension)")
    print("  2. Mark as NEW (adds _new before extension)")
    print("  3. Keep default names")
    choice = input(colored_prompt("Choice (1–3): ")).strip()

    def _add_suffix(path: str, suffix: str) -> str:
        base, ext = os.path.splitext(path)
        new_path = base + suffix + ext
        try:
            os.replace(path, new_path)
            return new_path
        except OSError:
            return path

    if choice == "1":
        new_csv_path = _add_suffix(new_csv_path, "_revised")
        new_html_path = _add_suffix(new_html_path, "_revised")
    elif choice == "2":
        new_csv_path = _add_suffix(new_csv_path, "_new")
        new_html_path = _add_suffix(new_html_path, "_new")

    print(f"\nSaved updated quote CSV:  {new_csv_path}")
    print(f"Saved updated quote HTML: {new_html_path}")
    print("Open the new HTML file in your browser and use 'Print → Save as PDF' for client copies.\n")



def log_receipts_loop_for_job(job_dir: str, job_id: str) -> None:
    """Prompt the user to enter receipts in a loop (used during invoice creation).

    This writes to the job's receipts CSV (creating it if needed) and reorganizes it
    after entry. The user can skip by answering 'n' to the first prompt.
    """
    print(colored_header("\n=== Receipts (Optional) ==="))
    if not prompt_yes_no("Do you want to enter receipts now?", default=False):
        return

    receipts_path = choose_receipt_file_for_logging(job_dir, job_id)
    print(colored_section(f"\nLogging receipts to: {receipts_path}"))

    while True:
        if not prompt_yes_no("Add a receipt?", default=True):
            break

        item_name = input(colored_prompt("Item / vendor / brief description: ")).strip()
        if not item_name:
            print("Description cannot be blank.")
            continue

        date_str_input = input(colored_prompt("Receipt date (MM-DD-YYYY, blank for today): ")).strip()
        if date_str_input:
            try:
                _ = parse_mmddyyyy(date_str_input)
                date_str = date_str_input
            except ValueError:
                print("Invalid date format. Using today instead.")
                date_str = format_date(datetime.date.today())
        else:
            date_str = format_date(datetime.date.today())

        cost = prompt_float("Receipt cost ($)")
        csv_path = append_receipt_csv(receipts_path, job_id, item_name, date_str, cost)
        print(f"  ✓ Receipt logged to {csv_path}\n")

    reorganize_receipts_file(receipts_path)
    print("Receipts file reorganized by date and item.\n")


def log_receipt() -> None:
    print(colored_header("\n=== Log Receipt ==="))
    job_id = input(colored_prompt("Enter Job ID (e.g. J0001): ")).strip().upper()
    if not job_id:
        print("Job ID is required to log receipts.")
        return

    job_dir = get_job_folder(job_id, project_name=None)
    receipts_path = choose_receipt_file_for_logging(job_dir, job_id)
    print(colored_section(f"\nLogging receipts to: {receipts_path}"))

    print("\nEnter receipts for this job.")
    print("Press ENTER on Item description when you are finished.\n")

    while True:
        item_name = input(colored_prompt("Item / vendor / brief description (or ENTER to finish): ")).strip()
        if not item_name:
            break

        date_str_input = input(colored_prompt("Receipt date (MM-DD-YYYY, blank for today): ")).strip()
        if date_str_input:
            try:
                _ = parse_mmddyyyy(date_str_input)
                date_str = date_str_input
            except ValueError:
                print("Invalid date format. Using today instead.")
                date_str = format_date(datetime.date.today())
        else:
            date_str = format_date(datetime.date.today())

        cost = prompt_float("Receipt cost ($)")
        csv_path = append_receipt_csv(receipts_path, job_id, item_name, date_str, cost)
        print(f"  ✓ Receipt logged to {csv_path}\n")

    reorganize_receipts_file(receipts_path)
    print("Receipts file reorganized by date and item.\n")
    print("Done logging receipts for this job.\n")


# =========================
#  MAIN MENU
# =========================

def main():
    ensure_root_dir()
    print(colored_header("=== Cienega Construction Tool ==="))
    print(f"Root jobs folder: {ROOT_BASE_DIR}\n")

    while True:
        print(colored_section("Select an option:"))
        print("  1. Create Quote")
        print("  2. Create Invoice")
        print("  3. Log Receipt")
        print("  4. Edit Existing Quote")
        print("  5. Exit")

        choice = input(colored_prompt("Choice (1–5): ")).strip()
        if choice == "1":
            create_quote_or_invoice("quote")
        elif choice == "2":
            create_quote_or_invoice("invoice")
        elif choice == "3":
            log_receipt()
        elif choice == "4":
            edit_existing_quote()
        elif choice == "5":
            print("Goodbye.")
            break
        else:
            print("Invalid choice. Please enter 1, 2, 3, 4, or 5.\n")


if __name__ == "__main__":
    main()
