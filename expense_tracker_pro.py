# expense_tracker_pro.py
# ------------------------------------------------------------
# Beautiful, persistent Expense Tracker (Console)
# - Saves to expenses.csv so your data survives closing.
# - Add expenses with date, item, category, amount.
# - View totals (overall, today, this month, by category).
# - Filter by date range.
# - Toggle pretty color themes (with curated vibes + preview).
# - CSV = easy to open in Excel/Sheets.
#
# Launch (recommended): money  (via money.bat or ./money launcher)
# If running directly: python expense_tracker_pro.py
# ------------------------------------------------------------

import csv
import os
from datetime import datetime, date

# ---------- Color setup with graceful fallback ----------
# If colorama isn't installed, we create "dummy" color constants.
USE_COLOR = True
try:
    from colorama import init, Fore, Style
    init(autoreset=True)
except Exception:
    USE_COLOR = False

    class _Dummy:
        def __getattr__(self, _):  # any attr => empty string
            return ""
    Fore = Style = _Dummy()
    
def _c(s: str) -> str:
    """Return color string or empty if colors disabled."""
    return s if USE_COLOR else ""

# ---------- Files ----------
DATA_FILE = "expenses.csv"
CONFIG_FILE = "expense_config.txt"  # stores the last selected theme name

# ---------- Themes ----------
# - title: banner/headings
# - menu: menu/options/table headers
# - prompt: input text + neutral text
# - accent: numbers/money/highlights
# - ok: success
# - warn: informational warnings
# - err: errors
# - muted: dividers/hints
THEMES = {
    "dark": {
        "title": _c(Fore.CYAN + Style.BRIGHT),
        "menu": _c(Fore.WHITE + Style.BRIGHT),
        "prompt": _c(Fore.LIGHTWHITE_EX),
        "accent": _c(Fore.MAGENTA + Style.BRIGHT),
        "ok": _c(Fore.GREEN + Style.BRIGHT),
        "warn": _c(Fore.YELLOW + Style.BRIGHT),
        "err": _c(Fore.RED + Style.BRIGHT),
        "muted": _c(Fore.LIGHTBLACK_EX),
    },
    "sunset": {
        "title": _c(Fore.LIGHTRED_EX + Style.BRIGHT),
        "menu": _c(Fore.LIGHTYELLOW_EX + Style.BRIGHT),
        "prompt": _c(Fore.LIGHTWHITE_EX),
        "accent": _c(Fore.LIGHTMAGENTA_EX + Style.BRIGHT),
        "ok": _c(Fore.LIGHTGREEN_EX + Style.BRIGHT),
        "warn": _c(Fore.LIGHTYELLOW_EX + Style.BRIGHT),
        "err": _c(Fore.LIGHTRED_EX + Style.BRIGHT),
        "muted": _c(Fore.LIGHTBLACK_EX),
    },
    "ocean": {
        "title": _c(Fore.LIGHTBLUE_EX + Style.BRIGHT),
        "menu": _c(Fore.LIGHTCYAN_EX + Style.BRIGHT),
        "prompt": _c(Fore.LIGHTWHITE_EX),
        "accent": _c(Fore.CYAN + Style.BRIGHT),
        "ok": _c(Fore.LIGHTGREEN_EX + Style.BRIGHT),
        "warn": _c(Fore.YELLOW + Style.BRIGHT),
        "err": _c(Fore.LIGHTRED_EX + Style.BRIGHT),
        "muted": _c(Fore.LIGHTBLACK_EX),
    },
    "neon": {
        "title": _c(Fore.LIGHTGREEN_EX + Style.BRIGHT),
        "menu": _c(Fore.LIGHTWHITE_EX + Style.BRIGHT),
        "prompt": _c(Fore.LIGHTWHITE_EX),
        "accent": _c(Fore.LIGHTMAGENTA_EX + Style.BRIGHT),
        "ok": _c(Fore.LIGHTGREEN_EX + Style.BRIGHT),
        "warn": _c(Fore.LIGHTYELLOW_EX + Style.BRIGHT),
        "err": _c(Fore.LIGHTRED_EX + Style.BRIGHT),
        "muted": _c(Fore.LIGHTBLACK_EX),
    },
    "forest": {
        "title": _c(Fore.GREEN + Style.BRIGHT),
        "menu": _c(Fore.LIGHTGREEN_EX + Style.BRIGHT),
        "prompt": _c(Fore.WHITE),
        "accent": _c(Fore.LIGHTCYAN_EX + Style.BRIGHT),
        "ok": _c(Fore.GREEN + Style.BRIGHT),
        "warn": _c(Fore.YELLOW + Style.BRIGHT),
        "err": _c(Fore.RED + Style.BRIGHT),
        "muted": _c(Fore.LIGHTBLACK_EX),
    },
    "retro": {
        "title": _c(Fore.CYAN + Style.BRIGHT),
        "menu": _c(Fore.LIGHTMAGENTA_EX + Style.BRIGHT),
        "prompt": _c(Fore.LIGHTWHITE_EX),
        "accent": _c(Fore.MAGENTA + Style.BRIGHT),
        "ok": _c(Fore.LIGHTGREEN_EX + Style.BRIGHT),
        "warn": _c(Fore.LIGHTYELLOW_EX + Style.BRIGHT),
        "err": _c(Fore.LIGHTRED_EX + Style.BRIGHT),
        "muted": _c(Fore.BLACK),
    },
    "mono": {
        "title": _c(Fore.WHITE + Style.BRIGHT),
        "menu": _c(Fore.LIGHTWHITE_EX + Style.BRIGHT),
        "prompt": _c(Fore.WHITE),
        "accent": _c(Fore.LIGHTWHITE_EX + Style.BRIGHT),
        "ok": _c(Fore.WHITE + Style.BRIGHT),
        "warn": _c(Fore.LIGHTBLACK_EX + Style.BRIGHT),
        "err": _c(Fore.LIGHTBLACK_EX + Style.BRIGHT),
        "muted": _c(Fore.LIGHTBLACK_EX),
    },
    "lavender": {
        "title": _c(Fore.LIGHTMAGENTA_EX + Style.BRIGHT),
        "menu": _c(Fore.MAGENTA + Style.BRIGHT),
        "prompt": _c(Fore.LIGHTWHITE_EX),
        "accent": _c(Fore.LIGHTBLUE_EX + Style.BRIGHT),
        "ok": _c(Fore.LIGHTGREEN_EX + Style.BRIGHT),
        "warn": _c(Fore.LIGHTYELLOW_EX + Style.BRIGHT),
        "err": _c(Fore.LIGHTRED_EX + Style.BRIGHT),
        "muted": _c(Fore.LIGHTBLACK_EX),
    },
    "gold": {
        "title": _c(Fore.YELLOW + Style.BRIGHT),
        "menu": _c(Fore.LIGHTYELLOW_EX + Style.BRIGHT),
        "prompt": _c(Fore.WHITE),
        "accent": _c(Fore.YELLOW + Style.BRIGHT),
        "ok": _c(Fore.LIGHTGREEN_EX + Style.BRIGHT),
        "warn": _c(Fore.YELLOW + Style.BRIGHT),
        "err": _c(Fore.RED + Style.BRIGHT),
        "muted": _c(Fore.LIGHTBLACK_EX),
    },
    "matrix": {
        "title": _c(Fore.GREEN + Style.BRIGHT),
        "menu": _c(Fore.LIGHTGREEN_EX + Style.BRIGHT),
        "prompt": _c(Fore.LIGHTWHITE_EX),
        "accent": _c(Fore.GREEN + Style.BRIGHT),
        "ok": _c(Fore.LIGHTGREEN_EX + Style.BRIGHT),
        "warn": _c(Fore.YELLOW + Style.BRIGHT),
        "err": _c(Fore.LIGHTRED_EX + Style.BRIGHT),
        "muted": _c(Fore.LIGHTBLACK_EX),
    },
    "ember": {
        "title": _c(Fore.RED + Style.BRIGHT),
        "menu": _c(Fore.LIGHTRED_EX + Style.BRIGHT),
        "prompt": _c(Fore.WHITE),
        "accent": _c(Fore.LIGHTYELLOW_EX + Style.BRIGHT),
        "ok": _c(Fore.LIGHTGREEN_EX + Style.BRIGHT),
        "warn": _c(Fore.YELLOW + Style.BRIGHT),
        "err": _c(Fore.LIGHTRED_EX + Style.BRIGHT),
        "muted": _c(Fore.LIGHTBLACK_EX),
    },
    "ice": {
        "title": _c(Fore.LIGHTCYAN_EX + Style.BRIGHT),
        "menu": _c(Fore.CYAN + Style.BRIGHT),
        "prompt": _c(Fore.LIGHTWHITE_EX),
        "accent": _c(Fore.LIGHTBLUE_EX + Style.BRIGHT),
        "ok": _c(Fore.LIGHTGREEN_EX + Style.BRIGHT),
        "warn": _c(Fore.YELLOW + Style.BRIGHT),
        "err": _c(Fore.LIGHTRED_EX + Style.BRIGHT),
        "muted": _c(Fore.LIGHTBLACK_EX),
    },
    "pastel": {
        "title": _c(Fore.LIGHTMAGENTA_EX + Style.BRIGHT),
        "menu": _c(Fore.LIGHTBLUE_EX + Style.BRIGHT),
        "prompt": _c(Fore.LIGHTWHITE_EX),
        "accent": _c(Fore.LIGHTCYAN_EX + Style.BRIGHT),
        "ok": _c(Fore.LIGHTGREEN_EX + Style.BRIGHT),
        "warn": _c(Fore.LIGHTYELLOW_EX + Style.BRIGHT),
        "err": _c(Fore.LIGHTRED_EX + Style.BRIGHT),
        "muted": _c(Fore.LIGHTBLACK_EX),
    },
}

DEFAULT_THEME = "dark"

def load_theme_name():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                name = f.read().strip()
                if name in THEMES:
                    return name
        except:
            pass
    return DEFAULT_THEME

CURRENT_THEME = THEMES[load_theme_name()]

def set_theme(name: str):
    """Persist and apply a theme by name."""
    global CURRENT_THEME
    if name in THEMES:
        CURRENT_THEME = THEMES[name]
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            f.write(name)
        echo(CURRENT_THEME["ok"], f"✓ Theme set to '{name}'.")
    else:
        echo(CURRENT_THEME["err"], f"✗ Unknown theme '{name}'. Available: {', '.join(THEMES.keys())}")

# ---------- Small UI helpers ----------
def line(char="─", width=60):
    return char * width

def echo(color, text=""):
    # color is already '' if colors disabled
    print(color + text + (Style.RESET_ALL if USE_COLOR else ""))

def banner():
    echo(CURRENT_THEME["title"], "\n" + line())
    echo(CURRENT_THEME["title"], "   💸 Expense Tracker — simple • persistent • pretty")
    echo(CURRENT_THEME["title"], line() + "\n")

def pause():
    input(CURRENT_THEME["muted"] + "Press Enter to continue...")

def preview_themes():
    """Visual preview so each theme feels like itself."""
    print()
    for name, t in THEMES.items():
        print(
            f"{t['title']}■{(Style.RESET_ALL if USE_COLOR else '')} "
            f"{t['menu']}{name:<10}{(Style.RESET_ALL if USE_COLOR else '')}  "
            f"{t['muted']}— {t['accent']}accent  "
            f"{t['ok']}ok  "
            f"{t['warn']}warn  "
            f"{t['err']}err{(Style.RESET_ALL if USE_COLOR else '')}"
        )
    print()

# ---------- CSV helpers ----------
def ensure_csv_exists():
    if not os.path.exists(DATA_FILE):
        with open(DATA_FILE, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["date", "item", "category", "amount"])  # header

def append_expense(row):
    with open(DATA_FILE, "a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(row)

def read_all():
    ensure_csv_exists()
    rows = []
    with open(DATA_FILE, "r", newline="", encoding="utf-8") as f:
        r = csv.DictReader(f)
        for rec in r:
            rows.append(rec)
    return rows

# ---------- Formatting ----------
def money(n):
    try:
        return f"${float(n):,.2f}"
    except:
        return "$0.00"

def parse_date(s):
    return datetime.strptime(s, "%Y-%m-%d").date()

def safe_date_input(prompt_text, default_date=None):
    while True:
        raw = input(CURRENT_THEME["prompt"] + prompt_text).strip()
        if raw == "" and default_date:
            return default_date
        try:
            return parse_date(raw)
        except ValueError:
            echo(CURRENT_THEME["err"], "✗ Use format YYYY-MM-DD (example: 2025-10-13).")

def safe_amount_input(prompt_text):
    while True:
        raw = input(CURRENT_THEME["prompt"] + prompt_text).strip().replace("$", "")
        try:
            val = float(raw)
            return round(val, 2)
        except ValueError:
            echo(CURRENT_THEME["err"], "✗ Enter a number (example: 12.50).")

# ---------- Features ----------
def add_expense():
    banner()
    echo(CURRENT_THEME["accent"], "➕ Add Expense")
    today_str = date.today().strftime("%Y-%m-%d")
    d = safe_date_input(f"Date [YYYY-MM-DD | Enter for {today_str}]: ", default_date=date.today())
    item = input(CURRENT_THEME["prompt"] + "Item (e.g., Chicken breast, Gas): ").strip() or "(unnamed)"
    category = input(CURRENT_THEME["prompt"] + "Category (e.g., Food, Gym, Bills): ").strip() or "Uncategorized"
    amount = safe_amount_input("Amount (e.g., 8.99): ")

    append_expense((d.strftime("%Y-%m-%d"), item, category, amount))
    echo(CURRENT_THEME["ok"], f"\n✓ Added: {item}  |  {category}  |  {money(amount)}  |  {d}")
    pause()

def show_all():
    banner()
    echo(CURRENT_THEME["accent"], "📜 All Expenses (latest first)")
    rows = read_all()
    if not rows:
        echo(CURRENT_THEME["warn"], "No records yet.")
        pause()
        return

    rows.sort(key=lambda r: r["date"], reverse=True)

    echo(CURRENT_THEME["muted"], line())
    print(CURRENT_THEME["menu"] + f"{'DATE':<12} {'ITEM':<24} {'CATEGORY':<14} {'AMOUNT':>10}")
    echo(CURRENT_THEME["muted"], line())
    total = 0.0
    for r in rows:
        amt = float(r["amount"])
        total += amt
        print(
            f"{CURRENT_THEME['prompt']}{r['date']:<12} "
            f"{CURRENT_THEME['prompt']}{r['item']:<24} "
            f"{CURRENT_THEME['prompt']}{r['category']:<14} "
            f"{CURRENT_THEME['accent']}{money(amt):>10}"
        )
    echo(CURRENT_THEME["muted"], line())
    echo(CURRENT_THEME["ok"], f"TOTAL: {money(total)}")
    pause()

def totals_menu():
    banner()
    rows = read_all()
    if not rows:
        echo(CURRENT_THEME["warn"], "No records yet.")
        pause()
        return

    today = date.today().strftime("%Y-%m-%d")
    ym_now = date.today().strftime("%Y-%m")

    overall = sum(float(r["amount"]) for r in rows)
    today_total = sum(float(r["amount"]) for r in rows if r["date"] == today)
    month_total = sum(float(r["amount"]) for r in rows if r["date"].startswith(ym_now + "-"))

    cat_totals = {}
    for r in rows:
        cat = r["category"] or "Uncategorized"
        cat_totals[cat] = cat_totals.get(cat, 0.0) + float(r["amount"])

    echo(CURRENT_THEME["accent"], "📊 Totals")
    echo(CURRENT_THEME["ok"], f"Overall: {money(overall)}")
    echo(CURRENT_THEME["ok"], f"Today ({today}): {money(today_total)}")
    echo(CURRENT_THEME["ok"], f"This month ({ym_now}): {money(month_total)}\n")

    echo(CURRENT_THEME["accent"], "🏷️  By Category (overall)")
    echo(CURRENT_THEME["muted"], line())
    print(CURRENT_THEME["menu"] + f"{'CATEGORY':<20} {'TOTAL':>12}")
    echo(CURRENT_THEME["muted"], line())
    for cat, amt in sorted(cat_totals.items(), key=lambda x: x[1], reverse=True):
        print(f"{CURRENT_THEME['prompt']}{cat:<20} {CURRENT_THEME['accent']}{money(amt):>12}")
    echo(CURRENT_THEME["muted"], line())
    pause()

def filter_by_date_range():
    banner()
    echo(CURRENT_THEME["accent"], "🔎 Filter by Date Range")
    start = safe_date_input("Start date [YYYY-MM-DD]: ")
    end = safe_date_input("End date   [YYYY-MM-DD]: ")
    if end < start:
        echo(CURRENT_THEME["err"], "✗ End date must be after start date.")
        pause()
        return

    rows = read_all()
    picked = []
    total = 0.0
    for r in rows:
        try:
            d = parse_date(r["date"])
        except:
            continue
        if start <= d <= end:
            picked.append(r)
            total += float(r["amount"])

    if not picked:
        echo(CURRENT_THEME["warn"], "No records in that range.")
        pause()
        return

    echo(CURRENT_THEME["muted"], line())
    print(CURRENT_THEME["menu"] + f"{'DATE':<12} {'ITEM':<24} {'CATEGORY':<14} {'AMOUNT':>10}")
    echo(CURRENT_THEME["muted"], line())
    for r in sorted(picked, key=lambda r: r["date"], reverse=True):
        print(
            f"{CURRENT_THEME['prompt']}{r['date']:<12} "
            f"{CURRENT_THEME['prompt']}{r['item']:<24} "
            f"{CURRENT_THEME['prompt']}{r['category']:<14} "
            f"{CURRENT_THEME['accent']}{money(r['amount']):>10}"
        )
    echo(CURRENT_THEME["muted"], line())
    echo(CURRENT_THEME["ok"], f"RANGE TOTAL: {money(total)}")
    pause()

def switch_theme():
    banner()
    echo(CURRENT_THEME["accent"], "🎨 Switch Theme")
    echo(CURRENT_THEME["menu"], "Preview:")
    preview_themes()
    echo(CURRENT_THEME["menu"], "Available themes:")
    print(", ".join(sorted(THEMES.keys())))
    choice = input(CURRENT_THEME["prompt"] + "Enter theme name: ").strip().lower()
    set_theme(choice)
    pause()

# ---------- Main loop ----------
def main():
    ensure_csv_exists()
    while True:
        banner()
        echo(CURRENT_THEME["menu"], "1) Add expense")
        echo(CURRENT_THEME["menu"], "2) Show all")
        echo(CURRENT_THEME["menu"], "3) Totals")
        echo(CURRENT_THEME["menu"], "4) Filter by date range")
        echo(CURRENT_THEME["menu"], "5) Switch theme")
        echo(CURRENT_THEME["menu"], "6) Exit")
        choice = input(CURRENT_THEME["prompt"] + "Choose: ").strip()

        if choice == "1":
            add_expense()
        elif choice == "2":
            show_all()
        elif choice == "3":
            totals_menu()
        elif choice == "4":
            filter_by_date_range()
        elif choice == "5":
            switch_theme()
        elif choice == "6":
            echo(CURRENT_THEME["ok"], "Bye! 🧾")
            break
        else:
            echo(CURRENT_THEME["err"], "✗ Invalid option.")
            pause()

if __name__ == "__main__":
    main()
