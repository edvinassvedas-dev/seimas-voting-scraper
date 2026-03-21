import requests
import xml.etree.ElementTree as ET
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import FreeSimpleGUI as sg
import io

# ── Google Sheets setup ────────────────────────────────────────────────────────

CREDENTIALS_PATH = 'credentials.json' #Google service account credentials
SPREADSHEET_NAME = 'your_spreadsheet' #Google spreadsheet file name

scope = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

creds  = Credentials.from_service_account_file(CREDENTIALS_PATH, scopes=scope)
gc     = gspread.authorize(creds)


def delete_voting_from_sheets(voting_id):
    """Delete all rows matching voting_id using read-filter-rewrite to avoid
    per-row API calls that trigger quota limits (429 errors)."""
    sheet1, sheet2, sheet3 = get_sheets()
    vid = str(voting_id)

    for sheet in (sheet1, sheet2, sheet3):
        all_rows = sheet.get_all_values()
        if not all_rows:
            continue
        header  = all_rows[0]
        kept    = [header] + [r for r in all_rows[1:] if r and str(r[0]) != vid]
        # Clear and rewrite in two API calls regardless of row count
        sheet.clear()
        if len(kept) > 1:
            sheet.update(values=kept, range_name="A1")


def get_sheets():
    """Return (sheet1, sheet2, sheet3) or raise on error."""
    wb = gc.open(SPREADSHEET_NAME)
    return wb.sheet1, wb.get_worksheet(1), wb.get_worksheet(2)


# ── Data fetching ──────────────────────────────────────────────────────────────

def fetch_voting(voting_id):
    """Fetch voting results for a given ID from the LRS API.
    Returns a DataFrame or an empty DataFrame if no records found."""
    url = (f"https://apps.lrs.lt/sip/p2b.ad_sp_balsavimo_rezultatai"
           f"?balsavimo_id=-{voting_id}")
    response = requests.get(url, timeout=15)
    response.raise_for_status()
    root = ET.fromstring(response.content)

    individual_votes = root.findall(".//IndividualusBalsavimoRezultatas")
    if not individual_votes:
        return pd.DataFrame()

    attr_names = {attr for vote in individual_votes for attr in vote.attrib.keys()}
    data       = {attr: [v.attrib.get(attr, "") for v in individual_votes]
                  for attr in attr_names}
    df = pd.DataFrame(data)
    df['voting'] = voting_id

    general = root.findall(".//BendriBalsavimoRezultatai")
    voting_time = general[0].attrib.get("balsavimo_laikas", "") if general else ""
    df['voting_time'] = pd.to_datetime(voting_time, errors='coerce')
    df['date']        = df['voting_time'].dt.date.astype(str)

    df['member']   = df.apply(lambda r: f"{r.get('vardas','')} {r.get('pavardė','')}".strip(), axis=1)
    df['fraction'] = df['frakcija']
    df['result']   = df['kaip_balsavo']

    # Try to auto-populate voting name from XML
    auto_name = ""
    if general:
        auto_name = general[0].attrib.get("balsavimo_pavadinimas", "")

    return df[['voting', 'date', 'member', 'fraction', 'result']], auto_name


# ── Database helpers ───────────────────────────────────────────────────────────

def get_existing_voting_ids():
    """Return a set of voting IDs already in sheet1."""
    try:
        sheet1, _, _ = get_sheets()
        col = sheet1.col_values(1)  # first column = 'voting'
        return set(str(v) for v in col[1:] if v)  # skip header
    except Exception:
        return set()


def check_duplicate(voting_id):
    """Return True if voting_id already exists in the database."""
    return str(voting_id) in get_existing_voting_ids()


def append_to_sheets(voting_df, votingname_df, votingurl_df):
    """Append data to all three sheets."""
    sheet1, sheet2, sheet3 = get_sheets()
    sheet1.append_rows(voting_df.values.tolist())
    sheet2.append_rows(votingname_df.values.tolist())
    sheet3.append_rows(votingurl_df.values.tolist())


def load_db_summary():
    """Return a list of [voting_id, record_count, date] rows for the summary table."""
    try:
        sheet1, sheet2, _ = get_sheets()
        data = sheet1.get_all_values()
        if len(data) < 2:
            return []
        df = pd.DataFrame(data[1:], columns=data[0])

        # Names from sheet2
        names_data = sheet2.get_all_values()
        names = {}
        if len(names_data) >= 2:
            for row in names_data[1:]:
                if len(row) >= 2:
                    names[str(row[0])] = row[1]

        summary = (df.groupby('voting')
                     .agg(records=('voting', 'count'), date=('date', 'first'))
                     .reset_index()
                     .sort_values('voting'))

        return [
            [row['voting'], names.get(str(row['voting']), ""), row['records'], row['date']]
            for _, row in summary.iterrows()
        ]
    except Exception as e:
        sg.popup_error(f"Error loading summary: {e}")
        return []


# ── GUI ────────────────────────────────────────────────────────────────────────

sg.theme("Reddit")

layout = [
    # ── Input section ──
    [sg.Text("Voting ID:",   size=(14, 1)), sg.InputText(key="-VOTING_ID-",   size=(8,  1))],
    [sg.Text("Voting Name:", size=(14, 1)), sg.InputText(key="-VOTING_NAME-", size=(70, 1))],
    [sg.Text("Voting URL:",  size=(14, 1)), sg.InputText(key="-VOTING_URL-",  size=(70, 1))],
    [
        sg.Button("Get Data"),
        sg.Button("Insert into DB"),
        sg.Button("Copy to Clipboard"),
        sg.Button("Exit"),
    ],
    [sg.HorizontalSeparator()],

    # ── Results table ──
    [sg.Text("Results", font=("Helvetica", 11, "bold"))],
    [sg.Table(
        values=[],
        headings=["Voting", "Date", "Member", "Fraction", "Result"],
        key="-RESULTS_TABLE-",
        auto_size_columns=False,
        col_widths=[6, 6, 20, 10, 10],
        display_row_numbers=True,
        num_rows=20,
        expand_x=True,
    )],
    [sg.HorizontalSeparator()],

    # ── Database summary ──
    [sg.Text("Database Summary", font=("Helvetica", 11, "bold")),
     sg.Button("Refresh Summary", key="-REFRESH_SUMMARY-"),
     sg.Button("Delete Selected", key="-DELETE_VOTING-", disabled=True,
               button_color=("white", "#C0392B"))],
    [sg.Table(
        values=[],
        headings=["Voting ID", "Voting Name", "Records", "Date"],
        key="-SUMMARY_TABLE-",
        auto_size_columns=False,
        col_widths=[6, 50, 5, 10],
        display_row_numbers=False,
        num_rows=15,
        expand_x=True,
        enable_events=True,
    )],
]

window = sg.Window("Seimas Voting App", layout, size=(900, 750), resizable=True, finalize=True)

# Initialise state
result_df = pd.DataFrame()
voting_id = ""

# Load summary on startup
window["-SUMMARY_TABLE-"].update(values=load_db_summary())

# ── Event loop ─────────────────────────────────────────────────────────────────

while True:
    event, values = window.read()

    if event in (sg.WIN_CLOSED, "Exit"):
        break

    # ── Get Data ──
    if event == "Get Data":
        voting_id = values["-VOTING_ID-"].strip()
        if not voting_id:
            sg.popup_error("Please enter a Voting ID.")
            continue
        try:
            fetch_result = fetch_voting(voting_id)
            if isinstance(fetch_result, tuple):
                result_df, auto_name = fetch_result
            else:
                result_df, auto_name = fetch_result, ""

            if result_df.empty:
                sg.popup(f"No records found for voting ID {voting_id}.")
                window["-RESULTS_TABLE-"].update(values=[])
            else:
                window["-RESULTS_TABLE-"].update(values=result_df.values.tolist())
                if auto_name and not values["-VOTING_NAME-"].strip():
                    window["-VOTING_NAME-"].update(auto_name)
        except Exception as e:
            sg.popup_error(f"Error fetching data: {e}")

    # ── Insert into DB ──
    elif event == "Insert into DB":
        if result_df.empty:
            sg.popup_error("Please fetch data before inserting.")
            continue

        if check_duplicate(voting_id):
            confirm = sg.popup_yes_no(
                f"Voting ID {voting_id} already exists in the database.\n"
                "Insert anyway?",
                title="Duplicate Warning"
            )
            if confirm != "Yes":
                continue

        voting_name = values["-VOTING_NAME-"].strip()
        voting_url  = values["-VOTING_URL-"].strip()

        votingname_df = pd.DataFrame({'voting': [voting_id], 'voting_name': [voting_name]})
        votingurl_df  = pd.DataFrame({'voting': [voting_id], 'voting_url':  [voting_url]})

        try:
            append_to_sheets(result_df, votingname_df, votingurl_df)
            sg.popup("Inserted successfully.")
            # Refresh summary
            window["-SUMMARY_TABLE-"].update(values=load_db_summary())
            # Clear inputs
            window["-VOTING_ID-"].update("")
            window["-VOTING_NAME-"].update("")
            window["-VOTING_URL-"].update("")
            window["-RESULTS_TABLE-"].update(values=[])
            result_df = pd.DataFrame()
        except Exception as e:
            sg.popup_error(f"Insertion failed: {e}")

    # ── Copy to Clipboard ──
    elif event == "Copy to Clipboard":
        if result_df.empty:
            sg.popup_error("No data to copy.")
            continue
        buf = io.StringIO()
        result_df.to_csv(buf, sep='\t', index=False)
        sg.clipboard_set(buf.getvalue())
        sg.popup("Copied to clipboard.")

    # ── Summary table selection ──
    elif event == "-SUMMARY_TABLE-":
        selected = values["-SUMMARY_TABLE-"]
        window["-DELETE_VOTING-"].update(disabled=not selected)

    # ── Delete selected voting ──
    elif event == "-DELETE_VOTING-":
        selected_rows = values["-SUMMARY_TABLE-"]
        if not selected_rows:
            continue
        summary_data = load_db_summary()
        sel_row = summary_data[selected_rows[0]]
        sel_vid = sel_row[0]
        sel_name = sel_row[1] or str(sel_vid)
        msg = f"Delete all records for voting ID {sel_vid} ({sel_name})? This will remove data from all three sheets."
        confirm = sg.popup_yes_no(msg, title="Confirm Delete")
        if confirm == "Yes":
            try:
                delete_voting_from_sheets(sel_vid)
                sg.popup(f"Voting {sel_vid} deleted.")
                window["-SUMMARY_TABLE-"].update(values=load_db_summary())
                window["-DELETE_VOTING-"].update(disabled=True)
            except Exception as e:
                sg.popup_error(f"Delete failed: {e}")

    # ── Refresh Summary ──
    elif event == "-REFRESH_SUMMARY-":
        window["-SUMMARY_TABLE-"].update(values=load_db_summary())
        window["-DELETE_VOTING-"].update(disabled=True)

window.close()
