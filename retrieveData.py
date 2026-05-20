import pandas as pd
from selenium import webdriver
import hashlib
import csv
from datetime import datetime
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from utils import read_settings, loginToEdata, read_am_codes

import time


# Valid klados codes as they exist in the database
KLADOS_CODES = {
    'ΠΕ60', 'ΠΕ06', 'ΠΕ70', 'ΠΕ11', 'ΔΕ01-ΕΒΠ', 'ΠΕ79.01', 'ΠΕ25',
    'ΠΕ71', 'ΠΕ23', 'ΠΕ30', 'ΠΕ21', 'ΠΕ07', 'ΠΕ86', 'ΠΕ08', 'ΠΕ29',
    'ΠΕ91.01', 'ΠΕ28', 'ΠΕ05', 'ΠΕ91.02', 'ΠΕ70.50', 'ΠΕ79.01.50',
    'ΠΕ91.01.50', 'ΠΕ11.01', 'ΠΕ60.50', 'ΠΕ86.50', 'ΔΕ1ΕΒΠ', 'ΠΕ87.02',
    'ΠΕ79', 'ΤΕ16', 'ΠΕ30-ΣΔΕΥ', 'ΠΕ08.50', 'ΠΕ61',
}

def navigate_to_teacher(driver, am_code):
    """
    Navigates to the teacher profile page for a given AM code.

    Args:
        driver: Selenium WebDriver instance
        am_code: 6-digit AM code of the teacher

    Returns:
        True if navigation succeeded, False otherwise
    """
    try:
        print(f"Searching for AM code: {am_code}")
        wait = WebDriverWait(driver, 10)

        # Step 1: Click "Εργαζόμενοι" to expand the menu section
        ergazomenoi_link = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'Εργαζόμενοι')]"))
            # ^^^ TODO: If this doesn't work, try: By.XPATH, "//a[@href='YOUR_HREF_HERE']"
        )
        ergazomenoi_link.click()
        time.sleep(1)

        # Step 2: Click "Βασικά Στοιχεία"
        vasika_link = wait.until(
            EC.element_to_be_clickable((By.ID, "ctl00_ContentMain_hyplnkWorker"))
        )
        vasika_link.click()
        time.sleep(2)

        # Enter the AM code in the search field
        am_input = wait.until(
            EC.presence_of_element_located((By.XPATH, "//span[@id='ctl00_ContentMain_dxpanelCriteria_lblRegistryNo']/../..//input[@type='text']"))
            # ^^^ TODO: Confirm this XPATH matches the AM input on the teacher search page
        )
        am_input.clear()
        am_input.send_keys(am_code)

        # Click the search button
        search_button = driver.find_element(By.ID, "ctl00_ContentMain_dxpanelCriteria_btnSearch")
        search_button.click()
        time.sleep(2)

        # Click the first result row to open the teacher profile
        result_link = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//tr[@id='ctl00_ContentMain_dxgridResults_DXDataRow0']//a"))
        )
        result_link.click()
        time.sleep(2)

        return True

    except Exception as e:
        print(f"Error navigating to teacher {am_code}: {e}")
        return False


def retrieve_teacher_data(driver, am_code):
    """
    Retrieves teacher information from their profile page.

    Args:
        driver: Selenium WebDriver instance
        am_code: the AM code being processed (used for logging)

    Returns:
        dict with teacher data, or None on failure
    """
    wait = WebDriverWait(driver, 10)
    data = {}

    data['am'] = am_code
    data['surname'] = driver.find_element(By.ID, "ctl00_ContentMain_dxpanelGeneral_txtLastName").get_attribute("value")
    data['name'] = driver.find_element(By.ID, "ctl00_ContentMain_dxpanelGeneral_txtName").get_attribute("value")
    data['afm'] = driver.find_element(By.ID, "ctl00_ContentMain_dxpanelGeneral_txtTaxNo").get_attribute("value")
    data['specialty'] = driver.find_element(By.ID, "ctl00_ContentMain_dxtabWorker_edclSectionSpeciality_cpLookup_ddeLookup_I").get_attribute("value")
    data['years'] = driver.find_element(By.ID, "ctl00_ContentMain_dxtabWorker_txtTotalEmploymentYears").get_attribute("value")
    data['months'] = driver.find_element(By.ID, "ctl00_ContentMain_dxtabWorker_txtTotalEmploymentMonths").get_attribute("value")
    data['days'] = driver.find_element(By.ID, "ctl00_ContentMain_dxtabWorker_txtTotalEmploymentDays").get_attribute("value")
    # Click the Επικοινωνία tab
    epikoinonia_tab = driver.find_element(By.ID, "ctl00_ContentMain_dxtabWorker_AT3T")
    driver.execute_script("arguments[0].click();", epikoinonia_tab)
    time.sleep(1)

    data['mobile'] = driver.find_element(By.ID, "ctl00_ContentMain_dxtabWorker_igmaskMobilePhone").get_attribute("value")
    data['email'] = driver.find_element(By.ID, "ctl00_ContentMain_dxtabWorker_dxtxtEMail_I").get_attribute("value")
    # Click the Υπηρετήσεις button
    ypiretiseis_btn = driver.find_element(By.XPATH, "//img[@alt='Υπηρετήσεις']")
    driver.execute_script("arguments[0].click();", ypiretiseis_btn)
    time.sleep(2)

    # Get all data rows from the table
    rows = driver.find_elements(By.XPATH, "//tr[contains(@id, 'ctl00_ContentMain_dxgridResults_DXDataRow')]")

    if rows:
        last_row = rows[-2] if '--- ΣΥΝΟΛΟ ΥΠΗΡΕΤΗΣΕΩΝ ---' in rows[-1].text else rows[-1]
        cells = last_row.find_elements(By.TAG_NAME, "td")
        data['service_school'] = cells[1].text.strip()   # Μονάδα Υπηρέτησης
        data['education_type'] = cells[3].text.strip()   # Τύπος Οργανικής Θέσης
    else:
        print("No service rows found")
        data['service_school'] = ''
        data['education_type'] = ''

    # print(data)
    return data

def load_from_xlsx(xlsx_file):
    """
    Loads teacher records from an Excel file into a list of dicts
    that generate_sql can consume.

    Expected columns (same names written by the scraper):
        am, surname, name, afm, specialty, years, months, days,
        mobile, email, service_school, education_type

    Args:
        xlsx_file: path to the Excel file

    Returns:
        list of dicts, one per teacher row
    """
    df = pd.read_excel(xlsx_file, dtype=str)
    df.fillna('', inplace=True)

    required_columns = {'am', 'surname', 'name', 'afm', 'specialty',
                        'mobile', 'email', 'education_type'}
    missing = required_columns - set(df.columns)
    if missing:
        raise ValueError(f"Excel file is missing required columns: {missing}")

    return df.to_dict(orient='records')


def load_school_mapping(schools_xlsx, db_csv):
    """
    Builds a mapping from school name → database id (organiki_id).

    Step 1 — schools_xlsx  (the Σχολικές Μονάδες file):
        Reads every row that has a numeric value in column B (index 1) and a
        non-empty string in column C (index 2).  Those two columns give us:
            school_name → school_code

    Step 2 — db_csv (the database export):
        Parses the CSV (single-quoted values) to produce:
            school_code → school_id

    Combining both steps yields:
            school_name → school_id   (used as organiki_id)

    Args:
        schools_xlsx : path to the Σχολικές Μονάδες Excel file
        db_csv       : path to the database schools CSV export

    Returns:
        dict  { school_name (str) : school_id (int) }
    """
    # ── Step 1: name → code from the schools xlsx ──────────────────────────
    # The file has 3 preamble rows (empty / title / empty) then the real
    # column-name row at index 3: Εποπτεύων Φορέας | Κωδικός | Όνομα | …
    # header=3 tells pandas to use that row as column names directly.
    df_schools = pd.read_excel(schools_xlsx, header=3, dtype=str)
    df_schools.fillna('', inplace=True)

    name_to_code = {}
    for _, row in df_schools.iterrows():
        code = row['Κωδικός'].strip()
        name = row['Όνομα'].strip()
        # Keep only rows where Κωδικός is a numeric school code;
        # this skips regional-authority headers and blank rows automatically.
        if code.isdigit() and name:
            name_to_code[name] = code

    print(f"  Loaded {len(name_to_code)} school names from {schools_xlsx}")

    # ── Step 2: code → id from the database CSV ─────────────────────────────
    # Use csv.DictReader with quotechar="'" so commas inside quoted address
    # fields (e.g. 'ΑΓΙΟΥ ΛΕΟΝΤΙΟΥ, 28') don't break field counting.
    
    code_to_id = {}
    with open(db_csv, newline='', encoding='utf-8') as f:
        reader = csv.DictReader(f, quotechar="'", skipinitialspace=True)
        for row in reader:
            school_id   = (row.get('id') or '').strip()
            school_code = (row.get('code') or '').strip()
            if school_code and school_id and school_id.isdigit():
                code_to_id[school_code] = int(school_id)

    print(f"  Loaded {len(code_to_id)} school codes from {db_csv}")

    # ── Combine ──────────────────────────────────────────────────────────────
    name_to_id = {}
    unmatched  = []
    for name, code in name_to_code.items():
        if code in code_to_id:
            name_to_id[name] = code_to_id[code]
        else:
            unmatched.append(f"  code {code!r} ({name!r}) not found in DB")

    if unmatched:
        print(f"  WARNING: {len(unmatched)} school(s) in xlsx have no DB match:")
        for msg in unmatched:
            print(msg)

    print(f"  Final mapping: {len(name_to_id)} school name → id pairs")
    return name_to_id

def resolve_klados(specialty):
    """
    Maps an eData specialty string to a database klados code.

    Tries in order:
      1. Exact match (specialty already IS the code, e.g. 'ΠΕ70')
      2. Prefix match — the specialty string starts with a known code
         followed by a separator (space, dash, slash), e.g. 'ΠΕ70 - ΔΑΣΚΑΛΟΙ'
      3. Returns None if nothing matches (caller will warn + write NULL)
    """
    s = specialty.strip()

    # 1. Exact match
    if s in KLADOS_CODES:
        return s

    # 2. Prefix match — sort by length descending so 'ΠΕ79.01' is tried
    #    before 'ΠΕ79', avoiding a short code stealing a longer one's records
    for code in sorted(KLADOS_CODES, key=len, reverse=True):
        if s.startswith(code) and len(s) > len(code) and s[len(code)] in ' -/(':
            return code

    return None

def generate_sql(results, output_file="teachers_insert.sql", school_mapping=None):
    """
    Generates SQL INSERT statements for the teachers table.

    Args:
        results        : list of teacher data dicts (from scraper or load_from_xlsx)
        output_file    : path to the output .sql file
        school_mapping : optional dict { service_school name → organiki_id (int) }
                         If provided, organiki_id is looked up from service_school.
                         Teachers whose school is not found get organiki_id = NULL
                         and a warning is printed.
    """
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    lines        = []
    unresolved   = []

    for r in results:
        md5     = hashlib.md5(r['afm'].encode()).hexdigest()
        name    = r['name'].replace("'", "\\'")
        surname = r['surname'].replace("'", "\\'")
        afm     = r['afm'].replace("'", "\\'")
        mobile  = r['mobile'].replace("'", "\\'")
        email   = r['email'].replace("'", "\\'")
        klados_resolved = resolve_klados(r.get('specialty', ''))
        if klados_resolved is None:
            print(f"  WARNING: unresolved specialty {r.get('specialty', '')!r} for AM {r.get('am', '')!r}")
        klados = (klados_resolved or r['specialty']).replace("'", "\\'")
        am      = r['am'].replace("'", "\\'")
        org_eae = 1 if r.get('education_type', '') == 'Ειδική Αγωγή' else 0

        # Resolve organiki_id from service_school name
        if school_mapping is not None:
            service_school = r.get('service_school', '').strip()
            organiki_id    = school_mapping.get(service_school)
            if organiki_id is None:
                unresolved.append(f"  AM {r['am']!r}: service_school {service_school!r}")
                organiki_id_sql = 'NULL'
            else:
                organiki_id_sql = str(organiki_id)
        else:
            organiki_id_sql = '1'   # placeholder when no mapping provided

        sql = (
            f"INSERT INTO `teachers` "
            f"(`md5`, `created_at`, `updated_at`, `name`, `surname`, `fname`, `mname`, "
            f"`afm`, `gender`, `telephone`, `mail`, `sch_mail`, `klados`, `am`, "
            f"`sxesi_ergasias_id`, `org_eae`, `organiki_id`, `organiki_type`, "
            f"`active`, `sent_link_mail`, `is_director`, `is_subdirector`) VALUES ("
            f"'{md5}', '{now}', '{now}', '{name}', '{surname}', '', '', "
            f"'{afm}', '', '{mobile}', '{email}', NULL, '{klados}', '{am}', "
            f"1, {org_eae}, {organiki_id_sql}, 'App\\Models\\School', "
            f"1, 0, 0, 0) "
            f"ON DUPLICATE KEY UPDATE "
            f"`updated_at` = '{now}', `name` = '{name}', "
            f"`surname` = '{surname}', "
            f"`telephone` = '{mobile}', `mail` = '{email}', "
            f"`klados` = '{klados}', `am` = '{am}', "
            f"`sxesi_ergasias_id` = 1, `org_eae` = {org_eae}, "
            f"`organiki_id` = {organiki_id_sql}, `organiki_type` = 'App\\Models\\School', "
            f"`active` = 1, `is_director` = 0, `is_subdirector` = 0;"
        )
        lines.append(sql)

    if unresolved:
        print(f"\nWARNING: {len(unresolved)} teacher(s) could not be mapped to a school id "
              f"(organiki_id set to NULL):")
        for msg in unresolved:
            print(msg)

    with open(f"files/{output_file}", 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines))

    print(f"\nSQL file saved to files/{output_file} ({len(lines)} records)")


def _prompt_school_mapping():
    """
    Asks the user whether to load a school mapping and returns the mapping dict
    (or None if the user skips).
    """
    use_mapping = input("\nDo you want to map service_school → organiki_id? [y/N]: ").strip().lower()
    if use_mapping != 'y':
        return None

    schools_xlsx = input("  Path to Σχολικές Μονάδες xlsx file: ").strip()
    db_csv       = input("  Path to database schools CSV export : ").strip()
    print("  Building school mapping...")
    return load_school_mapping(schools_xlsx, db_csv)


def run_scraper():
    """Scrapes teacher data from eData and saves results to Excel + SQL."""
    settings_file = "files/settings.xlsx"
    username, password, _ = read_settings(settings_file)

    am_file = input("Enter path to AM codes file (xlsx): ").strip()
    am_codes = read_am_codes(am_file)

    driver = webdriver.Chrome()
    loginToEdata(driver, username, password)

    results = []

    try:
        for am_code in am_codes:
            success = navigate_to_teacher(driver, am_code)
            if success:
                data = retrieve_teacher_data(driver, am_code)
                if data:
                    results.append(data)
            else:
                print(f"Navigation failed for AM code: {am_code}")

        if results:
            df = pd.DataFrame(results)
            df.to_excel("/files/teacher_data.xlsx", index=False)
            print(f"Data saved to /files/teacher_data.xlsx ({len(results)} records)")

            school_mapping = _prompt_school_mapping()
            generate_sql(results, school_mapping=school_mapping)

    except Exception as e:
        print(f"An error occurred: {e}")

    finally:
        input("Press Enter to close the browser...")
        driver.quit()


def run_sql_from_xlsx():
    """Generates SQL from a manually edited Excel file (no browser needed)."""
    xlsx_file   = input("Enter path to teacher data Excel file: ").strip()
    output_file = input("Enter output SQL file name [teachers_insert.sql]: ").strip()
    if not output_file:
        output_file = "teachers_insert.sql"

    try:
        results = load_from_xlsx(xlsx_file)
        print(f"Loaded {len(results)} records from {xlsx_file}")

        school_mapping = _prompt_school_mapping()
        generate_sql(results, output_file, school_mapping=school_mapping)
    except Exception as e:
        print(f"Error: {e}")


def main():
    print("Select mode:")
    print("  1 - Scrape data from eData")
    print("  2 - Generate SQL from existing Excel file")
    choice = input("Enter 1 or 2: ").strip()

    if choice == '1':
        run_scraper()
    elif choice == '2':
        run_sql_from_xlsx()
    else:
        print("Invalid choice. Exiting.")


if __name__ == "__main__":
    main()