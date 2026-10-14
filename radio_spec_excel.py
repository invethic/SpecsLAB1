"""
Radio Electronic Systems - Specification Template (Excel version)
-----------------------------------------------------------------
Generates Radio_Spec_Template.xlsx with:
  - "Specification" sheet: all sections of the web tool, drop-down lists,
    dynamic hyperlinks (Genorma, 360compliance, FCC ID, EFIS, App Note)
  - "Lists" sheet: editable sources of every drop-down (dynamic ranges:
    add a value at the bottom of a column and it appears in the drop-down)
  - "Link Budget" sheet: live free-space link budget + radio horizon + chart

Requirements : pip install openpyxl
Usage        : python radio_spec_excel.py
"""
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.chart import ScatterChart, Reference, Series
from openpyxl.utils import get_column_letter

OUTPUT = "Radio_Spec_Template.xlsx"

# ---------------------------------------------------------------- styles
THIN = Side(style="thin", color="D1D5DB")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
INPUT_FILL = PatternFill("solid", fgColor="FFFBEB")
NOTE_FILL = PatternFill("solid", fgColor="F3F4F6")
GREY_FILL = PatternFill("solid", fgColor="E5E7EB")
LABEL_FONT = Font(bold=True, color="374151")
HELP_FONT = Font(italic=True, color="6B7280", size=9)
LINK_FONT = Font(color="2563EB", underline="single", bold=True)
WRAP_TOP = Alignment(wrap_text=True, vertical="top")
CENTER_V = Alignment(vertical="center", wrap_text=True)

SRD = "Short Range Device (SRD)"

# ---------------------------------------------------------------- drop-down sources
# strict=True  -> only list values accepted (same as the web tool selects)
# strict=False -> suggestions only, free text still accepted
LISTS = {
    "DeviceCategory": (True, [SRD, "Cellular", "Satellite", "Other"]),
    "CEPTAnnex": (True, [
        "Annex 1: Non-specific SRD",
        "Annex 2: Tracking, Tracing & Data Acquisition",
        "Annex 3: Wideband Data Transmission",
        "Annex 4: Alarms",
        "Annex 5: Model Control",
        "Annex 6: Inductive Applications",
        "Annex 7: Active Medical Implants (Ultra Low Power)",
        "Annex 8: RFID",
        "Annex 9: Radiodetermination Applications",
        "Annex 10: Ultra Wideband (UWB)",
        "Annex 11: Road Transport Telematics (RTT)",
        "Annex 12: Tank Level Probing Radar (TLPR)",
        "Annex 13: Wireless Audio Applications",
    ]),
    "AntennaImpl": (True, ["Antenna Connector", "PCB Printed Antenna",
                           "Chipset Recommended", "Custom Design"]),
    "ChipsetMfr": (True, ["Nordic", "Espressif", "Texas Instruments",
                          "STMicroelectronics", "NXP", "Silicon Labs",
                          "Quectel", "Simcom", "u-blox", "Other"]),
    "SARLimit": (True, ["2.0 W/kg EU Head/Body", "4.0 W/kg EU Limbs",
                        "1.6 W/kg USA Head/Body", "Other"]),
    "YesNo": (True, ["Yes", "No"]),
    # suggestion lists (free text allowed)
    "StationType": (False, ["End-User Terminal", "Base Station", "Repeater",
                            "Gateway", "Fixed Station", "Mobile Station"]),
    "OperatingMode": (False, ["Simplex", "Half-duplex", "Full-duplex"]),
    "Modulation": (False, ["FM", "AM", "OOK/ASK", "FSK", "GFSK", "LoRa (CSS)",
                           "BPSK", "QPSK", "16-QAM", "64-QAM", "OFDM"]),
    "Polarization": (False, ["Vertical", "Horizontal",
                             "Circular (RHCP)", "Circular (LHCP)"]),
    # ---- standards & certifications (suggestions, extend freely in 'Lists')
    "IEC_Safety": (False, ["IEC 62368-1", "IEC 60950-1", "IEC 61010-1",
                           "IEC 60601-1"]),
    "IEC_EMC": (False, ["IEC 61000-6-1", "IEC 61000-6-2", "IEC 61000-6-3",
                        "IEC 61000-6-4"]),
    "CENELEC_Safety": (False, ["EN 62368-1", "EN 62311", "EN 62479",
                               "EN 50385"]),
    "CENELEC_EMC": (False, ["EN 55032", "EN 55035", "EN 61000-6-1",
                            "EN 61000-6-2", "EN 61000-6-3", "EN 61000-6-4"]),
    "ETSI": (False, ["ETSI EN 301 489-1", "ETSI EN 301 489-3",
                     "ETSI EN 301 489-17", "ETSI EN 300 328",
                     "ETSI EN 300 220-1", "ETSI EN 300 220-2",
                     "ETSI EN 301 893", "ETSI EN 303 645"]),
    "IEEE": (False, ["IEEE 802.11", "IEEE 802.15.1", "IEEE 802.15.4",
                     "IEEE 802.3"]),
    "3GPP": (False, ["3GPP TS 36.101", "3GPP TS 24.301", "3GPP TS 38.101",
                     "3GPP TS 36.521"]),
    "Certifications": (False, ["CE (RED)", "FCC", "IC (ISED)", "UKCA",
                               "PTCRB", "GCF", "TELEC (MIC)", "SRRC"]),
    # ---- cybersecurity
    "CyberSec": (False, ["IEC 62443-3-2", "IEC 62443-3-3", "IEC 62443-4-1",
                         "IEC 62443-4-2", "EN 18031-1", "EN 18031-2",
                         "EN 18031-3", "ETSI EN 303 645", "ISO/IEC 27001"]),
}


# ---------------------------------------------------------------- field helpers
def F(key, label, kind="text", lst=None, help="", default=None,
      lines=1, formula=None, url=None):
    return dict(key=key, label=label, kind=kind, lst=lst, help=help,
                default=default, lines=lines, formula=formula, url=url)


def NOTE(text):
    return dict(kind="note", text=text)


def MULTI(key, label, lst, n, help=""):
    """n drop-down rows for the same category (one standard per cell).
    Change n to get more / fewer slots."""
    return [F(f"{key}{i}", f"{label} — {i}", "list", lst,
              help=help if i == 1 else "")
            for i in range(1, n + 1)]


# dynamic hyperlink formulas (r = dict key -> absolute cell ref)
def genorma_link(r):
    a = r["application"]
    kw = f'IF({a}="","SRD",{a})'
    url = (f'"https://genorma.com/en/project/advanced-search?keywords="'
           f'&SUBSTITUTE({kw}," ","%20")'
           '&"&status%5B0%5D=PUBLISHED&status%5B2%5D=WITHDRAWN'
           '&operatorCommittee=1&operatorStdType=3&listMode=DEFAULT"')
    return f'=HYPERLINK({url},"→ Search "&{kw}&" standards on Genorma")'


def market_link(r):
    m = r["mainTargetMarket"]
    kw = f'IF({m}="","USA",{m})'
    url = (f'"https://360compliance.co/marketaccess/"'
           f'&LOWER(SUBSTITUTE({kw}," ","%20"))&"/"')
    return f'=HYPERLINK({url},"→ Market access requirements for "&{kw})'


def fcc_link(r):
    i = r["fccId"]
    return (f'=IF({i}="","(fill in the FCC-ID field in section 7)",'
            f'HYPERLINK("https://fcc.report/FCC-ID/"&{i}&"/",'
            f'"→ View FCC ID report: "&{i}))')


def cept_help(r):
    c, own = r["deviceCategory"], r["ceptAnnex"]
    return (f'=IF({c}="{SRD}","Select the ERC Rec. 70-03 annex matching the application",'
            f'IF({own}="","Only applicable when Device Category = SRD",'
            f'"⚠ Annex selected but Device Category is not SRD"))')


def appnote_help(r):
    own = r["applicationNoteUrl"]
    return f'=IF({own}="","Paste the datasheet / application note URL",HYPERLINK({own},"→ Open Application Note"))'


# ---------------------------------------------------------------- sections
SECTIONS = [
    ("PROJECT INFORMATION", "1F4E79", [
        F("projectName", "Project Name", help="e.g. VHF System"),
        F("designer", "Designer"),
        F("date", "Date", "date", default=date.today()),
        F("version", "Version", default="1.0", help="e.g. 1.0"),
    ]),
    ("1. GENERAL SPECIFICATIONS", "2563EB", [
        F("systemType", "System Type", help="e.g. Transceiver"),
        F("deviceCategory", "Device Category", "list", "DeviceCategory"),
        F("ceptAnnex", "CEPT Rec. 70-03 Annex", "list", "CEPTAnnex", help=cept_help),
        F("stationType", "Station Type", "list", "StationType",
          help="Suggestions in the list, free text accepted"),
        F("application", "Application",
          help="e.g. Commercial — also drives the Genorma search link (section 9)"),
        F("intendedUse", "Intended Use", help="e.g. WiFi Router, IoT Gateway"),
        F("mainTargetMarket", "Main Target Market",
          help="e.g. France — drives the market access link (one country works best)"),
        F("operatingMode", "Operating Mode", "list", "OperatingMode",
          help="Suggestions in the list, free text accepted"),
    ]),
    ("2. FREQUENCY SPECIFICATIONS", "2563EB", [
        F("efisLink", "EFIS database", "static_link",
          default="→ Open EFIS (ECO Frequency Information System)",
          url="https://efis.cept.org/", help="Check band allocation per country"),
        F("freqRange", "Frequency Range", help="e.g. 144-148 MHz"),
        F("centerFreq", "Center Frequency", help="e.g. 146 MHz"),
        F("channelSpacing", "Channel Spacing", help="e.g. 25 kHz"),
        F("bandwidth", "Bandwidth", help="e.g. 16 kHz"),
        F("efisVerified", "EFIS Verified", "list", "YesNo", default="No"),
        F("efisNotes", "EFIS Notes", "long", lines=2),
    ]),
    ("3. TRANSMITTER SPECIFICATIONS", "DC2626", [
        F("txPower", "Output Power", help="e.g. 5 W"),
        F("txFreqStability", "Frequency Stability", help="e.g. ±2.5 ppm"),
        F("spuriousEmission", "Spurious Emission", help="e.g. < -60 dBc"),
        F("modulation", "Modulation", "list", "Modulation",
          help="Suggestions in the list, free text accepted"),
    ]),
    ("4. RECEIVER SPECIFICATIONS", "CA8A04", [
        F("rxSensitivity", "Sensitivity", help="e.g. -120 dBm"),
        F("selectivity", "Selectivity", help="e.g. > 70 dB"),
        F("noiseFigure", "Noise Figure", help="e.g. < 5 dB"),
        F("dynamicRange", "Dynamic Range", help="e.g. 90 dB"),
    ]),
    ("5. ANTENNA SPECIFICATIONS", "DB2777", [
        F("antennaImplementation", "Implementation Type", "list", "AntennaImpl"),
        F("antennaType", "Type", help="e.g. Dipole"),
        F("antennaGain", "Gain", help="e.g. 2.15 dBi"),
        F("vswr", "VSWR", help="e.g. < 1.5:1"),
        F("polarization", "Polarization", "list", "Polarization",
          help="Suggestions in the list, free text accepted"),
    ]),
    ("6. FEATURES & CHARACTERISTICS", "0891B2", [
        F("inputVoltage", "Input Voltage", help="e.g. 13.8 V DC"),
        F("powerConsumption", "Power Consumption", help="e.g. RX: 1.5 A, TX: 10 A"),
        F("tempRange", "Temperature Range", help="e.g. -20 °C to +60 °C"),
        F("humidity", "Humidity", help="e.g. 5 % to 95 % RH"),
        F("dimensions", "Dimensions", help="e.g. 200×100×250 mm"),
        F("weight", "Weight", help="e.g. 2.5 kg"),
    ]),
    ("7. RF CHIPSET", "0D9488", [
        F("chipsetManufacturer", "Manufacturer", "list", "ChipsetMfr"),
        F("chipsetModel", "Model", help="e.g. nRF52840, SIM7028"),
        F("chipsetDetails", "Details", "long", lines=2,
          help="Integrated PA, protocol support, etc."),
        F("applicationNoteUrl", "Application Note URL", help=appnote_help),
        F("fccId", "FCC-ID (chipset / radio module)",
          help="e.g. 2AJYU-8EC0002 — drives the FCC report link (section 9)"),
    ]),
    ("8. CYBERSECURITY", "059669", [
        NOTE("Common standards: IEC 62443-3/-4 (Industrial security) • "
             "EN 18031-1/-2/-3 (Consumer IoT security) • ETSI EN 303 645 • "
             "ISO/IEC 27001 — pick one per row, free text accepted"),
        *MULTI("cyberSecurityStandard", "Applicable CyberSecurity Standard",
               "CyberSec", 6, help="Drop-down list, values editable in 'Lists' sheet"),
        F("cyberSecurityNotes", "CyberSecurity Notes", "long", lines=2,
          help="Scope, security level (SL), conformity assessment route…"),
    ]),
    ("9. STANDARDS & CERTIFICATIONS", "475569", [
        F("genormaLink", "Genorma standards search", "formula", formula=genorma_link,
          help="Built from the Application field"),
        F("marketLink", "Market access (360compliance)", "formula", formula=market_link,
          help="Built from the Main Target Market field"),
        F("fccLink", "FCC ID report", "formula", formula=fcc_link,
          help="Built from the FCC-ID field"),
        NOTE("One standard per row — pick from the drop-down or type freely. "
             "Add your own values in the 'Lists' sheet; change the slot count "
             "in MULTI(..., n) in the script."),
        *MULTI("iecSafety", "IEC (Safety)", "IEC_Safety", 2),
        *MULTI("iecEmc", "IEC (EMC)", "IEC_EMC", 3),
        *MULTI("cenelecSafety", "CENELEC (Safety)", "CENELEC_Safety", 3),
        *MULTI("cenelecEmc", "CENELEC (EMC)", "CENELEC_EMC", 3),
        *MULTI("etsi", "ETSI", "ETSI", 4),
        *MULTI("ieee", "IEEE", "IEEE", 2),
        *MULTI("threeGpp", "3GPP", "3GPP", 2),
        *MULTI("certifications", "Certification", "Certifications", 4),
    ]),
    ("10. SAR SAFETY", "EA580C", [
        NOTE("SAR measures RF energy absorption (EU: 2 W/kg head/body, 4 W/kg limbs; "
             "USA: 1.6 W/kg head/body)"),
        F("sarHeadContact", "Foreseeable use: Head contact", "list", "YesNo", default="No"),
        F("sarBodyContact", "Foreseeable use: Body contact", "list", "YesNo", default="No"),
        F("sarLimbsContact", "Foreseeable use: Arms/Legs contact", "list", "YesNo", default="No"),
        F("sarLimit", "SAR Limit", "list", "SARLimit"),
        F("sarNotes", "SAR Notes"),
    ]),
    ("11. USE CASES", "4F46E5", [
        F("useCase1", "Use Case 1", "long", lines=3,
          help="e.g. Home user connects multiple devices to internet via WiFi"),
        F("useCase2", "Use Case 2", "long", lines=3,
          help="e.g. Small business office network with 20+ devices"),
        F("useCase3", "Use Case 3", "long", lines=3,
          help="e.g. IoT sensor network for industrial monitoring"),
    ]),
    ("12. FUNCTIONAL REQUIREMENTS", "9333EA", [
        F("requirement1", "Main Requirement 1", help="e.g. Network Connectivity"),
        F("requirement1sub1", "      1.1 Sub-requirement", help="e.g. Support 2.4 GHz WiFi"),
        F("requirement1sub2", "      1.2 Sub-requirement", help="e.g. Support 5 GHz WiFi"),
        F("requirement2", "Main Requirement 2", help="e.g. Data Throughput"),
        F("requirement2sub1", "      2.1 Sub-requirement", help="e.g. Minimum 100 Mbps"),
        F("requirement2sub2", "      2.2 Sub-requirement", help="e.g. Maximum latency 50 ms"),
        F("requirement3", "Main Requirement 3", help="e.g. Security Features"),
        F("requirement3sub1", "      3.1 Sub-requirement", help="e.g. WPA3 encryption"),
        F("requirement3sub2", "      3.2 Sub-requirement", help="e.g. Firewall enabled"),
    ]),
]


# ---------------------------------------------------------------- Lists sheet
def build_lists(wb):
    ws = wb.create_sheet("Lists")
    validations = {}
    for col_idx, (name, (strict, values)) in enumerate(LISTS.items(), start=1):
        col = get_column_letter(col_idx)
        head = ws.cell(1, col_idx, name)
        head.font = Font(bold=True, color="FFFFFF")
        head.fill = PatternFill("solid", fgColor="1F4E79")
        for i, v in enumerate(values, start=2):
            ws.cell(i, col_idx, v)
        ws.column_dimensions[col].width = max(18, max(len(v) for v in values) + 2)
        # dynamic range: grows automatically when values are added below
        src = f"=OFFSET(Lists!${col}$2,0,0,COUNTA(Lists!${col}:${col})-1,1)"
        dv = DataValidation(type="list", formula1=src, allow_blank=True)
        if strict:
            dv.showErrorMessage = True
            dv.errorStyle = "stop"
            dv.errorTitle = "Invalid value"
            dv.error = "Please pick a value from the list (edit the 'Lists' sheet to add one)."
        else:
            dv.showErrorMessage = False  # free text accepted
        validations[name] = dv
    ws.freeze_panes = "A2"
    note = ws.cell(len(max((v for _, v in LISTS.values()), key=len)) + 4, 1,
                   "Add new values at the bottom of a column: drop-downs update automatically.")
    note.font = HELP_FONT
    return validations


# ---------------------------------------------------------------- Specification sheet
def build_spec(wb, validations):
    ws = wb.active
    ws.title = "Specification"
    ws.column_dimensions["A"].width = 36
    ws.column_dimensions["B"].width = 62
    ws.column_dimensions["C"].width = 58

    ws.merge_cells("A1:C1")
    t = ws["A1"]
    t.value = "RADIO ELECTRONIC SYSTEMS — Technical Specification"
    t.font = Font(bold=True, size=16, color="FFFFFF")
    t.fill = PatternFill("solid", fgColor="312E81")
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32

    ws.merge_cells("A2:C2")
    s = ws["A2"]
    s.value = ("Yellow cells = input. Drop-downs from the 'Lists' sheet. "
               "Blue links update automatically from your entries.")
    s.font = HELP_FONT
    s.alignment = Alignment(horizontal="center")

    for col, h in zip("ABC", ("Parameter", "Value", "Example / Help")):
        c = ws[f"{col}3"]
        c.value = h
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="4B5563")
        c.border = BORDER

    refs, cept_row, row = {}, None, 4

    for title, color, fields in SECTIONS:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
        h = ws.cell(row, 1, title)
        h.font = Font(bold=True, size=12, color="FFFFFF")
        h.fill = PatternFill("solid", fgColor=color)
        h.alignment = Alignment(vertical="center", indent=1)
        ws.row_dimensions[row].height = 22
        row += 1

        for fd in fields:
            if fd["kind"] == "note":
                ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
                n = ws.cell(row, 1, fd["text"])
                n.font = Font(italic=True, color="374151", size=10)
                n.fill = NOTE_FILL
                n.alignment = CENTER_V
                ws.row_dimensions[row].height = 30
                row += 1
                continue

            refs[fd["key"]] = f"$B${row}"
            lab = ws.cell(row, 1, fd["label"])
            lab.font = LABEL_FONT
            lab.alignment = CENTER_V
            lab.border = BORDER

            val = ws.cell(row, 2)
            val.border = BORDER
            kind = fd["kind"]

            if kind == "formula":
                val.value = fd["formula"](refs)
                val.font = LINK_FONT
                val.alignment = CENTER_V
            elif kind == "static_link":
                val.value = fd["default"]
                val.hyperlink = fd["url"]
                val.font = LINK_FONT
                val.alignment = CENTER_V
            else:
                val.fill = INPUT_FILL
                val.alignment = WRAP_TOP if kind == "long" else CENTER_V
                if fd["default"] is not None:
                    val.value = fd["default"]
                if kind == "date":
                    val.number_format = "yyyy-mm-dd"
                if kind == "list":
                    validations[fd["lst"]].add(val.coordinate)

            hlp = fd["help"]
            hc = ws.cell(row, 3, hlp(refs) if callable(hlp) else hlp)
            hc.font = LINK_FONT if fd["key"] == "applicationNoteUrl" else HELP_FONT
            hc.alignment = WRAP_TOP if kind == "long" else CENTER_V
            hc.border = BORDER

            lines = max(fd["lines"], hlp.count("\n") + 1 if isinstance(hlp, str) else 1)
            ws.row_dimensions[row].height = max(20, 16 * lines)

            if fd["key"] == "ceptAnnex":
                cept_row = row
            row += 1
        row += 1  # spacer

    for dv in validations.values():
        ws.add_data_validation(dv)

    # grey out CEPT annex row when Device Category is not SRD
    if cept_row:
        ws.conditional_formatting.add(
            f"A{cept_row}:C{cept_row}",
            FormulaRule(formula=[f'{refs["deviceCategory"]}<>"{SRD}"'],
                        fill=GREY_FILL, font=Font(color="9CA3AF")))

    ws.freeze_panes = "A4"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_title_rows = "3:3"


# ---------------------------------------------------------------- Link Budget sheet
def build_link_budget(wb):
    ws = wb.create_sheet("Link Budget")
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 14
    for c in "DEFG":
        ws.column_dimensions[c].width = 16

    ws.merge_cells("A1:G1")
    t = ws["A1"]
    t.value = "LINK BUDGET — free-space model + radio horizon"
    t.font = Font(bold=True, size=14, color="FFFFFF")
    t.fill = PatternFill("solid", fgColor="312E81")

    inputs = [
        (3, "TX output power (W)", 5, True),
        (4, "TX output power (dBm)", "=10*LOG10(B3*1000)", False),
        (5, "TX antenna gain (dBi)", 2.15, True),
        (6, "RX antenna gain (dBi)", 2.15, True),
        (7, "Cable / other losses (dB)", 0, True),
        (8, "RX sensitivity (dBm)", -120, True),
        (9, "Frequency (MHz)", 146, True),
        (10, "Fade margin (dB)", 10, True),
        (11, "TX antenna height (m)", 10, True),
        (12, "RX antenna height (m)", 1.5, True),
        (14, "Max free-space range (km)",
         "=10^((B4+B5+B6-B7-B8-B10-32.45-20*LOG10(B9))/20)", False),
        (15, "Radio horizon (km)", "=4.12*(SQRT(B11)+SQRT(B12))", False),
        (16, "Estimated max range (km)", "=MIN(B14,B15)", False),
    ]
    for r, label, v, is_input in inputs:
        a = ws.cell(r, 1, label)
        a.font = LABEL_FONT
        a.border = BORDER
        b = ws.cell(r, 2, v)
        b.border = BORDER
        b.number_format = "0.00"
        if is_input:
            b.fill = INPUT_FILL
        else:
            b.font = Font(bold=True, color="1E40AF")
    ws["A16"].font = Font(bold=True, color="059669")
    ws["B16"].font = Font(bold=True, color="059669", size=12)

    ws.merge_cells("A18:B21")
    n = ws["A18"]
    n.value = ("Free-space loss is optimistic: at VHF/UHF the range is usually limited "
               "by the radio horizon, terrain and buildings. Use the estimate as an "
               "upper bound.")
    n.font = HELP_FONT
    n.alignment = WRAP_TOP

    heads = ["Distance (km)", "FSPL (dB)", "RX power (dBm)", "Sens. + margin (dBm)"]
    for i, h in enumerate(heads):
        c = ws.cell(3, 4 + i, h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="4B5563")
        c.alignment = CENTER_V

    first, n_pts = 4, 41  # 0.1 km -> 1000 km, 10 points per decade
    for i in range(n_pts):
        r = first + i
        ws.cell(r, 4, f"=10^(-1+0.1*{i})").number_format = "0.00"
        ws.cell(r, 5, f"=32.45+20*LOG10($B$9)+20*LOG10(D{r})").number_format = "0.0"
        ws.cell(r, 6, f"=$B$4+$B$5+$B$6-$B$7-E{r}").number_format = "0.0"
        ws.cell(r, 7, "=$B$8+$B$10").number_format = "0.0"
    last = first + n_pts - 1

    chart = ScatterChart()
    chart.title = "Received power vs distance"
    chart.x_axis.title = "Distance (km, log)"
    chart.y_axis.title = "Power (dBm)"
    chart.x_axis.scaling.logBase = 10
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.height, chart.width = 9, 18
    xs = Reference(ws, min_col=4, min_row=first, max_row=last)
    for col in (6, 7):
        ys = Reference(ws, min_col=col, min_row=3, max_row=last)
        s = Series(ys, xs, title_from_data=True)
        s.marker.symbol = "none"
        s.smooth = False
        chart.series.append(s)
    ws.add_chart(chart, "I3")


# ---------------------------------------------------------------- main
def main():
    wb = Workbook()
    validations = build_lists(wb)
    build_spec(wb, validations)
    build_link_budget(wb)
    wb.move_sheet("Lists", offset=1)  # order: Specification, Link Budget, Lists
    wb.save(OUTPUT)
    print(f"✅ {OUTPUT} generated")


if __name__ == "__main__":
    main()
