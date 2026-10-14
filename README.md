# SpecsLAB1
How to define the specifications of Wireless Product
# Radio Electronic Systems — Specification Template (Excel)

A structured Excel template for writing the technical specification of a radio product with **regulatory compliance built in from day one**. A Python script (`radio_spec_excel.py`) generates the workbook.

---

## 1. Why this template exists

This template comes from **17 years of experience in radio equipment certification and conformity assessment**. Over that time, one pattern kept coming back: most compliance problems are not discovered in the test lab. They are designed in, months earlier, when the specification was written without regulatory input.

Some typical examples:

- A frequency band is chosen without checking its national allocation, and the product cannot be sold in a key market.
- Nobody asks early enough whether the device will be used against the head or body, and SAR testing appears late in the project, with its cost and delays.
- A pre-certified radio module is integrated without its FCC ID or test reports being collected, and the certification evidence has to be rebuilt.
- Cybersecurity requirements are discovered at the end of the project. In the EU, articles 3.3(d), (e) and (f) of the Radio Equipment Directive have applied to many connected products since August 2025, through the harmonised standards EN 18031-1/-2/-3.
- The applicable standards are listed from memory and are incomplete or outdated, so the test plan and the Declaration of Conformity rest on weak foundations.

The template addresses this by making the designer answer the questions a certification engineer would ask, **at the specification stage**. Each section maps to a part of the conformity assessment file:

| Template section | Conformity assessment relevance |
|---|---|
| 1. General specifications | Product scope, device category, intended use, target market. These decide which regulations apply (RED, FCC, ISED…). |
| 2. Frequency specifications | Spectrum use (RED art. 3.2). Band allocation is checked in the CEPT EFIS database. For SRDs, the matching annex of ERC Recommendation 70-03 is selected. |
| 3–5. Transmitter, receiver, antenna | Radio parameters that drive radio testing: output power, spurious emissions, receiver performance, antenna gain. |
| 6. Features & characteristics | Power supply and environmental conditions, which define the test conditions. |
| 7. RF chipset | Traceability of the radio module and reuse of existing certification evidence (FCC ID, application notes). |
| 8. Cybersecurity | RED art. 3.3(d)(e)(f) / EN 18031, ETSI EN 303 645, IEC 62443 for industrial products, ISO/IEC 27001. |
| 9. Standards & certifications | Safety (art. 3.1a), EMC (art. 3.1b), radio standards (ETSI, IEEE, 3GPP) and target marks (CE, FCC, IC, UKCA…). |
| 10. SAR safety | RF exposure assessment, triggered by the foreseeable use (head, body, limbs). |
| 11–12. Use cases & functional requirements | Link between intended use and technical requirements. The intended use must be consistent across the whole technical file. |

The aim is not to replace a test lab or a Notified Body. The aim is to **arrive at the lab with a coherent, complete specification**, which is the cheapest way to pass first time.

---

## 2. What the workbook contains

**Specification sheet.** This is the main form, with 12 colour-coded sections. Yellow cells are inputs and column C gives examples and guidance. Most fields with predefined values use drop-down lists:

- the device category
- the 13 annexes of ERC Rec. 70-03
- antenna implementation and chipset manufacturer
- modulation, polarization and operating mode
- SAR limits
- every standard family (IEC, CENELEC, ETSI, IEEE, 3GPP, certifications, cybersecurity)

The blue links update automatically from your entries. The Genorma standards search is built from the Application field, the 360compliance market-access page from the target market, and the FCC ID report from the module's FCC ID. The EFIS link is fixed.

**Link Budget sheet.** This gives a quick sizing of the radio link. You enter the TX power, antenna gains, losses, receiver sensitivity, frequency, fade margin and antenna heights. The sheet returns the free-space range, the radio horizon, and the estimated maximum range (the smaller of the two), with a chart. Free space is an optimistic model, so treat the result as an upper bound.

**Lists sheet.** This holds the sources of every drop-down. It is meant to be completed over time with your own standards, chipsets and certification marks.

---

## 3. Generating the workbook

### Requirements

You need **Python 3.8 or later** and the **openpyxl** library. Excel (2013 or later) or LibreOffice Calc can open the result.

### Step-by-step

1. Check that Python is installed:
   ```bash
   python --version
   ```
   On macOS or Linux, use `python3` if `python` is not found.

2. Install the dependency:
   ```bash
   pip install openpyxl
   ```
   (or `pip3 install openpyxl` / `python -m pip install openpyxl`)

3. Put `radio_spec_excel.py` in a folder and open a terminal in that folder.

4. Run the script:
   ```bash
   python radio_spec_excel.py
   ```
   The terminal should display:
   ```
   ✅ Radio_Spec_Template.xlsx generated
   ```

5. Open `Radio_Spec_Template.xlsx`. Keep a blank copy as your master, and save one filled copy per project, for example `Radio_Spec_<Project>_v1.0.xlsx`.

If an older generated file is open in Excel, close it before rerunning the script. Otherwise the save fails with a permission error.

---

## 4. Customising the template

### Adding values to a drop-down (no Python needed)

Open the **Lists** sheet and type the new value at the bottom of the matching column. The drop-down updates immediately, because the ranges are dynamic.

There are two kinds of lists:

- **Strict lists** (device category, CEPT annex, antenna implementation, chipset manufacturer, SAR limit, Yes/No) only accept values from the list.
- **Suggestion lists** (standards, modulation, polarization…) show the drop-down but also accept free text.

### Changing the number of rows per standard family

In the script, each standard family is declared with `MULTI(key, label, list, n)`. The last argument is the number of rows. For example:

```python
*MULTI("etsi", "ETSI", "ETSI", 4),   # 4 rows
*MULTI("etsi", "ETSI", "ETSI", 6),   # 6 rows
```

Rerun the script after any change.

### Changing the initial list contents

Edit the `LISTS` dictionary at the top of the script. Each entry is `name: (strict, [values])`. To make a list strict, set the first value to `True`; to accept free text, set it to `False`.

---

## 5. Troubleshooting

| Problem | Solution |
|---|---|
| `ModuleNotFoundError: No module named 'openpyxl'` | Run `pip install openpyxl`. If you have several Python versions, use `python -m pip install openpyxl`. |
| `PermissionError` when saving | Close the `.xlsx` file in Excel and run the script again. |
| Links do not open | Use Ctrl+click (Cmd+click on Mac), and make sure editing is enabled if Excel opened the file in Protected View. |
| The chart's log axis looks odd in LibreOffice | Right-click the X axis, choose *Format Axis*, then *Scale*, and tick *Logarithmic scale*. |
| A drop-down does not show a value I added | Make sure the value is directly below the last entry in its column, with no empty row in between. |

---

## 6. Disclaimer
This template is an engineering and teaching aid. It helps you structure a specification and anticipate regulatory requirements. It does not replace an analysis of the applicable legislation, testing by an accredited laboratory, or a Notified Body assessment where one is required. Standards are regularly revised, so always check the current version and harmonised status of each standard (for example in the Official Journal of the EU or on Genorma) before relying on it.

This template is an engineering and teaching aid. It helps you structure a specification and anticipate regulatory requirements. It does not replace an analysis of the applicable legislation, testing by an accredited laboratory, or a Notified Body assessment where one is required. Standards are regularly revised, so always check the current version and harmonised status of each standard (for example in the Official Journal of the EU or on Genorma) before relying on it.
