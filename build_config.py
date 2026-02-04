"""
GFB Config Builder - Creates JSON mapping configuration

This script analyzes the Excel file and creates a JSON config that maps
each output column to its source location by searching for text patterns.

The OUTPUT ORDER is fixed (hardcoded in map.py).
This config only stores WHERE to find each measure in the source Excel.
"""

import pandas as pd
import json
from datetime import datetime
import os
import re


def clean_text(text):
    """Clean and normalize text for matching"""
    if pd.isna(text) or text is None:
        return ""
    return str(text).strip()


def find_text_in_sheet(df, search_patterns, start_row=0, end_row=None, parent_section=None):
    """
    Search for text pattern in column A of the sheet with context awareness

    Args:
        df: DataFrame to search
        search_patterns: List of regex patterns to match
        start_row: Start searching from this row
        end_row: Stop searching at this row (None = search all)
        parent_section: Parent section text that must appear in the previous non-empty row

    Returns:
        Row index or None if not found
    """
    if end_row is None:
        end_row = len(df)

    for row_idx in range(start_row, min(end_row, len(df))):
        cell_value = clean_text(df.iloc[row_idx, 0])

        if not cell_value:
            continue

        # Check if pattern matches
        for pattern in search_patterns:
            if re.search(pattern, cell_value, re.IGNORECASE):
                # If parent_section is specified, check if it appears in recent rows above
                if parent_section:
                    # Look back up to 15 rows to find the parent section (Own holdings has many sub-items)
                    parent_found = False
                    for prev_idx in range(row_idx - 1, max(start_row - 1, row_idx - 16, -1), -1):
                        prev_value = clean_text(df.iloc[prev_idx, 0])
                        if prev_value and parent_section.lower() in prev_value.lower():
                            parent_found = True
                            break

                    if parent_found:
                        return row_idx  # Found with correct parent section
                    else:
                        # Parent section not found nearby, continue searching
                        break
                else:
                    # No parent section required, return immediately
                    return row_idx

    return None


def find_date_row(df):
    """Find the row containing dates"""
    print("  Searching for date row...")

    for row_idx in range(min(20, len(df))):
        # Check columns B onwards for date patterns
        row_values = df.iloc[row_idx, 1:10].astype(str).tolist()

        # Look for date patterns
        date_patterns = [
            r'\d{4}-\d{2}',  # YYYY-MM
            r'\d{1,2}[/-]\d{1,2}[/-]\d{2,4}',  # Date format
            r'(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)',  # Month names
        ]

        matches = sum(1 for val in row_values if any(re.search(pat, str(val), re.IGNORECASE) for pat in date_patterns))

        if matches >= 3:
            print(f"    OK Found date row at row {row_idx + 1} (index {row_idx})")
            return row_idx

    print(f"    X Date row not found, defaulting to row 14 (index 13)")
    return 13  # Default


def create_measure_definitions():
    """
    Create the fixed output measure definitions in the correct order.
    Each measure has search patterns to find it in the source Excel.
    """

    # 54 BORROWING measures in fixed order
    borrowing_measures = [
        {
            "code": "GFB.BORR.FINBUDGET.LOANFIN.M",
            "description": " Financing Federal budget, special funds; loan financing FMS & ESF",
            "search_patterns": [
                r"Financing Federal budget, special funds.*loan financing",
                r"Finanzierungsbedarf.*Bundeshaushalt.*Sondervermögen.*Darlehensfin"
            ]
        },
        {
            "code": "GFB.BORR.FINBUDGET.M",
            "description": "Gross Borrowing Requirement: Financing Federal budget and special funds",
            "search_patterns": [
                r"^Financing Federal budget and special funds",
                r"^Finanzierungsbedarf Bundeshaushalt und Sondervermögen"
            ]
        },
        {
            "code": "GFB.BORR.BREAKPURP.M",
            "description": "Gross Borrowing Requirement: Breakdown by purpose",
            "search_patterns": [
                r"^Breakdown by purpose",
                r"^Gliederung nach Zweck"
            ]
        },
        {
            "code": "GFB.BORR.FEDBUDGET.M",
            "description": "Gross Borrowing Requirement: Federal budget",
            "search_patterns": [
                r"^\s*Federal budget\s*$",
                r"^\s*Bundeshaushalt\s*$"
            ]
        },
        {
            "code": "GFB.BORR.FINMARKETFUND.M",
            "description": "Gross Borrowing Requirement: Financial Market Stabilisation Fund",
            "search_patterns": [
                r"Financial Market Stabilisation Fund",
                r"Finanzmarktstabilisierungsfonds"
            ]
        },
        {
            "code": "GFB.BORR.FMSLOANEXP.M",
            "description": "Gross Borrowing Requirement: FMS loans for expenses acc. to section 9 (1)  StFG",
            "search_patterns": [
                r"FMS loans for expenses.*section 9.*1",
                r"FMS-Darlehen für Ausgaben.*§ 9.*1"
            ]
        },
        {
            "code": "GFB.BORR.FMSLOANWINDAGE.M",
            "description": "Gross Borrowing Requirement: FMS loans for wind-up agencies acc. to section 9 (5)  StFG",
            "search_patterns": [
                r"FMS loans for wind-up agencies.*section 9.*5",
                r"FMS-Darlehen für Abwicklungsanstalten.*§ 9.*5"
            ]
        },
        {
            "code": "GFB.BORR.INVREDFUND.M",
            "description": "Gross Borrowing Requirement: Investment and Redemption Fund",
            "search_patterns": [
                r"Investment and Redemption Fund",
                r"Investitions- und Tilgungsfonds"
            ]
        },
        {
            "code": "GFB.BORR.ECONSTABFUND.M",
            "description": "Gross Borrowing Requirement: Economic Stabilisation Fund",
            "search_patterns": [
                r"^Economic Stabilisation Fund\s*$",
                r"^Wirtschaftsstabilisierungsfonds\s*$"
            ]
        },
        {
            "code": "GFB.BORR.ESFLOANRECAPMEAS.M",
            "description": "Gross Borrowing Requirement: ESF loans for recapitalisation measures acc. to section 22 StFG",
            "search_patterns": [
                r"ESF loans for recapitalisation.*section 22",
                r"WSF-Darlehen für Rekapitalisierung.*§ 22"
            ]
        },
        {
            "code": "GFB.BORR.ESFLOANKFW23.M",
            "description": "Gross Borrowing Requirement: ESF loans for KfW acc. to section 23 StFG",
            "search_patterns": [
                r"ESF loans for KfW.*section 23",
                r"WSF-Darlehen an KfW.*§ 23"
            ]
        },
        {
            "code": "GFB.BORR.ESFLOANENERGCRIS.M",
            "description": "Gross Borrowing Requirement: ESF loans to mitigate consequences of the energy crisis acc. to sect. 26a (1) no. 1-4 StFG",
            "search_patterns": [
                r"ESF loans.*energy crisis.*26a.*1.*1-4",
                r"WSF-Darlehen.*Energiekrise.*26a.*1.*1-4"
            ]
        },
        {
            "code": "GFB.BORR.ESFLOANKFW26.M",
            "description": "Gross Borrowing Requirement: ESF loans to KfW acc. to section 26a (1) no 5 StFG",
            "search_patterns": [
                r"ESF loans to KfW.*section 26a.*1.*5",
                r"WSF-Darlehen an KfW.*§ 26a.*1.*5"
            ]
        },
        {
            "code": "GFB.BORR.SPECFUNDBUND.M",
            "description": "Gross Borrowing Requirement: Special Fund for the Bundeswehr",
            "search_patterns": [
                r"Special Fund for the Bundeswehr",
                r"Sondervermögen Bundeswehr"
            ]
        },
        {
            "code": "GFB.BORR.RESTRFUND.M",
            "description": "Gross Borrowing Requirement: Restructuring Fund",
            "search_patterns": [
                r"^Restructuring Fund\s*$",
                r"^Restrukturierungsfonds\s*$"
            ]
        },
        {
            "code": "GFB.BORR.HARDCOALEQFUND.M",
            "description": "Gross Borrowing Requirement: Hard Coal Equalisation Fund",
            "search_patterns": [
                r"Hard Coal Equalisation Fund",
                r"Steinkohlefinanzierungsfonds"
            ]
        },
        {
            "code": "GFB.BORR.FEDRAILFUND.M",
            "description": "Gross Borrowing Requirement: Federal Railways Fund",
            "search_patterns": [
                r"Federal Railways Fund",
                r"Bundeseisenbahnvermögen"
            ]
        },
        {
            "code": "GFB.BORR.COMPFUND.M",
            "description": "Gross Borrowing Requirement: Compensation Fund",
            "search_patterns": [
                r"^Compensation Fund\s*$",
                r"^Ausgleichsfonds\s*$"
            ]
        },
        {
            "code": "GFB.BORR.REDFUNDLIAB.M",
            "description": "Gross Borrowing Requirement: Redemption Fund for Inherited Liabilities",
            "search_patterns": [
                r"Redemption Fund for Inherited Liabilities",
                r"Erblastentilgungsfonds"
            ]
        },
        {
            "code": "GFB.BORR.ERPSPECFUND.M",
            "description": "Gross Borrowing Requirement: ERP Special Fund",
            "search_patterns": [
                r"ERP Special Fund",
                r"ERP-Sondervermögen"
            ]
        },
        {
            "code": "GFB.BORR.GERUNITFUND.M",
            "description": "Gross Borrowing Requirement: German Unity Fund",
            "search_patterns": [
                r"German Unity Fund",
                r"Fonds Deutsche Einheit"
            ]
        },
        {
            "code": "GFB.BORR.EQUALBURFUND.M",
            "description": "Gross Borrowing Requirement: Equalisation of Burdens Fund",
            "search_patterns": [
                r"Equalisation of Burdens Fund",
                r"Lastenausgleichsfonds"
            ]
        },
        {
            "code": "GFB.BORR.BREAKDEPTTYPE.M",
            "description": "Gross Borrowing Requirement: Breakdown by debt type",
            "search_patterns": [
                r"^Breakdown by debt type\s*$",
                r"^Gliederung nach Schuldarten\s*$"
            ]
        },
        {
            "code": "GFB.BORR.BREAKDEPTTYPE.FEDSEC.M",
            "description": "Gross Borrowing Requirement: Breakdown by debt type: Federal securities",
            "search_patterns": [
                r"^\s*Federal securities\s*$",
                r"^\s*Bundeswertpapiere\s*$"
            ]
        },
        {
            "code": "GFB.BORR.BREAKDEPTTYPE.CONVFEDSEC.M",
            "description": "Gross Borrowing Requirement: Breakdown by debt type: Conventional Federal securities",
            "search_patterns": [
                r"Conventional Federal securities",
                r"Konventionelle Bundeswertpapiere"
            ]
        },
        {
            "code": "GFB.BORR.BREAKDEPTTYPE.FEDBOND.M",
            "description": "Gross Borrowing Requirement: Breakdown by debt type: Federal bonds",
            "search_patterns": [
                r"^\s*Federal bonds\s*$",
                r"^\s*Bundesanleihen\s*$"
            ]
        },
        {
            "code": "GFB.BORR.BREAKDEPTTYPE.30YFEDBOND.M",
            "description": "Gross Borrowing Requirement: Breakdown by debt type: 30-year Federal bonds",
            "search_patterns": [
                r"30-year Federal bonds",
                r"30-jährige Bundesanleihen"
            ]
        },
        {
            "code": "GFB.BORR.BREAKDEPTTYPE.15YFEDBOND.M",
            "description": "Gross Borrowing Requirement: Breakdown by debt type: 15-year Federal bonds",
            "search_patterns": [
                r"15-year Federal bonds",
                r"15-jährige Bundesanleihen"
            ]
        },
        {
            "code": "GFB.BORR.BREAKDEPTTYPE.10YFEDBOND.M",
            "description": "Gross Borrowing Requirement: Breakdown by debt type: 10-year Federal bonds",
            "search_patterns": [
                r"10-year Federal bonds",
                r"10-jährige Bundesanleihen"
            ]
        },
        {
            "code": "GFB.BORR.BREAKDEPTTYPE.7YFEDBOND.M",
            "description": "Gross Borrowing Requirement: Breakdown by debt type: 7-year Federal bonds",
            "search_patterns": [
                r"7-year Federal bonds",
                r"7-jährige Bundesanleihen"
            ]
        },
        {
            "code": "GFB.BORR.BREAKDEPTTYPE.FEDNOTE.M",
            "description": "Gross Borrowing Requirement: Breakdown by debt type: Federal notes",
            "search_patterns": [
                r"^\s*Federal notes\s*$",
                r"^\s*Bundesobligationen\s*$"
            ]
        },
        {
            "code": "GFB.BORR.BREAKDEPTTYPE.FEDTREASNOTE.M",
            "description": "Gross Borrowing Requirement: Breakdown by debt type: Federal Treasury notes",
            "search_patterns": [
                r"Federal Treasury notes",
                r"Bundesschatzanweisungen"
            ]
        },
        {
            "code": "GFB.BORR.BREAKDEPTTYPE.TREASDISCPAP.M",
            "description": "Gross Borrowing Requirement: Breakdown by debt type: Treasury discount paper",
            "search_patterns": [
                r"Treasury discount paper",
                r"Unverzinsliche Schatzanweisungen"
            ]
        },
        {
            "code": "GFB.BORR.BREAKDEPTTYPE.INFFEDSEC.M",
            "description": "Gross Borrowing Requirement: Breakdown by debt type: Inflation-linked Federal securities",
            "search_patterns": [
                r"Inflation-linked Federal securities",
                r"Inflationsindexierte Bundeswertpapiere"
            ]
        },
        {
            "code": "GFB.BORR.BREAKDEPTTYPE.GREENFEDSEC.M",
            "description": "Gross Borrowing Requirement: Breakdown by debt type: Green Federal securities",
            "search_patterns": [
                r"Green Federal securities",
                r"Grüne Bundeswertpapiere"
            ]
        },
        {
            "code": "GFB.BORR.BREAKDEPTTYPE.SUPSECFEDGOV.M",
            "description": "Gross Borrowing Requirement: Breakdown by debt type: Supplementary securities issued by the federal government",
            "search_patterns": [
                r"Supplementary securities issued by the federal government",
                r"Ergänzungswertpapiere des Bundes"
            ]
        },
        {
            "code": "GFB.BORR.BREAKDEPTTYPE.LOANSUPFEDGOV.M",
            "description": "Gross Borrowing Requirement: Breakdown by debt type: Loans from suppl. issues by the federal government for the ESF acc. to sect. 26b StFG",
            "search_patterns": [
                r"Loans from suppl.*federal government for the ESF.*26b",
                r"Darlehen aus Ergänzungsemissionen.*WSF.*26b"
            ]
        },
        {
            "code": "GFB.BORR.BREAKDEPTTYPE.ESFINVFEDGOVSEC.M",
            "description": "Gross Borrowing Requirement: Breakdown by debt type: ESF investment in federal government securities acc. to sect. 26b (5) StFG",
            "search_patterns": [
                r"ESF investment in federal government securities.*26b.*5",
                r"WSF-Anlage in Bundeswertpapieren.*26b.*5"
            ]
        },
        {
            "code": "GFB.BORR.BREAKDEPTTYPE.OTHERFEDSEC.M",
            "description": "Gross Borrowing Requirement: Breakdown by debt type: Other Federal securities",
            "search_patterns": [
                r"Other Federal securities",
                r"Sonstige Bundeswertpapiere"
            ]
        },
        {
            "code": "GFB.BORR.BREAKDEPTTYPE.PROMNOTE.M",
            "description": "Gross Borrowing Requirement: Breakdown by debt type: Promissory notes",
            "search_patterns": [
                r"^\s*Promissory notes\s*$",
                r"^\s*Schuldscheindarlehen\s*$"
            ]
        },
        {
            "code": "GFB.BORR.BREAKDEPTTYPE.OTHERLOANORDDEPT.M",
            "description": "Gross Borrowing Requirement: Breakdown by debt type: Other loans and ordinary debts",
            "search_patterns": [
                r"Other loans and ordinary debts",
                r"Sonstige Darlehen und gewöhnliche Schulden"
            ]
        },
        {
            "code": "GFB.BORR.OWNHOLD.M",
            "description": "Gross Borrowing Requirement: Own holdings",
            "search_patterns": [
                r"^\s*Own holdings\s*$",
                r"^\s*Eigenbestand\s*$"
            ]
        },
        {
            "code": "GFB.BORR.OWNHOLD.CONVFEDSEC.M",
            "description": "Gross Borrowing Requirement: Own holdings: Conventional Federal securities",
            "search_patterns": [
                r"^\s*Conventional Federal securities\s*$",
                r"^\s*Konventionelle Bundeswertpapiere\s*$"
            ],
            "parent_section": "Own holdings"
        },
        {
            "code": "GFB.BORR.OWNHOLD.30YFEDBOND.M",
            "description": "Gross Borrowing Requirement: Own holdings: 30-year Federal bonds",
            "search_patterns": [
                r"^\s*30-year Federal bonds\s*$",
                r"^\s*30-jährige Bundesanleihen\s*$"
            ],
            "parent_section": "Own holdings"
        },
        {
            "code": "GFB.BORR.OWNHOLD.15YFEDBOND.M",
            "description": "Gross Borrowing Requirement: Own holdings: 15-year Federal bonds",
            "search_patterns": [
                r"^\s*15-year Federal bonds\s*$",
                r"^\s*15-jährige Bundesanleihen\s*$"
            ],
            "parent_section": "Own holdings"
        },
        {
            "code": "GFB.BORR.OWNHOLD.10YFEDBOND.M",
            "description": "Gross Borrowing Requirement: Own holdings: 10-year Federal bonds",
            "search_patterns": [
                r"^\s*10-year Federal bonds\s*$",
                r"^\s*10-jährige Bundesanleihen\s*$"
            ],
            "parent_section": "Own holdings"
        },
        {
            "code": "GFB.BORR.OWNHOLD.7YFEDBOND.M",
            "description": "Gross Borrowing Requirement: Own holdings: 7-year Federal bonds",
            "search_patterns": [
                r"^\s*7-year Federal bonds\s*$",
                r"^\s*7-jährige Bundesanleihen\s*$"
            ],
            "parent_section": "Own holdings"
        },
        {
            "code": "GFB.BORR.OWNHOLD.FEDNOTE.M",
            "description": "Gross Borrowing Requirement: Own holdings: Federal notes",
            "search_patterns": [
                r"^\s*Federal notes\s*$",
                r"^\s*Bundesobligationen\s*$"
            ],
            "parent_section": "Own holdings"
        },
        {
            "code": "GFB.BORR.OWNHOLD.FEDTREASNOTE.M",
            "description": "Gross Borrowing Requirement: Own holdings: Federal Treasury notes",
            "search_patterns": [
                r"^\s*Federal Treasury notes\s*$",
                r"^\s*Bundesschatzanweisungen\s*$"
            ],
            "parent_section": "Own holdings"
        },
        {
            "code": "GFB.BORR.OWNHOLD.TREASDISCPAP.M",
            "description": "Gross Borrowing Requirement: Own holdings: Treasury discount paper",
            "search_patterns": [
                r"^\s*Treasury discount paper\s*$",
                r"^\s*Unverzinsliche Schatzanweisungen\s*$"
            ],
            "parent_section": "Own holdings"
        },
        {
            "code": "GFB.BORR.OWNHOLD.INFFEDSEC.M",
            "description": "Gross Borrowing Requirement: Own holdings: Inflation-linked Federal securities",
            "search_patterns": [
                r"^\s*Inflation-linked Federal securities\s*$",
                r"^\s*Inflationsindexierte Bundeswertpapiere\s*$"
            ],
            "parent_section": "Own holdings"
        },
        {
            "code": "GFB.BORR.OWNHOLD.GREENFEDSEC.M",
            "description": "Gross Borrowing Requirement: Own holdings: Green Federal securities",
            "search_patterns": [
                r"^\s*Green Federal securities\s*$",
                r"^\s*Grüne Bundeswertpapiere\s*$"
            ],
            "parent_section": "Own holdings"
        },
        {
            "code": "GFB.BORR.OWNHOLD.OTHERFEDSEC.M",
            "description": "Gross Borrowing Requirement: Own holdings: Other Federal securities",
            "search_patterns": [
                r"^\s*Other Federal securities\s*$",
                r"^\s*Sonstige Bundeswertpapiere\s*$"
            ],
            "parent_section": "Own holdings"
        },
        {
            "code": "GFB.BORR.ESFINVFEDGOVSEC.M",
            "description": "Gross Borrowing Requirement: ESF investment in federal government securities acc. to sect. 26b (5) StFG",
            "search_patterns": [
                r"^ESF investment in federal government securities.*26b.*5",
                r"^WSF-Anlage in Bundeswertpapieren.*26b.*5"
            ]
        }
    ]

    # Create 54 REDEMPTION measures by copying and modifying borrowing measures
    redemption_measures = []
    for measure in borrowing_measures:
        redemption_measure = {
            "code": measure["code"].replace("BORR", "REDEM"),
            "description": measure["description"].replace("Gross Borrowing Requirement:", "Redemption Payments:").replace(" Financing Federal budget", "Redemption Payments: Financing Federal budget"),
            "search_patterns": measure["search_patterns"]  # Same patterns work for both sheets
        }
        # Copy parent_section if it exists
        if "parent_section" in measure:
            redemption_measure["parent_section"] = measure["parent_section"]
        redemption_measures.append(redemption_measure)

    return borrowing_measures, redemption_measures


def build_config(source_file):
    """
    Build the JSON configuration by analyzing the Excel file

    Args:
        source_file: Path to source Excel file

    Returns:
        Configuration dictionary
    """
    print("=" * 70)
    print("GFB Config Builder - Context-Based Mapping")
    print("=" * 70)
    print(f"\nSource file: {source_file}")

    # Read sheets
    print("\n1. Reading Excel sheets...")
    borrowing_df = pd.read_excel(source_file, sheet_name='rpgBorrowing', header=None)
    redemption_df = pd.read_excel(source_file, sheet_name='rpgRedemptions', header=None)
    print(f"   OK Borrowing: {borrowing_df.shape[0]} rows x {borrowing_df.shape[1]} cols")
    print(f"   OK Redemption: {redemption_df.shape[0]} rows x {redemption_df.shape[1]} cols")

    # Find date rows
    print("\n2. Finding date rows...")
    borrowing_date_row = find_date_row(borrowing_df)
    redemption_date_row = find_date_row(redemption_df)

    # Get measure definitions
    print("\n3. Loading fixed output column definitions...")
    borrowing_measures, redemption_measures = create_measure_definitions()
    print(f"   OK {len(borrowing_measures)} borrowing measures")
    print(f"   OK {len(redemption_measures)} redemption measures")

    # Map each measure to its source row
    print("\n4. Mapping borrowing measures to source rows...")
    borrowing_mappings = []
    for i, measure in enumerate(borrowing_measures, 1):
        # Check if this measure has a parent section requirement
        parent_section = measure.get("parent_section", None)

        row_idx = find_text_in_sheet(
            borrowing_df,
            measure["search_patterns"],
            start_row=borrowing_date_row + 1,
            parent_section=parent_section
        )

        if row_idx is not None:
            label = clean_text(borrowing_df.iloc[row_idx, 0])
            print(f"   {i:2d}. OK Row {row_idx+1:3d}: {label[:60]}")
        else:
            label = None
            print(f"   {i:2d}. NOT FOUND: {measure['code']}")

        borrowing_mappings.append({
            "code": measure["code"],
            "description": measure["description"],
            "search_patterns": measure["search_patterns"],
            "source_row": row_idx,
            "source_label": label
        })

    print("\n5. Mapping redemption measures to source rows...")
    redemption_mappings = []
    for i, measure in enumerate(redemption_measures, 1):
        # Check if this measure has a parent section requirement
        parent_section = measure.get("parent_section", None)

        row_idx = find_text_in_sheet(
            redemption_df,
            measure["search_patterns"],
            start_row=redemption_date_row + 1,
            parent_section=parent_section
        )

        if row_idx is not None:
            label = clean_text(redemption_df.iloc[row_idx, 0])
            print(f"   {i:2d}. OK Row {row_idx+1:3d}: {label[:60]}")
        else:
            label = None
            print(f"   {i:2d}. NOT FOUND: {measure['code']}")

        redemption_mappings.append({
            "code": measure["code"],
            "description": measure["description"],
            "search_patterns": measure["search_patterns"],
            "source_row": row_idx,
            "source_label": label
        })

    # Create config
    config = {
        "version": "1.0",
        "created": datetime.now().isoformat(),
        "source_file": os.path.basename(source_file),
        "borrowing_sheet": {
            "name": "rpgBorrowing",
            "date_row": borrowing_date_row,
            "date_column_start": 1,
            "measures": borrowing_mappings
        },
        "redemption_sheet": {
            "name": "rpgRedemptions",
            "date_row": redemption_date_row,
            "date_column_start": 1,
            "measures": redemption_mappings
        }
    }

    # Summary
    print("\n" + "=" * 70)
    print("Mapping Summary")
    print("=" * 70)
    borr_found = sum(1 for m in borrowing_mappings if m["source_row"] is not None)
    redem_found = sum(1 for m in redemption_mappings if m["source_row"] is not None)
    print(f"Borrowing: {borr_found}/{len(borrowing_mappings)} measures mapped")
    print(f"Redemption: {redem_found}/{len(redemption_mappings)} measures mapped")
    print(f"Total: {borr_found + redem_found}/{len(borrowing_mappings) + len(redemption_mappings)} measures mapped")

    if borr_found < len(borrowing_mappings) or redem_found < len(redemption_mappings):
        print("\n\! WARNING: Some measures were not found!")
        print("This may indicate:")
        print("  - Search patterns need adjustment")
        print("  - Excel structure has changed")
        print("  - Measures don't exist in this version")

    return config


def save_config(config, output_file="gfb_config.json"):
    """Save config to JSON file"""
    print(f"\nSaving configuration to {output_file}...")

    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(config, f, indent=2, ensure_ascii=False)

    file_size = os.path.getsize(output_file) / 1024
    print(f"OK Configuration saved!")
    print(f"  File: {output_file}")
    print(f"  Size: {file_size:.1f} KB")


if __name__ == "__main__":
    # Find source Excel file
    import glob

    excel_files = glob.glob("gfb_downloads/*.xlsx")
    if not excel_files:
        excel_files = glob.glob("**/*.xlsx", recursive=True)
        excel_files = [f for f in excel_files if not f.startswith('GFB_DATA_')]

    if not excel_files:
        print("ERROR: No source Excel files found!")
        exit(1)

    source_file = excel_files[0]

    # Build and save config
    config = build_config(source_file)
    save_config(config, "gfb_config.json")

    print("\nOK Config generation complete!")
    print("\nNext step: Run map.py (it will automatically use gfb_config.json)")
