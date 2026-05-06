"""
Seed categories.json with known PO → standardized-name mappings,
then write a 'PO Name Standardized' column into every vessel sheet.
"""

import sys
import openpyxl
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from classifier import bulk_train, classify

# ── Training pairs: (raw pattern, standardized name) ──────────────────────────
TRAINING = [
    # Lubricants
    ("Lubricants - Pre Delivery",                          "Lubricants"),
    ("Pre Delivery - Additional Grade Copco Roto-XtendDuty","Lubricants"),
    ("Predelivery - Additional Grades",                    "Lubricants"),
    ("Predelivery Lubes - Order # 2 (5713853485)",         "Lubricants"),
    ("Pre Delivery Lubes Order# 3 (5714478288)",           "Lubricants"),
    ("Pre Delivery - Lubes Order# 3 (5715069528)",         "Lubricants"),
    ("Pre-Delivery - Lubes",                               "Lubricants"),
    ("Final Sailing Qty of Lubes",                         "Lubricants"),
    ("PD Lubes - VLSQ61941 - Kaeser Fluid Rotary S-46X",  "Lubricants"),
    ("RSC Lube oil - qty / Total for Glovis Lighthouse",   "Lubricants"),
    ("RSC Lube oil - qty",                                 "Lubricants"),
    ("lube oil for Sailing quantity",                      "Lubricants"),
    ("Predelivery - Visga 46 5715010143",                  "Lubricants"),
    ("Glycol",                                             "Lubricants"),
    ("PD - RSC Envirologic",                               "Lubricants"),
    ("AST PO",                                             "Lubricants"),
    ("FS -AST-LANDER",                                     "Lubricants"),
    ("FS -AST-LINK",                                       "Lubricants"),
    ("FS-AST-SAMPLE",                                      "Lubricants"),
    ("FS-AST-LOGOS",                                       "Lubricants"),
    ("FS-AST-LOYOL",                                       "Lubricants"),

    # IT Equipment - ITBP
    ("FS-ITBP  -lander",                                   "IT Equipment - ITBP"),
    ("FS-ITBP  -luminous",                                 "IT Equipment - ITBP"),
    ("FS-ITBP  -glovis lighthouse",                        "IT Equipment - ITBP"),
    ("FS-ITBP  -glovis link",                              "IT Equipment - ITBP"),
    ("FS-ITBP  -LOGOS",                                    "IT Equipment - ITBP"),
    ("FS-ITBP  -LOYAL",                                    "IT Equipment - ITBP"),
    ("FS-ITBP  -1",                                        "IT Equipment - ITBP"),

    # IT Equipment - ITBP Additional Items
    ("FS-ITBP Additional Item",                            "IT Equipment - ITBP Additional Items"),

    # IT Equipment - ITBP Lab Kit
    ("FS-ITBP MISC LAB KIT",                               "IT Equipment - ITBP Lab Kit"),

    # IT Equipment - ITBP Cables
    ("FS-ITBP CABLES",                                     "IT Equipment - ITBP Cables"),

    # IT Equipment - ITBP Licenses
    ("FS-ITBP-LICENSES",                                   "IT Equipment - ITBP Licenses"),

    # IT Equipment - ITBP Tax Invoice
    ("FS-ITBP tax invoice",                                "IT Equipment - ITBP Tax Invoice"),

    # IT Equipment - X10 System
    ("FS-X10-LANDER",                                      "IT Equipment - X10 System"),
    ("FS-X10-SILVERPEAK",                                  "IT Equipment - X10 System"),
    ("FS-X10",                                             "IT Equipment - X10 System"),
    ("fs-x10 sample test",                                 "IT Equipment - X10 System"),

    # IT Equipment - Computers (Lenovo)
    ("FS-LENOVO-Lander",                                   "IT Equipment - Computers (Lenovo)"),
    ("FS-LENOVO-Luminous",                                  "IT Equipment - Computers (Lenovo)"),
    ("FS-LENOVO-Lighthouse",                               "IT Equipment - Computers (Lenovo)"),
    ("FS-LENOVO-Link",                                     "IT Equipment - Computers (Lenovo)"),
    ("FS-LENOVO-Logos",                                    "IT Equipment - Computers (Lenovo)"),
    ("FS-LENOVO-Loyal",                                    "IT Equipment - Computers (Lenovo)"),

    # IT Equipment - Interface / Integration
    ("Interface PO -24110005",                             "IT Equipment - Interface / Integration"),
    ("Interface PO",                                       "IT Equipment - Interface / Integration"),
    ("Interface PO-H1612",                                 "IT Equipment - Interface / Integration"),
    ("FS-Interface-Lander",                                "IT Equipment - Interface / Integration"),
    ("FS-Interface-Light house",                           "IT Equipment - Interface / Integration"),
    ("FS-Interface MTA adapter",                           "IT Equipment - Interface / Integration"),
    ("Interface -equipment-lighthouse",                    "IT Equipment - Interface / Integration"),
    ("fs interface",                                       "IT Equipment - Interface / Integration"),

    # IT Equipment - Sihui Marine
    ("FS-SIHUI-lander",                                    "IT Equipment - Sihui Marine"),
    ("FS-SIHUI",                                           "IT Equipment - Sihui Marine"),
    ("FS-SIHUI-link",                                      "IT Equipment - Sihui Marine"),
    ("Sihui Marine-light house",                           "IT Equipment - Sihui Marine"),

    # Communications - VSAT (KVH)
    ("FS-KVH-LANDER UPDATED",                             "Communications - VSAT (KVH)"),
    ("FS-KVH-GSI",                                        "Communications - VSAT (KVH)"),
    ("FS-KVH -LIGHTHOUSE UPDATED",                        "Communications - VSAT (KVH)"),
    ("FS-KVH-LIGHTHOUSE",                                 "Communications - VSAT (KVH)"),
    ("FS-KVH-SWS",                                        "Communications - VSAT (KVH)"),
    ("FS-KVH-SWS-LOGOS",                                  "Communications - VSAT (KVH)"),
    ("FS-KVH-SWS-LOYAL",                                  "Communications - VSAT (KVH)"),

    # Communications - VLINK
    ("FS-VLINK-2411005",                                   "Communications - VLINK"),
    ("FS-VLINK-H1610",                                     "Communications - VLINK"),
    ("FS-VLINK-H1612",                                     "Communications - VLINK"),
    ("FS-VLINK-Accessories",                               "Communications - VLINK"),
    ("FS-VLINK CABLE",                                     "Communications - VLINK"),

    # Communications - Starlink Accessories
    ("FS-VLINK-Starlink Accessories",                      "Communications - Starlink Accessories"),

    # LNG Operation Manual
    ("Service: LNG OPERATION MANUAL Development",          "LNG Operation Manual"),
    ("LNG OPERATION MANUAL Development",                   "LNG Operation Manual"),

    # Tripods & Spray Pipes
    ("Shipyard Extra Buyers Supply - Tripods and Spray Pipes", "Tripods & Spray Pipes"),
    ("Pre Delivery - Spray Pipes",                         "Tripods & Spray Pipes"),

    # QCDC
    ("Pre Delivery - QCDC",                                "QCDC"),
    ("Shipyard Extra (buyer Supply) - QCDC -GSI",          "QCDC"),
    ("QC DC Supply",                                       "QCDC"),
    ("QC DC",                                              "QCDC"),

    # Gas Detection System
    ("Reliabiltiy Spares - Gas Detection System",          "Gas Detection System"),
    ("Reliability Spares - Gas Detection System",          "Gas Detection System"),

    # Navigation - HiNAS SVM
    ("Shipyard Extra Buyer Supply - Hinas SVM SEASPAN 10.8K PCTC",  "Navigation - HiNAS SVM"),
    ("Shipyard Extra Buyers Supply - Hinas SVM SEASPAN 10.8K PCTC", "Navigation - HiNAS SVM"),
    ("Shipyard Extra Buyers supply - Hinas SVM SEASPAN 10.8K PCTC", "Navigation - HiNAS SVM"),
    ("Shipyard Extra Buyer Supply - HiNAS SVM 10800 A-Q-2025-09-02_R0", "Navigation - HiNAS SVM"),
    ("HiNAS SVM (Surround View Monitoring)",               "Navigation - HiNAS SVM"),
    ("New Positioning System for Transiting Neopanamax - Buyer Supply", "Navigation - HiNAS SVM"),
    ("Buyer Supply - New Positioning System for Transiting Panama ( inc. Anti Jamming) (SAAB R6 NAV NEO)", "Navigation - HiNAS SVM"),

    # Navigation - Captains Eye
    ("Shipyard Extra Owner supply - Captains Eye",         "Navigation - Captains Eye"),
    ("Shipyard Extra (owner supply) - Captains Eye",       "Navigation - Captains Eye"),
    ("Captain eye",                                        "Navigation - Captains Eye"),

    # Thermal Camera (FLIR)
    ("FLIR ESPro and E6 Pro",                              "Thermal Camera (FLIR)"),
    ("Pre Delivery - FLIR Camera",                         "Thermal Camera (FLIR)"),
    ("Two (2) explosion-proof portable thermal imaging cameras", "Thermal Camera (FLIR)"),
    ("8. Explosion Proof Camera",                          "Thermal Camera (FLIR)"),
    ("Explosion Proof Camera",                             "Thermal Camera (FLIR)"),

    # Fire Detection System
    ("Reliability Spares - Fire detection system",         "Fire Detection System"),

    # Sample Cooler
    ("Pre Delivery - Sample Cooler",                       "Sample Cooler"),
    ("Sampling Cooler- Qty?",                              "Sample Cooler"),

    # Cryogenic PPE
    ("Pre Delivery - Cryogenic PPE",                       "Cryogenic PPE"),
    ("Cryogenic PPE - Scandia",                            "Cryogenic PPE"),

    # LNG Bunkering Equipment
    ("Pre Delivery - LNG Bunkering Equipment",             "LNG Bunkering Equipment"),

    # Seatab
    ("Pre Delivery - Seatab",                              "Seatab"),
    ("Seatab",                                             "Seatab"),

    # Garbage Compactor
    ("Pre Delivery - Garbage Compactor",                   "Garbage Compactor"),
    ("Garbage Compactor",                                  "Garbage Compactor"),
    ("Garbage compactor in garbage room and engine room",  "Garbage Compactor"),

    # Oil Skimmers - Abanaki
    ("SHIPYARD EXTRA OWNER SUPPLY- Abanaki Skimmers",      "Oil Skimmers - Abanaki"),
    ("SHIPYARD EXTRA ONWER SUPPLY - Abanaki Skimmers",     "Oil Skimmers - Abanaki"),
    ("Oil Skimmers for engine room",                       "Oil Skimmers - Abanaki"),

    # Boiler Spares
    ("Reliability Spares - PCTC-Seaspan - Boilers QUOT26012103",      "Boiler Spares"),
    ("Reliability Spares - PCTC-Seaspan - Boilers QUOT26012103_R00",  "Boiler Spares"),
    ("Reliability Spares - 10,8K CEU SWS - GVT- Korval",              "Boiler Spares"),
    ("Reliability Spares - 10,8K CEU SWS - GVT- Eltronic",            "Boiler Spares"),

    # Ultrasonic Cleaner
    ("Ultrasonic Cleaning system - 230V",                  "Ultrasonic Cleaner"),
    ("Ultrasonic Cleaner",                                 "Ultrasonic Cleaner"),

    # OPA 90 Kit
    ("OPA 90 KIT - 12-barrel kit",                         "OPA 90 Kit"),
    ("Pre-Delivery - OPA 90",                              "OPA 90 Kit"),

    # High Pressure Washer
    ("H.P machine - 2pcs (prefer Branded - Karcher/ KEW)", "High Pressure Washer"),
    ("Reliability HP Pump Skid",                           "High Pressure Washer"),
    ("HP Pump Skid",                                       "High Pressure Washer"),
    ("Air driven HP Pump - For hydraulic nuts for M/E",    "High Pressure Washer"),
    ("13. Air driven HP Pump - For hydraulic nuts for M/E - From Dan Marine or VLINK", "High Pressure Washer"),

    # Steam Splash Suit
    ("3. Steam Splash Suit",                               "Steam Splash Suit"),
    ("Steam Splash Suit",                                  "Steam Splash Suit"),

    # Gas Meters
    ("Gas meters",                                         "Gas Meters"),
    ("Gas Meter",                                          "Gas Meters"),
    ("Gas meters",                                         "Gas Meters"),

    # Oily Water Separator
    ("Oily Water Separator",                               "Oily Water Separator"),

    # BWTS Spares
    ("BWTS Spares",                                        "BWTS Spares"),
    ("Manuals for all equipment- priority for BWTS, STP, OWS, EMGCY AC, EMCY GEN", "Publications & Manuals"),

    # EV Lance
    ("EV Lance",                                           "EV Lance"),

    # RO-RO Equipment - Loose Fittings
    ("PCTC Loose Fittings Procurement",                    "RO-RO Equipment - Loose Fittings"),
    ("Jonghap - PCTC Loose Fittings Procurement",          "RO-RO Equipment - Loose Fittings"),
    ("Item no 11, loose equipment for emergency operation of roro equipment", "RO-RO Equipment - Loose Fittings"),

    # RO-RO Equipment - Spares
    ("RO-RO Equipment (Main & Centre Ramp spares + inner Ramp spares)", "RO-RO Equipment - Spares"),
    ("Reliability RO-RO Equipment (Main & Centre Ramp spares + inner Ramp spares)", "RO-RO Equipment - Spares"),

    # Engine Spares - Main Engine
    ("Main Engine Spares",                                 "Engine Spares - Main Engine"),
    ("Main Engine - Additional Spares",                    "Engine Spares - Main Engine"),
    ("Main Engine",                                        "Engine Spares - Main Engine"),

    # Engine Spares - Auxiliary Engine
    ("Aux. Engine Spares",                                 "Engine Spares - Auxiliary Engine"),
    ("Aux. Engine",                                        "Engine Spares - Auxiliary Engine"),

    # Sewage Treatment Plant
    ("Sewage Traetment Plant",                             "Sewage Treatment Plant"),
    ("Sewage Treatment Plant",                             "Sewage Treatment Plant"),

    # Sewage Discharge Pump
    ("Sewage Discharge P/P",                               "Sewage Discharge Pump"),

    # Mooring Ropes
    ("Mooring Rope",                                       "Mooring Ropes"),

    # Citadel VHF/GPS/UHF System
    ("Citadel VHF/GPS and UHF ANTENNA system",             "Citadel VHF/GPS/UHF System"),
    ("Buyer Supply - 10,800 PCTC Citadel VHF/GPS and UHF ANTENNA system", "Citadel VHF/GPS/UHF System"),
    ("Citadel Items",                                      "Citadel VHF/GPS/UHF System"),

    # UHF Radio
    ("UHF Radio",                                          "UHF Radio"),
    ("UHF Radio (Handheld)",                               "UHF Radio"),
    ("UHF radios",                                         "UHF Radio"),
    ("UHF Antenna",                                        "UHF Radio"),

    # Welding Equipment
    ("Pre Delivery - Welding Equipment",                   "Welding Equipment"),
    ("Welding machine in Workshop 440 v",                  "Welding Equipment"),

    # Engineering Products & Chemicals
    ("Pre Delivery - Engineering Products & Kits (DREW)",  "Engineering Products & Chemicals"),

    # Dual Fuel Tools
    ("DF Tools",                                           "Dual Fuel Tools"),
    ("DF Tools for Methanol Project",                      "Dual Fuel Tools"),

    # Safety Integration - SMiG
    ("Shipyard Extra (Owner supply) - 10800 CEU LNG PCTC _x0013_ Consilium SMiG Integration", "Safety Integration - SMiG"),
    ("SHIPYARD EXTRA OWNER SUPPLY - SWS-10.8k CEU extra cost for SMiG application",            "Safety Integration - SMiG"),
    ("SHIPYARD EXTRA OWNER SUPLPY - 10800 CEU LNG PCTC _x0013_ Consilium SMiG Integration",   "Safety Integration - SMiG"),
    ("Shipyard Extra (owner supply) 10800 CEU LNG PCTC _x0013_ Consilium SMiG Integration",   "Safety Integration - SMiG"),

    # Nitrogen Generator
    ("Reliability Nitrogen Generator",                     "Nitrogen Generator"),
    ("Nitrogen Generator",                                 "Nitrogen Generator"),

    # Training - Wartsila
    ("Wartsila Training PCTC",                             "Training - Wartsila"),

    # Fleet Systems - Services
    ("FS services",                                        "Fleet Systems - Services"),
    ("FS",                                                 "Fleet Systems - Services"),

    # Fleet Systems - General
    ("FS-VLINK-Accessories",                               "Fleet Systems - General"),

    # Treatment Chemicals
    ("Pre Delivery - Treatment Chemical",                  "Treatment Chemicals"),

    # Publications & Manuals
    ("Publications",                                       "Publications & Manuals"),
    ("Manuals for all equipment- priority for BWTS, STP, OWS, EMGCY AC, EMCY GEN", "Publications & Manuals"),

    # Refrigerant Equipment
    ("Refrigerant Equipment",                              "Refrigerant Equipment"),
    ("Reefer cable and plugs extensions",                  "Refrigerant Equipment"),

    # BOG Compressor Spares
    ("Reliability - Donghwa Spares for BOG Compressor",    "BOG Compressor Spares"),

    # IMPA Stores - BA Compressor
    ("Impa Stores - BA Compressor",                        "IMPA Stores - BA Compressor"),
    ("IMPA Stores - BA Compressor",                        "IMPA Stores - BA Compressor"),
    ("Breathing air compressor",                           "IMPA Stores - BA Compressor"),
    ("BA compressor",                                      "IMPA Stores - BA Compressor"),

    # IMPA Stores - Engine
    ("Impa Stores - Engine",                               "IMPA Stores - Engine"),
    ("IMPA Stores - Engine",                               "IMPA Stores - Engine"),
    ("IMPA Stores - Engine 1",                             "IMPA Stores - Engine"),

    # IMPA Stores - Galley
    ("Impa Stores - Galley",                               "IMPA Stores - Galley"),
    ("IMPA Stores - Galley",                               "IMPA Stores - Galley"),

    # IMPA Stores - Accommodation
    ("Impa Stores - ACC",                                  "IMPA Stores - Accommodation"),
    ("IMPA Stores - ACC",                                  "IMPA Stores - Accommodation"),

    # IMPA Stores - Bridge
    ("Impa Stores - Bridge",                               "IMPA Stores - Bridge"),
    ("IMPA Stores - Bridge",                               "IMPA Stores - Bridge"),

    # IMPA Stores - Gym
    ("Impa Stores - Gym",                                  "IMPA Stores - Gym"),
    ("IMPA Stores - Gym",                                  "IMPA Stores - Gym"),

    # IMPA Stores - Crew
    ("Impa Stores - Crew",                                 "IMPA Stores - Crew"),
    ("IMPA Stores - Crew",                                 "IMPA Stores - Crew"),

    # IMPA Stores - FFE (Fire Fighting Equipment)
    ("Impa Stores - FFE",                                  "IMPA Stores - FFE"),
    ("IMPA Stores - FFE",                                  "IMPA Stores - FFE"),

    # IMPA Stores - Company Stationery
    ("Impa Stores - Company Stationery",                   "IMPA Stores - Company Stationery"),
    ("IMPA Stores - Company Stationery",                   "IMPA Stores - Company Stationery"),

    # IMPA Stores - Deck
    ("Impa Stores - Deck",                                 "IMPA Stores - Deck"),
    ("IMPA Stores - Deck",                                 "IMPA Stores - Deck"),

    # IMPA Stores - SOPEP
    ("Impa Stores - SOPEP",                                "IMPA Stores - SOPEP"),
    ("IMPA Stores - SOPEP",                                "IMPA Stores - SOPEP"),

    # IMPA Stores - Electric
    ("Impa Stores - Electric",                             "IMPA Stores - Electric"),
    ("IMPA Stores - Electric",                             "IMPA Stores - Electric"),

    # IMPA Stores - Instrumentation
    ("Impa Stores - Instrumenation",                       "IMPA Stores - Instrumentation"),
    ("IMPA Stores - Instrumenation",                       "IMPA Stores - Instrumentation"),

    # IMPA Stores - Stationery
    ("Impa Stores Stationery",                             "IMPA Stores - Stationery"),
    ("IMPA Stores Stationery",                             "IMPA Stores - Stationery"),
    ("12. Standard Chinese stationery",                    "IMPA Stores - Stationery"),

    # Temperature Calibrators
    ("Temperature Calibrators",                            "Temperature Calibrators"),
    ("10. Temperature & Pressure Calibrators",             "Temperature Calibrators"),

    # Pressure Calibrators
    ("Pressure Calibrators",                               "Pressure Calibrators"),

    # Maintenance Chemicals
    ("Maintenance Chemicals",                              "Maintenance Chemicals"),
    ("Chemical",                                           "Maintenance Chemicals"),
    ("Naoh",                                               "Maintenance Chemicals"),

    # Industrial Gases
    ("Gases",                                              "Industrial Gases"),
    ("Gas Cylinders",                                      "Industrial Gases"),
    ("Industrial Gases",                                   "Industrial Gases"),

    # Sea Trial
    ("Seatrial Stores",                                    "Sea Trial"),
    ("Seatrial",                                           "Sea Trial"),

    # Safety & PPE
    ("2. LOTO Stickers",                                   "Safety & PPE"),
    ("LOTO Stickers",                                      "Safety & PPE"),
    ("4. Chemical tags",                                   "Safety & PPE"),
    ("Chemical tags",                                      "Safety & PPE"),
    ("5. Trelleborg Safe lanes",                           "Safety & PPE"),
    ("Trelleborg Safe lanes",                              "Safety & PPE"),
    ("Immersion Suit",                                     "Safety & PPE"),
    ("Fire Mans Outfit",                                   "Safety & PPE"),
    ("PPE - RMS",                                          "Safety & PPE"),
    ("Rubber Matting",                                     "Safety & PPE"),

    # VISWA / Line Samplers
    ("11. VISWA Line Samplers",                            "VISWA Line Samplers"),
    ("VISWA Line Samplers",                                "VISWA Line Samplers"),

    # Crew Care / Provisions
    ("Crew Care",                                          "Crew Care & Provisions"),
    ("Provisions and Bonds",                               "Crew Care & Provisions"),
    ("Medicines",                                          "Crew Care & Provisions"),
    ("Medox and Equipments",                               "Crew Care & Provisions"),
    ("Domestic refrigerators 200 litres for rec rooms to be added to stores supplies", "Crew Care & Provisions"),

    # Miscellaneous / Shipyard Equipment
    ("Stanchions",                                         "Shipyard / Miscellaneous"),
    ("Stancions",                                          "Shipyard / Miscellaneous"),
    ("Manual chain blocks",                                "Shipyard / Miscellaneous"),
    ("Service Car and Forlift",                            "Shipyard / Miscellaneous"),
    ("Service Car",                                        "Shipyard / Miscellaneous"),
    ("Forlift",                                            "Shipyard / Miscellaneous"),
    ("Alcometer",                                          "Shipyard / Miscellaneous"),
    ("Bosch GLM-50 or Laser Distance Measuring Tool",      "Shipyard / Miscellaneous"),
    ("Deck Sweeper",                                       "Shipyard / Miscellaneous"),
    ("Suez light",                                         "Shipyard / Miscellaneous"),
    ("suez light",                                         "Shipyard / Miscellaneous"),
    ("Soot Release Dosing Device- Qty?",                   "Shipyard / Miscellaneous"),
    ("Soot Releasing Dosing Unit",                         "Shipyard / Miscellaneous"),
    ("Soot",                                               "Shipyard / Miscellaneous"),
    ("Ozone sterilizer",                                   "Shipyard / Miscellaneous"),
    ("Ozone sterliser",                                    "Shipyard / Miscellaneous"),
    ("ozone sterliser",                                    "Shipyard / Miscellaneous"),
    ("Garbage Refrigerator",                               "Shipyard / Miscellaneous"),
    ("Garbage refrigerator in dry provision store",        "Shipyard / Miscellaneous"),
    ("Electric Hoist",                                     "Shipyard / Miscellaneous"),
    ("Receptacles",                                        "Shipyard / Miscellaneous"),
    ("Two (2) receptacles in officer's mess room, crew's mess room for electric wall type fan", "Shipyard / Miscellaneous"),
    ("One (1) receptacle in public space (conference room, gymnasium, galley, crew's recreation room, offcier's recreation room, ship office) for electric wall type fan.", "Shipyard / Miscellaneous"),
    ("Cable for electrical power shore connection",        "Shipyard / Miscellaneous"),
    ("Nitrogen Hose",                                      "Shipyard / Miscellaneous"),
    ("Thermal Stop PO",                                    "Shipyard / Miscellaneous"),
    ("Smart Docking system (Buyer Supply) -",              "Shipyard / Miscellaneous"),
    ("ORCA-AI",                                            "Shipyard / Miscellaneous"),
    ("Iron Test Kit from CMT",                             "Shipyard / Miscellaneous"),
    ("Oil Test Kit",                                       "Shipyard / Miscellaneous"),
    ("Oxygen fastening rack and strap",                    "Shipyard / Miscellaneous"),

    # Citadel (already handled above but keep UHF separately)

    # Test kit / calibration (iron, oil)

    # Additional miscellaneous from Glovis Lander
    ("2. LOTO Stickers",                                   "Safety & PPE"),
]

# ── Override for items that should NOT be "Fleet Systems - General"
# (We over-assigned FS-VLINK-Accessories above; correct it)
OVERRIDES = [
    ("FS-VLINK-Accessories",                               "Communications - VLINK"),
]


def seed():
    pairs = TRAINING + OVERRIDES
    print("Seeding categories.json …")
    bulk_train(pairs, silent=True)


def apply_to_workbook(wb_path: str):
    wb = openpyxl.load_workbook(wb_path)

    vessel_sheets = [
        "Glovis Lander", "Glovis Luminous", "Glovis Lighthouse",
        "Glovis Link", "Glovis Logos", "Glovis Loyal",
    ]

    HEADER_ROW = 10   # row containing column labels
    DATA_START  = 11  # first data row
    PO_NAME_COL = 3   # column C (1-indexed)

    for sheet_name in vessel_sheets:
        if sheet_name not in wb.sheetnames:
            print(f"  Sheet not found: {sheet_name}")
            continue

        ws = wb[sheet_name]

        # Find or create 'PO Name Standardized' header in row 10
        std_col = None
        for cell in ws[HEADER_ROW]:
            if cell.value and "standardized" in str(cell.value).lower():
                std_col = cell.column
                break

        if std_col is None:
            # Place it right after PO Name (col C = 3), so col D = 4
            # But col D may be merged/used; use col 4 and bump if needed
            std_col = PO_NAME_COL + 1  # column D
            ws.cell(row=HEADER_ROW, column=std_col, value="PO Name Standardized")

        classified = 0
        unmatched = []

        for row_idx in range(DATA_START, ws.max_row + 1):
            po_name = ws.cell(row=row_idx, column=PO_NAME_COL).value
            if not po_name or not isinstance(po_name, str):
                continue

            category, standardized, confidence = classify(po_name.strip())
            ws.cell(row=row_idx, column=std_col, value=standardized)

            if category is None:
                unmatched.append((row_idx, po_name.strip(), confidence))
            else:
                classified += 1

        print(f"\n{sheet_name}: {classified} classified, {len(unmatched)} unmatched")
        for row_idx, name, conf in unmatched:
            print(f"  row {row_idx:>4}  [{conf:.0%}]  {name}")

    wb.save(wb_path)
    print(f"\nSaved: {wb_path}")


if __name__ == "__main__":
    seed()
    apply_to_workbook("/home/user/fuzzy-happiness/Newbuilds_workingcopy.xlsx")
