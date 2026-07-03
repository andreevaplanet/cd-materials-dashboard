#!/usr/bin/env python3
"""
Convert the INFCOM CD Materials Excel workbook into materials.json for the dashboard.

Usage:
    python build_materials_json.py New_CD_dashboard.xlsx materials.json

Then place materials.json next to INFCOM_CD_Materials_Dashboard.html on the web
server. The dashboard loads it automatically on page load; no HTML edits needed.

Reads the 'Classified_materials' sheet, columns: Title, Type, Domain, Function, Link.
If a row's Link cell is filled in the workbook it is used as-is; otherwise a known
link from LINK_MAP below is applied. Add links straight into the workbook (or into
LINK_MAP) and re-run to refresh.

Requires: pandas, openpyxl   ->   pip install pandas openpyxl
"""
import sys, re, json
import pandas as pd

SHEET = "Classified_materials"

# Function value -> (display area name, spectrum colour, short code)
AREA_MAP = {
 "Information/data sharing and management": ("Information Management & Technology", "#1E5AA8", "IMT"),
 "Observing systems and monitoring networks": ("Earth Observing Systems & Networks", "#0E8A6B", "EOS"),
 "Satellite applications": ("Satellite Data Applications", "#4C4BB0", "SAT"),
 "Measurements, instruments and calibration": ("Measurements, Instrumentation & Traceability", "#C77D0A", "MIT"),
 "Integrated prediction systems/forecasting": ("Data-Processing & Prediction", "#7E3FA0", "DPP"),
 "Radio Frequency": ("Radio Frequency Coordination", "#C0335F", "RFC"),
}

# Recovered links (title -> URL), harvested from the published INFCOM dashboard.
# Applied only when the workbook's own Link cell is blank.
LINK_MAP = {
 "WIS Newsletter":"https://community.wmo.int/node/33669",
 "Operational Newsletter":"https://community.wmo.int/node/33765",
 "WIS Manuals":"https://library.wmo.int/index.php?lvl=notice_display&id=9254",
 "Guide to WIS":"https://library.wmo.int/idurl/4/28988",
 "WIS2 in a Box":"https://docs.wis2box.wis.wmo.int/en/latest/",
 "WIS2 Training":"https://community.wmo.int/node/37819",
 "WIGOS Newsletters":"https://community.wmo.int/node/33511",
 "Aircraft based Observation and AMDAR Newsletters":"https://community.wmo.int/node/37133",
 "WIGOS Manual":"https://library.wmo.int/?lvl=notice_display&id=19223",
 "Guide to WIGOS":"https://library.wmo.int/index.php?lvl=notice_display&id=20026",
 "WIGOS Metadata Standard":"https://library.wmo.int/index.php?lvl=more_results&mode=extended",
 "Technical Guidelines for RWC on WDQMS (WMO-No. 1224)":"https://library.wmo.int/idurl/4/56347",
 "High-level Guidance on Evolution of Global Observing Systems 2023-2027 (WMO-No. 1334)":"https://library.wmo.int/idurl/4/68862",
 "WIGOS Technical Reports":"https://community.wmo.int/activity-areas/WIGOS/Technical%20reports",
 "OSCAR/Surface User Manual":"https://library.wmo.int/idurl/4/56451",
 "WDQMS User Guide":"https://confluence.ecmwf.int/display/WIGOSWT/User+Guide",
 "User Manual on the Incident Management System (IMS-RWC)":"https://library.wmo.int/idurl/4/57436",
 "WIGOS Learning Portal":"https://etrp.wmo.int/course/view.php?id=146",
 "RA II WIGOS Newsletters":"https://www.jma.go.jp/jma/jma-eng/satellite/ra2wigosproject/ra2wigosnewsletter.html",
 "Space-based Weather and Climate Extremes Monitoring (SWCEM) Newsletters":"https://public.wmo.int/en/programmes/wmo-space-programme/swcem",
 "DBNet Newsletter":"https://community.wmo.int/node/37663",
 "Guide to the Direct Broadcast Network for Near-real-time Relay of Low Earth Orbit Satellite Data (WMO-No. 1185)":"https://library.wmo.int/idurl/4/55778",
 "Guidelines on Best Practices for User Readiness for New Meteorological Satellite (WMO-No. 1187)":"https://library.wmo.int/idurl/4/55542",
 "Guidelines on Satellite Skills and Knowledge for Operational Meteorologists (SP-12)":"https://library.wmo.int/?lvl=notice_display&id=19843",
 "Guidelines for Satellite-based Nowcasting in Africa (WMO-No. 1309)":"https://library.wmo.int/idurl/4/58348",
 "OSCAR/Space":"https://space.oscar.wmo.int/spacecapabilities",
 "WMO-CGMS VLab":"https://www.wmo-sat.info/vlab/",
 "Coordination Group on Satellite Data Requirements (RA III-IV-SDR Group)":"https://sdr.ucr.ac.cr/",
 "RA I Dissemination Expert Group":"https://community.wmo.int/node/36413",
 "RA II WIGOS Project":"https://www.jma.go.jp/jma/jma-eng/satellite/ra2wigosproject/ra2wigosproject-intro_en_jma.html",
 "SWCEM":"https://public.wmo.int/en/programmes/wmo-space-programme/swcem",
 "Surveys on Satellite Data Use":"https://community.wmo.int/node/36406",
 "Guide to Instruments and Methods of Observation (GIMO) (WMO-No. 8)":"https://library.wmo.int/idurl/4/41650",
 "Manual on the Observation of Clouds and Other Meteors - International Cloud Atlas (WMO-No.407)":"https://cloudatlas.wmo.int/en/home.html",
 "Guide to Operational Weather Radar Best Practices (WMO-No. 1257)":"https://library.wmo.int/idurl/4/68834",
 "IOM Report Series":"https://community.wmo.int/node/37839",
 "Generic AWS Tender Specifications":"https://library.wmo.int/index.php?lvl=notice_display&id=22031",
 "GBON Tender Specifications for AWS":"https://wmoomm.sharepoint.com/:b:/s/wmocpdb/Ecvr45QvsGZAsQA5FZFcnawBboMp_t4o_wAYFbSjsCXFeQ?e=yDMoD1",
 "GBON Tender Specifications for Upper-Air":"https://wmoomm.sharepoint.com/sites/wmocpdb/eve_activityarea/Forms/AllItems.aspx",
 "Knowledge Sharing Portal":"https://community.wmo.int/node/37036",
 "IMOP Learning Portal":"https://etrp.wmo.int/course/index.php?categoryid=38",
 "Guidance on Approaches for Traceability Assurance in GBON stations":"https://wmoomm.sharepoint.com/:b:/s/wmocpdb/Ecv9vy-3ipBCuIOyH7rgIRYBjTfUe-27uYyLMf3VHUpYKg?e=uJ7INs",
 "Training Course on Calibration":"https://etrp.wmo.int/course/view.php?id=375",
 "Community of Practice for RICs":"https://etrp.wmo.int/course/view.php?id=417",
 "WIPPS Newsletter":"https://community.wmo.int/en/news/wipps-newsletter",
 "Manual on WIPPS":"https://library.wmo.int/idurl/4/35703",
 "Guide to WIPPS":"https://library.wmo.int/idurl/4/28978",
 "WIPPS Community Site":"https://community.wmo.int/en/activity-areas/wmo-integrated-processing-and-prediction-system-wipps",
 "ERA Community Site":"https://community.wmo.int/node/33030",
 "WIPPS Web Portal":"https://community.wmo.int/node/37367",
 "WIPPS Webinar":"https://community.wmo.int/node/33832",
 "WIPPS Pilot Projects":"https://community.wmo.int/node/33836",
 "Implementation Status of WIPPS by Members (WMO Data Collection Campaign 2021)":"https://app.powerbi.com/view?r=eyJrIjoiZWI0NDc1M2YtOGY3OS00Y2Y5LTlkM2UtY2ZmNjI3MWNjNjIyIiwidCI6ImVhYTZiZTU0LTQ2ODctNDBjNC05ODI3LWMwNDRiZDhlOGQzYyIsImMiOjl9",
 "WMO Guidelines on Emerging Data Issues (WMO-No. 1239)":"https://library.wmo.int/idurl/4/56904",
 "The WMO Unified Data Policy (Resolution 1 (Cg-Ext.(2021)))":"https://library.wmo.int/doc_num.php?explnum_id=11256",
 "Guide to Participation in Radio-Frequency Coordination (WMO-No. 1159)":"https://library.wmo.int/idurl/4/54853",
 "ITU/WMO Handbook on Use of Radio Spectrum for Meteorology: Weather, Water and Climate Monitoring and Prediction (WMO-No. 1197)":"https://library.wmo.int/idurl/4/55658",
 "National Focal Points for Radio Frequency matters":"https://community.wmo.int/node/33376",
 "Expert Team on Radio Frequency Coordination (ET-RFC)":"https://community.wmo.int/node/33244",
 "Guide to Hydrological Practices, Volume I: From Measurement to Hydrological Information (WMO-No.168)":"https://community.wmo.int/site/knowledge-hub/programmes-and-initiatives/water-resources-assessment/hydrology-publications",
}

def norm(s):
    s = str(s).lower().replace("–", "-").replace("—", "-")
    return re.sub(r"\s+", " ", re.sub(r"[^a-z0-9]+", " ", s)).strip()

def clean(s):
    if pd.isna(s): return ""
    return re.sub(r"\s+", " ", str(s).replace("\t", "").replace("\n", "")).strip()

def main(xlsx_path, out_path="materials.json"):
    df = pd.read_excel(xlsx_path, sheet_name=SHEET, header=0)
    df = df.iloc[:, :5]
    df.columns = ["Title", "Type", "Domain", "Function", "Link"]
    df = df.dropna(subset=["Title"])
    for c in df.columns:
        df[c] = df[c].map(clean)

    link_by_norm = {norm(k): v for k, v in LINK_MAP.items()}
    records, missing = [], []
    for _, r in df.iterrows():
        area, color, code = AREA_MAP.get(r["Function"], ("Other", "#5b6472", "OTH"))
        link = r["Link"] or link_by_norm.get(norm(r["Title"]), "")
        if not link: missing.append(r["Title"])
        records.append({
            "title": r["Title"],
            "type": r["Type"] or "Other types",
            "domain": r["Domain"] or "All domains",
            "area": area, "areaColor": color, "areaCode": code,
            "link": link,
        })

    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=1)

    with_links = sum(1 for x in records if x["link"])
    print(f"Wrote {len(records)} materials to {out_path}")
    print(f"  {with_links} have a link; {len(records) - with_links} without.")
    for m in missing:
        print("  no link:", m)

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__); sys.exit(1)
    main(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else "materials.json")
