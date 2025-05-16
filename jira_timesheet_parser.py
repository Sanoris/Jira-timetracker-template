import xml.etree.ElementTree as ET
import pandas as pd
from dateutil import parser as date_parser
import re

def extract_jira_key_and_clean_desc(desc):
    match = re.match(r"\[(.*?)\]\s*(.*)", desc)
    if match:
        return match.group(1), match.group(2)
    return "", desc

def parse_rss_to_timesheet(rss_path, output_excel_path, domain="callahan"):
    tree = ET.parse(rss_path)
    root = tree.getroot()

    entries = []
    for item in root.findall(".//item"):
        title = item.findtext("title")
        description = item.findtext("description")
        pub_date_str = item.findtext("pubDate")
        pub_date = date_parser.parse(pub_date_str)
        date = pub_date.date().isoformat()

        if title and "-" in title:
            parts = title.split(" ", 1)
            ticket = parts[0]
            summary = parts[1] if len(parts) > 1 else ''
        else:
            ticket = title
            summary = ''

        hours = 1.5  # Estimated time per ticket event

        entries.append({
            "DATE": date,
            "HOURS": hours,
            "TICKET": ticket,
            "DESCRIPTION": summary
        })

    df = pd.DataFrame(entries)
    df[["TICKET", "DESCRIPTION"]] = df.apply(
        lambda row: pd.Series(extract_jira_key_and_clean_desc(row["DESCRIPTION"])),
        axis=1
    )
    df = df.groupby(["DATE", "TICKET", "DESCRIPTION"]).agg({"HOURS": "sum"}).reset_index()
    df["LINK"] = df.apply(
        lambda row: f'=HYPERLINK("https://{domain}.atlassian.net/browse/{row["TICKET"]}", "{row["TICKET"]}")', axis=1
    )

    df.to_excel(output_excel_path, index=False)
    print(f"[✓] Timesheet saved to {output_excel_path}")

if __name__ == "__main__":
    parse_rss_to_timesheet("jira_rss_export.xml", "callahan_timesheet.xlsx")
