import xml.etree.ElementTree as ET
import pandas as pd
import requests

# ── Data fetching (LRS API) ────────────

def fetch_voting(voting_id):
    """Fetch voting results for an ID from the LRS API"""
    url = (f"https://apps.lrs.lt/sip/p2b.ad_sp_balsavimo_rezultatai"
           f"?balsavimo_id=-{voting_id}")
    response = requests.get(url, timeout=15)
    response.raise_for_status()
    root = ET.fromstring(response.content)

    individual_votes = root.findall(".//IndividualusBalsavimoRezultatas")
    if not individual_votes:
        return pd.DataFrame(), ""

    attr_names = {a for v in individual_votes for a in v.attrib.keys()}
    data = {a: [v.attrib.get(a, "") for v in individual_votes] for a in attr_names}
    df = pd.DataFrame(data)
    df['voting'] = voting_id

    general = root.findall(".//BendriBalsavimoRezultatai")
    voting_time = general[0].attrib.get("balsavimo_laikas", "") if general else ""
    df['voting_time'] = pd.to_datetime(voting_time, errors='coerce')
    df['date'] = df['voting_time'].dt.date.astype(str)

    df['member'] = df.apply(
        lambda r: f"{r.get('vardas', '')} {r.get('pavardė', '')}".strip(), axis=1)
    df['fraction'] = df.get('frakcija', "")
    df['result'] = df.get('kaip_balsavo', "")

    auto_name = general[0].attrib.get("balsavimo_pavadinimas", "") if general else ""
    return df[['voting', 'date', 'member', 'fraction', 'result']], auto_name