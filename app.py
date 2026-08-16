# ═══════════════════════════════════════════════════════════════════════════════
#  MENKOR AVIATION — Bombardier Global Express XRS Market Intelligence
#  Version 1.0
# ═══════════════════════════════════════════════════════════════════════════════

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import math
import base64
import os
import shutil

st.set_page_config(
    page_title="Menkor Aviation — Global Express XRS Market",
    page_icon="✈",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ─────────────────────────────────────────────────────────────────────────────
# PHOTO HERO — encodée en base64 depuis le fichier local
# ─────────────────────────────────────────────────────────────────────────────
def get_excel_bytes():
    """Retourne le contenu du fichier Excel pour téléchargement."""
    excel_path = os.path.join(os.path.dirname(__file__), "gxrs_data.xlsx")
    if os.path.exists(excel_path):
        with open(excel_path, "rb") as f:
            return f.read()
    return None

def get_hero_b64():
    img_path = os.path.join(os.path.dirname(__file__), "gxrs_hero.png")
    if os.path.exists(img_path):
        with open(img_path, "rb") as f:
            return base64.b64encode(f.read()).decode()
    return ""

# ─────────────────────────────────────────────────────────────────────────────
# DONNÉES MARCHÉ — 20 APPAREILS
# ─────────────────────────────────────────────────────────────────────────────
AIRCRAFT_DATA = [
    {"SN":"9377","Year":2011,"Reg":"Off-Market","Location":"Paris Le Bourget (LFPB), France","Price_USD":17200000,"Price_Label":"Est. 17,2 M$","TTAF":2735,"Landings":1110,"Engine":"RR BR710A2-20","Airframe":"Smart Parts Preferred","Engine_Prog":"RR CorporateCare","APU":"MSP Gold","Avionics":"Batch 3.4 / DU-875 / HUD / EVS","Connectivity":"Ka-Band Satcom","Notes":"120M c/w ; EU OPS / CAMO EASA ; faible exposition cellule","Score":89.5,"Broker":"Boutsen Aviation","Email":"acquisitions@boutsen.com","Status":"Offre"},
    {"SN":"9347","Year":2010,"Reg":"M-AGMA","Location":"Zurich (LSZH), Suisse","Price_USD":16800000,"Price_Label":"Est. 16,8 M$","TTAF":1738,"Landings":710,"Engine":"RR BR710A2-20","Airframe":"Smart Parts Plus","Engine_Prog":"RR CorporateCare","APU":"MSP Gold","Avionics":"Batch 3.4 / DU-875 / SVS / FANS 1/A+","Connectivity":"Ka-Band Haut Débit","Notes":"TTAF le plus bas du marché (1 738h) ; 120M c/w ; Impeccable","Score":89,"Broker":"Sparfell Aviation","Email":"sales@sparfell.aero","Status":"Offre"},
    {"SN":"9312","Year":2009,"Reg":"N-reg","Location":"Guilford, CT, USA","Price_USD":15500000,"Price_Label":"Est. 15,5 M$","TTAF":3820,"Landings":1490,"Engine":"RR BR710A2-20","Airframe":"Smart Parts Preferred","Engine_Prog":"RRCC Enhanced","APU":"MSP Gold","Avionics":"Batch 3.4 / DU-875AF / SVS / HUD","Connectivity":"Ka-Band Satcom","Notes":"Tri-programme complet ; 120M c/w ; Appareil US clé en main","Score":88.5,"Broker":"Guardian Jet","Email":"sales@guardianjet.com","Status":"Offre"},
    {"SN":"9290","Year":2009,"Reg":"N797KB","Location":"Dallas, TX, USA","Price_USD":15200000,"Price_Label":"Est. 15,2 M$","TTAF":6324,"Landings":2796,"Engine":"RR BR710A2-20","Airframe":"Smart Parts Plus","Engine_Prog":"RR CorporateCare","APU":"MSP Gold","Avionics":"DU-875 / Batch 3.4 / FANS 1/A+ / CPDLC","Connectivity":"Starlink Haut Débit","Notes":"1 propriétaire US depuis neuf ; repeint 2024 ; intérieur refait 2025","Score":87,"Broker":"Guardian Jet","Email":"info@guardianjet.com","Status":"Offre"},
    {"SN":"9336","Year":2010,"Reg":"9H-OKI","Location":"Vienne / Malte","Price_USD":15900000,"Price_Label":"Est. 15,9 M$","TTAF":4850,"Landings":1920,"Engine":"RR BR710A2-20","Airframe":"Smart Parts Plus","Engine_Prog":"RR CorporateCare","APU":"MSP Gold","Avionics":"Batch 3.4 / CPDLC ATN-B1 / FANS 1/A+","Connectivity":"SwiftBroadband","Notes":"15 pax EASA Part-CAT ; prévision 120M/240M propre","Score":86,"Broker":"Avcon Jet AG","Email":"sales@avconjet.at","Status":"Offre"},
    {"SN":"9283","Year":2008,"Reg":"T7-MLR","Location":"Genève (LSGG) / Dubaï","Price_USD":14500000,"Price_Label":"Est. 14,5 M$","TTAF":3651,"Landings":1420,"Engine":"RR BR710A2-20","Airframe":"Smart Parts Plus","Engine_Prog":"RR CorporateCare","APU":"MSP Gold","Avionics":"Primus 2000XP / FANS 1/A / EVS / HUD","Connectivity":"Inmarsat Satcom","Notes":"Faible temps (3 651h) ; 14 pax ; Conformité EASA complète","Score":86.5,"Broker":"Elit'Avia","Email":"aircraftsales@elitavia.com","Status":"Offre"},
    {"SN":"9280","Year":2008,"Reg":"N-reg","Location":"Fort Lauderdale, FL, USA","Price_USD":14400000,"Price_Label":"Est. 14,4 M$","TTAF":4720,"Landings":1860,"Engine":"RR BR710A2-20","Airframe":"Smart Parts Plus","Engine_Prog":"RR CorporateCare","APU":"MSP Gold","Avionics":"Batch 3.4 / DU-875 / FANS 1/A / CPDLC","Connectivity":"Gogo AVANCE L5","Notes":"120M c/w ; vendeur motivé ; galley avant 14 pax","Score":86,"Broker":"Avpro Inc.","Email":"sales@avprojets.com","Status":"Offre"},
    {"SN":"9236","Year":2008,"Reg":"N624BR","Location":"Raleigh, NC, USA","Price_USD":14200000,"Price_Label":"Est. 14,2 M$","TTAF":5210,"Landings":2180,"Engine":"RR BR710A2-20","Airframe":"Smart Parts Plus","Engine_Prog":"RR CorporateCare","APU":"MSP Gold","Avionics":"Batch 3 / Collins CASP / FANS 1/A / CPDLC","Connectivity":"Ku-Band Satcom","Notes":"120M & train effectués ; Collins Venue CMS","Score":85.5,"Broker":"Jetcraft Corporation","Email":"sales@jetcraft.com","Status":"Offre"},
    {"SN":"9259","Year":2008,"Reg":"N-reg","Location":"Chicago, IL, USA","Price_USD":14000000,"Price_Label":"Est. 14,0 M$","TTAF":4920,"Landings":2040,"Engine":"RR BR710A2-20","Airframe":"Smart Parts Plus","Engine_Prog":"RR CorporateCare","APU":"MSP Gold","Avionics":"Batch 3.4 / FANS 1/A+ / CPDLC / Collins Venue","Connectivity":"Gogo 5G Haut Débit","Notes":"14 pax galley avant ; 120M c/w ; sans dommage","Score":85,"Broker":"Central Business Jets","Email":"info@cbjets.com","Status":"Offre"},
    {"SN":"9295","Year":2009,"Reg":"M-reg","Location":"Londres Luton (EGGW), UK","Price_USD":14800000,"Price_Label":"Est. 14,8 M$","TTAF":5430,"Landings":2210,"Engine":"RR BR710A2-20","Airframe":"Smart Parts Plus","Engine_Prog":"RR CorporateCare","APU":"MSP Gold","Avionics":"Batch 3.4 / DU-875 / SVS / LPV / ADS-B Out","Connectivity":"Ka-Band Satcom","Notes":"120M c/w 2019 ; ops privées EASA ; repeint neuf 2023","Score":85,"Broker":"Freestream Aircraft","Email":"sales@freestream.com","Status":"Offre"},
    {"SN":"9248","Year":2007,"Reg":"VH-VSK","Location":"Sydney (YSSY), Australie","Price_USD":13800000,"Price_Label":"Est. 13,8 M$","TTAF":5640,"Landings":2310,"Engine":"RR BR710A2-20","Airframe":"Smart Parts Plus","Engine_Prog":"RR CorporateCare","APU":"MSP Gold","Avionics":"Batch 3.3 / ADS-B Out / TCAS 7.1","Connectivity":"SwiftBroadband","Notes":"Géré ExecuJet ; clé en main Pacifique ; 120M effectuée","Score":83.5,"Broker":"ExecuJet MRO & Sales","Email":"aircraftsales@execujet.com","Status":"Offre"},
    {"SN":"9271","Year":2008,"Reg":"N-reg","Location":"Atlanta, GA, USA","Price_USD":13900000,"Price_Label":"Est. 13,9 M$","TTAF":5810,"Landings":2450,"Engine":"RR BR710A2-20","Airframe":"Smart Parts Plus","Engine_Prog":"RR CorporateCare","APU":"MSP Gold","Avionics":"Batch 3.3 / FANS 1/A / CPDLC / TCAS 7.1","Connectivity":"Gogo AVANCE L5","Notes":"120M c/w ; 13 pax avec stateroom arrière","Score":83,"Broker":"OGara Jets","Email":"sales@ogarajets.com","Status":"Offre"},
    {"SN":"9165","Year":2006,"Reg":"N633BA","Location":"Seattle, WA, USA","Price_USD":13000000,"Price_Label":"Est. 13,0 M$","TTAF":5517,"Landings":2349,"Engine":"RR BR710A2-20","Airframe":"Smart Parts Plus","Engine_Prog":"RR CorporateCare","APU":"MSP Gold","Avionics":"Batch 3.3 / DU-870 / Venue CMS / SVS","Connectivity":"Ka-Band Satcom","Notes":"Mod altitude cabine réduite ; 240M c/w ; cuir neuf","Score":84,"Broker":"Guardian Jet","Email":"sales@guardianjet.com","Status":"Offre"},
    {"SN":"9214","Year":2007,"Reg":"LX-reg","Location":"Luxembourg / Nice","Price_USD":13400000,"Price_Label":"Est. 13,4 M$","TTAF":6450,"Landings":2680,"Engine":"RR BR710A2-20","Airframe":"Smart Parts Plus","Engine_Prog":"RR CorporateCare","APU":"MSP Gold","Avionics":"Batch 3.4 / DU-875 / FANS 1/A+ / CPDLC","Connectivity":"SwiftBroadband","Notes":"EU OPS CAT conforme ; 120M c/w ; train propre","Score":84.5,"Broker":"Global Jet Monaco","Email":"sales@globaljet.mc","Status":"Offre"},
    {"SN":"9202","Year":2006,"Reg":"OE-LCA","Location":"Athènes / Vienne","Price_USD":13200000,"Price_Label":"Est. 13,2 M$","TTAF":7152,"Landings":2890,"Engine":"RR BR710A2-20","Airframe":"Smart Parts Plus","Engine_Prog":"RR CorporateCare","APU":"MSP","Avionics":"Batch 3.3 / HUD & EVS / FANS 1/A / CPDLC","Connectivity":"Inmarsat Satcom","Notes":"Visite lourde 240M (20 ans) c/w ; chambre privée arrière","Score":82,"Broker":"Avcon Jet AG","Email":"sales@avconjet.at","Status":"Offre"},
    {"SN":"9188","Year":2006,"Reg":"N-reg","Location":"Houston, TX, USA","Price_USD":12600000,"Price_Label":"Est. 12,6 M$","TTAF":6980,"Landings":2910,"Engine":"RR BR710A2-20","Airframe":"Smart Parts Plus","Engine_Prog":"RR CorporateCare","APU":"MSP","Avionics":"Batch 3.3 / ADS-B Out / TCAS 7.1 / LPV","Connectivity":"Gogo AVANCE L5","Notes":"240M effectuée ; historique corporate US","Score":81.5,"Broker":"Leading Edge Aviation","Email":"sales@leas.com","Status":"Offre"},
    {"SN":"9222","Year":2007,"Reg":"OE-reg","Location":"Linz / Salzbourg, Autriche","Price_USD":13100000,"Price_Label":"Est. 13,1 M$","TTAF":6780,"Landings":2840,"Engine":"RR BR710A2-20","Airframe":"Smart Parts Plus","Engine_Prog":"RR CorporateCare","APU":"MSP Gold","Avionics":"Batch 3.3 / CPDLC / FANS 1/A / Double HF","Connectivity":"SwiftBroadband","Notes":"EASA Part-CAT ; 120M c/w ; 240M due 2027","Score":81,"Broker":"Manning Aviation","Email":"sales@manningaviation.com","Status":"Offre"},
    {"SN":"9175","Year":2006,"Reg":"N-reg","Location":"Atlanta Rgnl (KFFC), USA","Price_USD":12250000,"Price_Label":"12,25 M$","TTAF":7650,"Landings":3180,"Engine":"RR BR710A2-20","Airframe":"Smart Parts Plus","Engine_Prog":"RR CorporateCare","APU":"MSP Gold","Avionics":"Batch 3.3 / FANS 1/A+ / CPDLC / MTOW élevé","Connectivity":"ATG-5000","Notes":"Révision train c/w ; mod MTOW élevé (SB 700-11)","Score":80.5,"Broker":"Vertical Jet Sales","Email":"sales@verticaljetsales.com","Status":"Prix ferme"},
    {"SN":"9113","Year":2005,"Reg":"N954SP","Location":"Fort Lauderdale, FL, USA","Price_USD":11500000,"Price_Label":"11,5 M$","TTAF":8808,"Landings":4651,"Engine":"RR BR710A2-20","Airframe":"Smart Parts Plus","Engine_Prog":"RR CorporateCare","APU":"Honeywell MSP","Avionics":"Batch 3.3 / FANS 1/A+ / CPDLC / TCAS 7.1","Connectivity":"Gogo ATG-5000","Notes":"Valeur d'entrée temps élevé ; révision train c/w","Score":79,"Broker":"Vertical Jet Sales","Email":"sales@verticaljetsales.com","Status":"Prix ferme"},
    {"SN":"9152","Year":2005,"Reg":"VP-Bxx","Location":"Dubaï World Central (DWC), EAU","Price_USD":11800000,"Price_Label":"Est. 11,8 M$","TTAF":8210,"Landings":3420,"Engine":"RR BR710A2-20","Airframe":"Smart Parts Plus","Engine_Prog":"RR CorporateCare","APU":"MSP","Avionics":"Batch 3.3 / FANS 1/A / Double Collins HF","Connectivity":"Inmarsat Satcom","Notes":"Visite 240M / 20 ans effectuée T1 2026","Score":80,"Broker":"JetHQ Middle East","Email":"sales@jethq.com","Status":"Offre"},
]

df = pd.DataFrame(AIRCRAFT_DATA)

# Liens photos réels extraits du fichier Excel Menkor Aviation
LINKS = {
    "9377": {"photo": "https://www.controller.com/listings/for-sale/bombardier/global-express-xrs/jet-aircraft/3", "email": "acquisitions@boutsen.com"},
    "9347": {"photo": "https://www.avbuyer.com/aircraft/private-jets/bombardier-global/express-xrs", "email": "sales@sparfell.aero"},
    "9312": {"photo": "https://www.guardianjet.com/aircraft-for-sale-listings/aircraft-for-sale.cfm?aid=Bombardier-Global%20Express%20XRS-9312-2619", "email": "sales@guardianjet.com"},
    "9290": {"photo": "https://aircraftexchange.com/aircraft-for-sale/6/bombardier", "email": "info@guardianjet.com"},
    "9336": {"photo": "https://www.avbuyer.com/aircraft/private-jets/bombardier-global/express-xrs", "email": "sales@avconjet.at"},
    "9283": {"photo": "https://aviapages.com/aircraft/t7-mlr/", "email": "aircraftsales@elitavia.com"},
    "9280": {"photo": "https://www.avprojets.com/aircraft-for-sale/", "email": "sales@avprojets.com"},
    "9236": {"photo": "https://www.jetcraft.com/inventory/aircraft/", "email": "sales@jetcraft.com"},
    "9259": {"photo": "https://www.controller.com/listings/for-sale/bombardier/global-express-xrs/jet-aircraft/3", "email": "info@cbjets.com"},
    "9295": {"photo": "https://www.freestream.com/aircraft-for-sale/", "email": "sales@freestream.com"},
    "9248": {"photo": "https://www.avbuyer.com/aircraft/private-jets/bombardier-global/express-xrs", "email": "aircraftsales@execujet.com"},
    "9271": {"photo": "https://ogarajets.com/inventory/", "email": "sales@ogarajets.com"},
    "9202": {"photo": "https://aviapages.com/aircraft/oe-lca/", "email": "sales@avconjet.at"},
    "9165": {"photo": "https://www.guardianjet.com/aircraft-for-sale-listings/", "email": "sales@guardianjet.com"},
    "9214": {"photo": "https://www.globaljet.aero/en/aircraft-sales", "email": "sales@globaljet.mc"},
    "9188": {"photo": "https://leas.com/aircraft-for-sale/", "email": "sales@leas.com"},
    "9222": {"photo": "https://www.avbuyer.com/aircraft/private-jets/bombardier-global/express-xrs", "email": "sales@manningaviation.com"},
    "9175": {"photo": "https://aviapages.com/jet_for_sale/global-express-xrs/n117tp/", "email": "sales@verticaljetsales.com"},
    "9113": {"photo": "https://aviapages.com/jet_for_sale/global-express-xrs/n117tp/", "email": "sales@verticaljetsales.com"},
    "9152": {"photo": "https://jethq.com/aircraft-inventory/", "email": "sales@jethq.com"},
}

# ─────────────────────────────────────────────────────────────────────────────
# PERFORMANCES GLOBAL EXPRESS XRS
# ─────────────────────────────────────────────────────────────────────────────
PERF = {
    "Rayon d'action (plein pax)": "5 904 NM / 10 934 km",
    "Vitesse de croisière": "Mach 0.88 / 488 kts",
    "Altitude maximale": "51 000 ft / FL510",
    "Longueur cabine": "14,66 m / 48,1 ft",
    "Largeur cabine": "2,49 m / 8,2 ft",
    "Hauteur cabine": "1,91 m / 6,3 ft",
    "Passagers typiques": "13 – 17",
    "Passagers max": "19",
    "Motorisation": "2× Rolls-Royce BR710A2-20",
    "MTOW": "43 998 kg",
    "Production": "2005 – 2012",
    "Flotte active": "≈ 160 appareils",
}

# ─────────────────────────────────────────────────────────────────────────────
# DESTINATIONS — rayon depuis Genève
# ─────────────────────────────────────────────────────────────────────────────
GENEVA = (46.2381, 6.1080)
XRS_RANGE_KM = 10934

DESTINATIONS = {
    "Genève":(46.24,6.11),"Londres":(51.48,-0.46),"Paris":(49.01,2.55),
    "Dubaï":(25.25,55.36),"New York":(40.64,-73.78),"Singapour":(1.36,103.99),
    "Tokyo":(35.55,139.78),"Miami":(25.79,-80.29),"Hong Kong":(22.31,113.92),
    "Los Angeles":(33.94,-118.41),"Johannesburg":(-26.13,28.24),"Sydney":(-33.95,151.18),
    "Moscou":(55.97,37.41),"Riyad":(24.96,46.70),"Mumbai":(19.09,72.87),
    "Pékin":(40.08,116.60),"Nairobi":(-1.32,36.93),"São Paulo":(-23.43,-46.47),
    "Chicago":(41.98,-87.91),"Toronto":(43.68,-79.63),"Le Cap":(-33.96,18.60),
    "Doha":(25.27,51.61),"Abu Dhabi":(24.43,54.65),"Istanbul":(40.98,28.82),
    "Le Caire":(30.11,31.40),"Casablanca":(33.37,-7.59),"Bangkok":(13.69,100.75),
    "Kuala Lumpur":(2.75,101.71),"Séoul":(37.46,126.44),"Mexico":(19.44,-99.07),
    "Lagos":(6.58,3.32),"Athènes":(37.94,23.95),"Rome":(41.80,12.25),
    "Madrid":(40.47,-3.57),"Amsterdam":(52.31,4.76),"Zurich":(47.46,8.55),
    "Vienne":(48.11,16.57),"Varsovie":(52.17,20.97),"Stockholm":(59.65,17.92),
    "Oslo":(60.20,11.08),"Lisbonne":(38.77,-9.13),"Barcelone":(41.30,2.08),
    "Bruxelles":(50.90,4.48),"Francfort":(50.03,8.57),"Munich":(48.35,11.79),
    "Nice":(43.66,7.22),"Luxembourg":(49.63,6.21),"Reykjavik":(63.99,-22.63),
    "Marrakech":(31.61,-8.04),"Delhi":(28.56,77.10),"Jakarta":(-6.13,106.66),
    "Dallas":(32.90,-97.04),"Houston":(29.99,-95.34),"Atlanta":(33.64,-84.43),
    "Boston":(42.36,-71.01),"Washington":(38.95,-77.46),"Seattle":(47.45,-122.31),
    "San Francisco":(37.62,-122.38),"Vancouver":(49.19,-123.18),"Montréal":(45.47,-73.74),
    "Buenos Aires":(-34.82,-58.54),"Santiago":(-33.39,-70.79),
}

def haversine(lat1, lon1, lat2, lon2):
    R = 6371
    a = math.sin(math.radians(lat2-lat1)/2)**2 + \
        math.cos(math.radians(lat1))*math.cos(math.radians(lat2)) * \
        math.sin(math.radians(lon2-lon1)/2)**2
    return R * 2 * math.asin(math.sqrt(a))


# ─────────────────────────────────────────────────────────────────────────────
# CSS
# ─────────────────────────────────────────────────────────────────────────────
def inject_css():
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Cormorant+Garamond:wght@300;400;600;700&family=Inter:wght@300;400;500;600&display=swap');

    html, body, [class*="css"] { font-family:'Inter',sans-serif; background-color:#F0F2F5!important; }
    .stApp { background-color:#F0F2F5!important; }

    /* ── HERO PHOTO ── */
    .hero-wrap {
        position:relative; width:100%; overflow:hidden;
        margin:-1rem -1rem 0 -1rem;
        max-height:500px;
    }
    .hero-wrap img {
        width:100%; height:500px;
        object-fit:cover; object-position:center 40%;
        filter:brightness(0.52) saturate(0.9);
        display:block;
    }
    .hero-overlay {
        position:absolute; bottom:0; left:0; right:0;
        padding:2.5rem 3.5rem 2rem;
        background:linear-gradient(to top,rgba(5,15,35,0.95) 0%,rgba(5,15,35,0.5) 55%,transparent 100%);
        border-bottom:3px solid #C9A84C;
    }
    .hero-eyebrow {
        font-size:.68rem; letter-spacing:.22em; text-transform:uppercase;
        color:#C9A84C; font-weight:500; margin-bottom:.5rem;
    }
    .hero-title {
        font-family:'Cormorant Garamond',serif;
        font-size:3rem; font-weight:300; color:#FFFFFF;
        letter-spacing:.04em; line-height:1.1; margin:0;
        text-shadow:0 2px 16px rgba(0,0,0,0.6);
    }
    .hero-title span { color:#C9A84C; font-weight:600; }
    .hero-sub { font-size:.82rem; color:#B0C4DE; margin-top:.7rem; letter-spacing:.06em; }

    /* ── STAT PILLS ── */
    .stat-pill {
        display:inline-block; background:#FFFFFF; border:1px solid #D8DFE8;
        border-radius:40px; padding:.45rem 1.1rem; font-size:.78rem;
        color:#1A2A4A; font-weight:500; margin:.2rem;
        box-shadow:0 1px 4px rgba(11,22,41,0.08);
    }
    .stat-pill b { color:#C9A84C; }

    /* ── SECTION LABEL ── */
    .section-eyebrow {
        font-size:.65rem; letter-spacing:.22em; text-transform:uppercase;
        color:#C9A84C; font-weight:600; margin-bottom:1rem;
        padding-bottom:.5rem; border-bottom:1px solid #D8DFE8;
    }

    /* ── PERF GRID ── */
    .perf-grid { display:grid; grid-template-columns:repeat(3,1fr); gap:.7rem; margin-bottom:1.2rem; }
    .perf-item {
        background:#FFFFFF; border-radius:8px; padding:.85rem 1rem;
        border-left:3px solid #C9A84C; box-shadow:0 1px 4px rgba(11,22,41,0.07);
    }
    .perf-label { font-size:.63rem; letter-spacing:.1em; text-transform:uppercase; color:#8496B0; margin-bottom:.25rem; }
    .perf-value { font-size:.86rem; font-weight:600; color:#1A2A4A; }

    /* ── SCORE BADGE ── */
    .score-high { color:#1a7a3c; font-weight:700; }
    .score-mid  { color:#b07a00; font-weight:700; }
    .score-low  { color:#a03020; font-weight:700; }

    /* ── FOOTER ── */
    .footer-band {
        background:#050F23; border-top:2px solid #C9A84C;
        padding:1.2rem 3rem; margin:2rem -1rem -1rem;
        text-align:center; font-size:.72rem; color:#8496B0; letter-spacing:.06em;
    }
    .footer-band b { color:#C9A84C; }

    #MainMenu, footer, header { visibility:hidden; }
    .block-container { padding-top:0!important; padding-bottom:2rem; }
    </style>
    """, unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# GRAPHIQUES
# ─────────────────────────────────────────────────────────────────────────────
def chart_prix_vs_annee(df_f):
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=df_f["Year"], y=df_f["Price_USD"]/1_000_000,
        mode="markers+text",
        marker=dict(
            size=df_f["TTAF"].apply(lambda x: max(8, min(30, x/320))),
            color=df_f["Score"],
            colorscale=[[0,"#1A3A6E"],[0.5,"#C9A84C"],[1,"#4CAF50"]],
            showscale=True,
            colorbar=dict(title="Score", thickness=10, len=0.6,
                          tickfont=dict(size=9, family="Inter")),
            line=dict(color="#FFFFFF", width=1.5), opacity=0.9,
        ),
        text=df_f["SN"].apply(lambda x: f"SN {x}"),
        textposition="top center",
        textfont=dict(size=8, color="#1A2A4A", family="Inter"),
        hovertemplate=(
            "<b>SN %{customdata[0]}</b> (%{x})<br>"
            "Valeur : %{y:.2f} M$<br>"
            "TTAF : %{customdata[1]:,}h<br>"
            "Score : %{customdata[2]}/100<br>"
            "%{customdata[3]}<extra></extra>"
        ),
        customdata=df_f[["SN","TTAF","Score","Location"]].values,
    ))
    fig.update_layout(
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        margin=dict(t=10,b=50,l=10,r=60), height=330,
        xaxis=dict(title="Année", showgrid=True, gridcolor="#E8EDF3",
                   tickfont=dict(family="Inter",size=10,color="#8496B0")),
        yaxis=dict(title="Valeur (M$)", showgrid=True, gridcolor="#E8EDF3",
                   tickprefix="$", ticksuffix="M",
                   tickfont=dict(family="Inter",size=10,color="#8496B0")),
        showlegend=False,
    )
    return fig


def chart_range_from_geneva():
    lat_gva, lon_gva = GENEVA
    reachable, unreachable = [], []
    for city,(lat,lon) in DESTINATIONS.items():
        dist = haversine(lat_gva,lon_gva,lat,lon)
        (reachable if dist<=XRS_RANGE_KM else unreachable).append((city,lat,lon,dist))

    fig = go.Figure()

    # Cercle de portée
    R_earth = 6371
    clats, clons = [], []
    for angle in [i*(360/72) for i in range(73)]:
        ang_rad = math.radians(angle)
        d = XRS_RANGE_KM/R_earth
        lat2 = math.asin(math.sin(math.radians(lat_gva))*math.cos(d) +
                         math.cos(math.radians(lat_gva))*math.sin(d)*math.cos(ang_rad))
        lon2 = math.radians(lon_gva) + math.atan2(
            math.sin(ang_rad)*math.sin(d)*math.cos(math.radians(lat_gva)),
            math.cos(d)-math.sin(math.radians(lat_gva))*math.sin(lat2))
        clats.append(math.degrees(lat2)); clons.append(math.degrees(lon2))

    fig.add_trace(go.Scattergeo(
        lat=clats, lon=clons, mode="lines",
        line=dict(color="#C9A84C",width=1.5,dash="dot"),
        fill="toself", fillcolor="rgba(201,168,76,0.09)",
        name=f"Rayon : {XRS_RANGE_KM:,} km", hoverinfo="skip"))

    if unreachable:
        fig.add_trace(go.Scattergeo(
            lat=[c[1] for c in unreachable], lon=[c[2] for c in unreachable],
            mode="markers", marker=dict(size=5,color="#BBBBBB",opacity=0.4),
            text=[f"{c[0]}<br>{c[3]:,.0f} km — Hors portée" for c in unreachable],
            hovertemplate="%{text}<extra></extra>", showlegend=False))

    if reachable:
        fig.add_trace(go.Scattergeo(
            lat=[c[1] for c in reachable], lon=[c[2] for c in reachable],
            mode="markers+text",
            marker=dict(size=8,color="#C9A84C",line=dict(color="#FFFFFF",width=1.5),opacity=0.9),
            text=[c[0] for c in reachable], textposition="top center",
            textfont=dict(size=7.5,color="#1A2A4A",family="Inter"),
            hovertemplate="<b>%{text}</b><br>%{customdata:,.0f} km de Genève<extra></extra>",
            customdata=[c[3] for c in reachable], name="À portée"))

    for city,lat,lon,dist in reachable[:12]:
        fig.add_trace(go.Scattergeo(
            lat=[lat_gva,lat], lon=[lon_gva,lon], mode="lines",
            line=dict(color="#C9A84C",width=0.5), opacity=0.18,
            showlegend=False, hoverinfo="skip"))

    fig.add_trace(go.Scattergeo(
        lat=[lat_gva], lon=[lon_gva], mode="markers+text",
        marker=dict(size=14,color="#050F23",line=dict(color="#C9A84C",width=3),symbol="star"),
        text=["Genève"], textposition="bottom center",
        textfont=dict(size=11,color="#1A2A4A",family="Inter"),
        name="Genève (LSGG)",
        hovertemplate="<b>Genève — Origine</b><br>Rayon Global XRS : 10 934 km<extra></extra>"))

    fig.update_layout(
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        margin=dict(t=0,b=0,l=0,r=0), height=440, showlegend=True,
        legend=dict(bgcolor="rgba(255,255,255,0.92)",bordercolor="#D8DFE8",borderwidth=1,
                    font=dict(family="Inter",size=10,color="#1A2A4A"),x=0.01,y=0.99),
        geo=dict(
            projection_type="natural earth",
            showland=True, landcolor="#E8EDF3",
            showocean=True, oceancolor="#C8D8E8",
            showcoastlines=True, coastlinecolor="#B8C8D8",
            showcountries=True, countrycolor="#D0D8E0",
            showframe=False, bgcolor="rgba(0,0,0,0)",
            center=dict(lat=30,lon=15),
            lataxis=dict(range=[-55,85]), lonaxis=dict(range=[-130,170])),
    )
    return fig


def chart_ttaf(df_f):
    fig = go.Figure()
    fig.add_trace(go.Histogram(
        x=df_f["TTAF"], nbinsx=8,
        marker_color="#C9A84C",
        marker_line=dict(color="#FFFFFF",width=1.5),
        opacity=0.9,
        hovertemplate="TTAF : %{x}h<br>Nb : %{y}<extra></extra>"))
    fig.update_layout(
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        margin=dict(t=10,b=30,l=10,r=10), height=185, bargap=0.12,
        xaxis=dict(showgrid=False,tickfont=dict(family="Inter",size=9,color="#8496B0")),
        yaxis=dict(showgrid=True,gridcolor="#E8EDF3",tickfont=dict(family="Inter",size=9,color="#8496B0")),
        showlegend=False)
    return fig


# ─────────────────────────────────────────────────────────────────────────────
# APPLICATION PRINCIPALE
# ─────────────────────────────────────────────────────────────────────────────
def main():
    inject_css()

    # ── HERO PHOTO ───────────────────────────────────────────────────────────
    hero_b64 = get_hero_b64()
    if hero_b64:
        st.markdown(f"""
        <div class="hero-wrap">
            <img src="data:image/png;base64,{hero_b64}" alt="Bombardier Global Express XRS" />
            <div class="hero-overlay">
                <div class="hero-eyebrow">✦ Menkor Aviation GBL — Veille Marché</div>
                <div class="hero-title">Bombardier<br><span>Global Express XRS</span><br>Marché de l'Occasion</div>
                <div class="hero-sub">
                    20 appareils recensés dans le monde &nbsp;·&nbsp; Mis à jour août 2026
                    &nbsp;·&nbsp; Valeurs en USD &nbsp;·&nbsp; Source : Menkor Aviation Research
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown("""
        <div style="background:linear-gradient(135deg,#050F23 0%,#1A3A6E 55%,#050F23 100%);
                    border-bottom:3px solid #C9A84C; padding:3rem 3.5rem 2.5rem;
                    margin:-1rem -1rem 0 -1rem;">
            <div style="font-size:.68rem;letter-spacing:.22em;text-transform:uppercase;color:#C9A84C;font-weight:500;margin-bottom:.5rem;">
                ✦ Menkor Aviation GBL — Veille Marché</div>
            <div style="font-family:'Cormorant Garamond',serif;font-size:3rem;font-weight:300;
                        color:#FFFFFF;letter-spacing:.04em;line-height:1.1;">
                Bombardier <span style="color:#C9A84C;font-weight:600;">Global Express XRS</span><br>Marché de l'Occasion</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("<div style='margin-top:1.5rem'></div>", unsafe_allow_html=True)

    # ── INDICATEURS CLÉS ─────────────────────────────────────────────────────
    total = len(df)
    prix_min = df["Price_USD"].min()/1e6
    prix_max = df["Price_USD"].max()/1e6
    prix_moy = df["Price_USD"].mean()/1e6
    ttaf_moy = df["TTAF"].mean()
    score_moy = df["Score"].mean()
    prix_ferme = df[df["Status"]=="Prix ferme"].shape[0]

    st.markdown(f"""
    <div style="margin-bottom:1.2rem;">
        <span class="stat-pill">✈ <b>{total}</b> appareils recensés</span>
        <span class="stat-pill">💰 Fourchette <b>{prix_min:.1f} M$</b> – <b>{prix_max:.1f} M$</b></span>
        <span class="stat-pill">📊 Valeur moy. <b>{prix_moy:.1f} M$</b></span>
        <span class="stat-pill">⏱ TTAF moy. <b>{ttaf_moy:,.0f}h</b></span>
        <span class="stat-pill">⭐ Score moy. <b>{score_moy:.1f}/100</b></span>
        <span class="stat-pill">📢 <b>{prix_ferme}</b> avec prix ferme affiché</span>
        <span class="stat-pill">📅 Millésimes <b>2005–2011</b></span>
    </div>
    """, unsafe_allow_html=True)

    # ── BOUTON TÉLÉCHARGEMENT EXCEL ──────────────────────────────────────────
    excel_bytes = get_excel_bytes()
    if excel_bytes:
        col_dl, col_spacer = st.columns([1, 4])
        with col_dl:
            st.download_button(
                label="⬇️ Télécharger l'étude complète (Excel)",
                data=excel_bytes,
                file_name="Menkor_GlobalXRS_Market_Intelligence_2026.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="Fichier Excel complet avec données, programmes et liens"
            )

    # ── FILTRES ───────────────────────────────────────────────────────────────
    col_f1, col_f2, col_f3, col_f4 = st.columns([1,1,1,1])
    with col_f1:
        year_range = st.slider("Année de fabrication", 2005, 2011, (2005, 2011), key="yr")
    with col_f2:
        price_range = st.slider("Valeur estimée (M$)", 11, 18, (11, 18), key="pr")
    with col_f3:
        ttaf_max = st.slider("TTAF maximum (heures)", 1000, 9000, 9000, step=500, key="tt")
    with col_f4:
        score_min = st.slider("Score minimum (/100)", 75, 92, 75, step=1, key="sc")

    df_f = df[
        (df["Year"]>=year_range[0]) & (df["Year"]<=year_range[1]) &
        (df["Price_USD"]>=price_range[0]*1e6) & (df["Price_USD"]<=price_range[1]*1e6) &
        (df["TTAF"]<=ttaf_max) & (df["Score"]>=score_min)
    ].copy()

    st.markdown(
        f"<div style='font-size:.78rem;color:#8496B0;margin-bottom:.8rem;'>"
        f"Affichage de <b style='color:#1A2A4A'>{len(df_f)}</b> appareil(s) sur {total}</div>",
        unsafe_allow_html=True)

    # ── LAYOUT ───────────────────────────────────────────────────────────────
    col_left, col_right = st.columns([1.65, 1])

    with col_left:
        st.markdown('<div class="section-eyebrow">📋 Inventaire du Marché</div>', unsafe_allow_html=True)

        df_display = df_f[[
            "SN","Year","Reg","Location","Price_Label","TTAF","Landings",
            "Airframe","Engine_Prog","APU","Avionics","Connectivity","Notes","Score","Broker","Status"
        ]].copy()

        df_display["📷 Photos"] = df_display["SN"].apply(
            lambda sn: LINKS.get(sn, {}).get("photo", ""))
        df_display["✉️ Contact"] = df_display["SN"].apply(
            lambda sn: f"mailto:{LINKS.get(sn, {}).get('email', '')}")

        df_display.columns = [
            "SN","Année","Immat.","Localisation","Valeur Est.","TTAF","Atterr.",
            "Pgm Cellule","Pgm Moteur","APU","Avionique","Connectivité","Notes","Score","Courtier","Statut",
            "📷 Photos","✉️ Contact"
        ]

        st.dataframe(
            df_display, hide_index=True, use_container_width=True, height=460,
            column_config={
                "SN": st.column_config.TextColumn("SN", width=60),
                "Année": st.column_config.NumberColumn("Année", width=58, format="%d"),
                "Immat.": st.column_config.TextColumn("Immat.", width=80),
                "Localisation": st.column_config.TextColumn("Localisation", width=170),
                "Valeur Est.": st.column_config.TextColumn("Valeur Est.", width=95),
                "TTAF": st.column_config.NumberColumn("TTAF", width=65, format="%d h"),
                "Atterr.": st.column_config.NumberColumn("Atterr.", width=58, format="%d"),
                "Pgm Cellule": st.column_config.TextColumn("Pgm Cellule", width=140),
                "Pgm Moteur": st.column_config.TextColumn("Pgm Moteur", width=155),
                "APU": st.column_config.TextColumn("APU", width=90),
                "Avionique": st.column_config.TextColumn("Avionique", width=205),
                "Connectivité": st.column_config.TextColumn("Connectivité", width=130),
                "Notes": st.column_config.TextColumn("Notes", width=255),
                "Score": st.column_config.NumberColumn("Score", width=58, format="%.1f"),
                "Courtier": st.column_config.TextColumn("Courtier", width=155),
                "Statut": st.column_config.TextColumn("Statut", width=82),
                "📷 Photos": st.column_config.LinkColumn("📷 Photos", width=85, display_text="Voir →"),
                "✉️ Contact": st.column_config.LinkColumn("✉️ Contact", width=85, display_text="Email →"),
            })

        st.markdown('<div class="section-eyebrow" style="margin-top:1.4rem">📈 Valeur vs Année · Couleur = Score · Taille = TTAF</div>', unsafe_allow_html=True)
        st.plotly_chart(chart_prix_vs_annee(df_f), use_container_width=True,
                        config={"displayModeBar": False}, key="scatter")

    with col_right:
        st.markdown('<div class="section-eyebrow">⚡ Global Express XRS — Caractéristiques</div>', unsafe_allow_html=True)
        st.markdown('<div class="perf-grid">' + "".join([
            f'<div class="perf-item"><div class="perf-label">{k}</div><div class="perf-value">{v}</div></div>'
            for k,v in PERF.items()
        ]) + '</div>', unsafe_allow_html=True)

        st.markdown('<div class="section-eyebrow" style="margin-top:.8rem">⏱ Répartition TTAF (filtré)</div>', unsafe_allow_html=True)
        st.plotly_chart(chart_ttaf(df_f), use_container_width=True,
                        config={"displayModeBar": False}, key="ttaf_hist")

        st.markdown('<div class="section-eyebrow" style="margin-top:.8rem">🌍 Rayon d\'action depuis Genève (LSGG)</div>', unsafe_allow_html=True)
        st.plotly_chart(chart_range_from_geneva(), use_container_width=True,
                        config={"displayModeBar": False}, key="range_map")

        reachable_count = sum(1 for c,(lat,lon) in DESTINATIONS.items()
                              if haversine(GENEVA[0],GENEVA[1],lat,lon)<=XRS_RANGE_KM)
        st.markdown(
            f"<div style='font-size:.78rem;color:#8496B0;text-align:center;margin-top:-.5rem;'>"
            f"<b style='color:#C9A84C'>{reachable_count}</b> grandes villes accessibles sans escale depuis Genève"
            f"</div>", unsafe_allow_html=True)

    # ── ANALYSES MARCHÉ ──────────────────────────────────────────────────────
    st.markdown('<div class="section-eyebrow" style="margin-top:1.5rem">💡 Analyses Marché — Août 2026</div>', unsafe_allow_html=True)
    col_i1, col_i2, col_i3 = st.columns(3)
    for col, (title, text) in zip([col_i1,col_i2,col_i3],[
        ("Positionnement & Valeur",
         "Le Global Express XRS (2005–2011) se positionne entre 11,5 M$ et 17,2 M$ selon le millésime, le TTAF et l'état des programmes. Les unités Batch 3.4 / DU-875 avec CorporateCare complète occupent le haut du spectre et se négocient rapidement."),
        ("Avionique & Liquidité",
         "La suite Batch 3.4 avec écrans DU-875 LCD, FANS 1/A+, CPDLC ATN-B1 et SVS est devenue le standard de liquidité. Les appareils encore équipés de CRT (Batch 3.0/3.1) subissent une décote significative et sont à éviter pour une revente à court terme."),
        ("Points de Vigilance",
         "La visite lourde 120M (10 ans) représente 750 K$ à 1,1 M$. La 240M (20 ans) avec révision des trains atteint 2,4 M$. Tout appareil avec échéance dans les 18 mois doit faire l'objet d'une provision dans la négociation du prix."),
    ]):
        with col:
            st.markdown(f"""
            <div style="background:#FFFFFF;border-radius:8px;padding:1.2rem 1.3rem;
                        border-top:3px solid #C9A84C;box-shadow:0 2px 8px rgba(11,22,41,0.07);">
                <div style="font-size:.72rem;letter-spacing:.1em;text-transform:uppercase;
                            color:#C9A84C;font-weight:700;margin-bottom:.6rem;">{title}</div>
                <div style="font-size:.83rem;color:#3A4A5E;line-height:1.55;">{text}</div>
            </div>""", unsafe_allow_html=True)

    # ── PROGRAMMES DE MAINTENANCE ─────────────────────────────────────────────
    st.markdown("---")
    st.markdown(
        "<div style='font-size:.65rem;letter-spacing:.22em;text-transform:uppercase;"
        "color:#C9A84C;font-weight:600;margin-bottom:.3rem;'>Programmes de Maintenance</div>",
        unsafe_allow_html=True)
    st.markdown("## Programmes de Maintenance & Protection de la Valeur Résiduelle")
    st.markdown(
        "Le Global Express XRS repose sur un **triptyque de programmes d'entretien à l'heure de vol** "
        "qui sécurisent les coûts d'exploitation et protègent la valeur résiduelle : "
        "**CorporateCare Rolls-Royce** pour les deux moteurs BR710, "
        "**Smart Parts Bombardier** pour la cellule et les systèmes, "
        "et **MSP Gold Honeywell** pour l'APU. "
        "La présence simultanée de ces trois contrats est aujourd'hui le standard exigé par le marché secondaire."
    )

    st.markdown("<div style='margin-top:1.2rem'></div>", unsafe_allow_html=True)
    col_m1, col_m2, col_m3 = st.columns(3)

    with col_m1:
        st.markdown(
            "<div style='font-size:.6rem;letter-spacing:.18em;text-transform:uppercase;"
            "color:#C9A84C;font-weight:700;'>Programme 01 · Moteurs</div>",
            unsafe_allow_html=True)
        st.markdown("### Rolls-Royce CorporateCare")
        st.caption("Programme OEM — 2× BR710A2-20")
        st.markdown(
            "Programme moteurs de référence géré directement par Rolls-Royce, "
            "couvrant les deux réacteurs BR710A2-20 avec accès prioritaire "
            "au réseau MRO mondial de l'OEM."
        )
        st.markdown(
            "- **CorporateCare (RRCC)** — Révisions programmées et non programmées, "
            "main-d'œuvre 100%, pièces de maintenance de ligne\n"
            "- **CorporateCare Enhanced (RRCCE)** — Tout le RRCC, plus couverture des nacelles, "
            "inverseurs de poussée, tuyères et corrosion non programmée\n"
            "- Le **RRCCE** est le niveau premium attendu sur les derniers millésimes "
            "et constitue un atout majeur lors de la revente\n"
            "- Un appareil avec moteurs hors programme représente un risque "
            "significatif que le marché valorise immédiatement à la baisse"
        )

    with col_m2:
        st.markdown(
            "<div style='font-size:.6rem;letter-spacing:.18em;text-transform:uppercase;"
            "color:#C9A84C;font-weight:700;'>Programme 02 · Cellule & Systèmes</div>",
            unsafe_allow_html=True)
        st.markdown("### Bombardier Smart Parts")
        st.caption("Programme OEM — Cellule, systèmes & avionique")
        st.markdown(
            "Programme cellule de Bombardier couvrant l'ensemble des systèmes embarqués "
            "et permettant de prévisibiliser les dépenses de maintenance hors moteurs."
        )
        st.markdown(
            "- **Smart Parts Plus** — Systèmes de vol, hydraulique, commandes de vol, "
            "calculateurs électriques, trains d'atterrissage et boîtiers avioniques\n"
            "- **Smart Parts Preferred** — Niveau étendu avec couverture renforcée "
            "sur les composants structuraux et systèmes critiques\n"
            "- Élimine l'exposition aux défaillances coûteuses sur les LRU Bombardier\n"
            "- **Visites 120M / 240M** — Les grandes visites décennales et vingtennales "
            "constituent les postes de coût les plus significatifs de la vie de l'appareil "
            "et doivent être vérifiées lors de toute acquisition"
        )

    with col_m3:
        st.markdown(
            "<div style='font-size:.6rem;letter-spacing:.18em;text-transform:uppercase;"
            "color:#C9A84C;font-weight:700;'>Programme 03 · APU & Avionique</div>",
            unsafe_allow_html=True)
        st.markdown("### Honeywell MSP Gold")
        st.caption("APU RE220(GX) + Suite Primus 2000XP")
        st.markdown(
            "Programme Honeywell couvrant le groupe auxiliaire de puissance RE220(GX) "
            "et, en option, la suite avionique Primus 2000XP."
        )
        st.markdown(
            "- **MSP Gold APU** — Révisions majeures, inspections section chaude, "
            "remplacement des composants LRU et dépannages non programmés\n"
            "- **Batch 3.4 / DU-875 LCD** — La mise à niveau avionique vers les écrans "
            "LCD DU-875, FANS 1/A+, CPDLC ATN-B1 et SVS est un facteur de liquidité "
            "déterminant — les appareils CRT (Batch 3.0) sont pénalisés sur le marché\n"
            "- Un APU sous MSP Gold est un **critère de sélection** pour les acheteurs "
            "institutionnels et les gestionnaires de flotte"
        )

    # ── FOOTER ────────────────────────────────────────────────────────────────
    st.markdown("""
    <div class="footer-band">
        <b>MENKOR AVIATION GBL</b> &nbsp;|&nbsp;
        Ce récapitulatif marché est fourni à titre informatif uniquement. Valeurs estimatives basées sur
        les recherches de Menkor Aviation — peuvent différer des prix de transaction réels.
        &nbsp;|&nbsp; © 2026 Menkor Aviation GBL
    </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
