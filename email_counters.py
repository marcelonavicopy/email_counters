import os
from datetime import datetime
from imap_tools import MailBox, AND
from bs4 import BeautifulSoup
from lxml import etree
from dotenv import load_dotenv
import pandas as pd

load_dotenv()

EMAIL_HOST = os.getenv("EMAIL_HOST")
EMAIL_USER = os.getenv("EMAIL_USER")
EMAIL_PSWD = os.getenv("EMAIL_PSWD")

OPT_TONER = False
OPT_RAW_COUNTER = True

df_counter = pd.DataFrame(
    columns=[
        "Date",
        "MachineModel",
        "SerialNumber",
        "TotalCounter",
        "Mode",
        "Color",
        "Type",
        "Large",
        "Small",
        "Total",
    ]
)

if OPT_TONER:
    df_toner = pd.DataFrame(
        columns=[
            "Date",
            "MachineModel",
            "SerialNumber",
            "TotalCounter",
            "ColorCode",
            "RemainingQuantity",
        ]
    )

mailbox = MailBox(EMAIL_HOST)
mailbox.login(EMAIL_USER, EMAIL_PSWD, "INBOX")
for msg in mailbox.fetch(AND(subject="COUNTER NOTIFICATION")):
    print(msg.uid, msg.date, msg.subject, len(msg.text or msg.html))
    for att in msg.attachments:
        if att.filename == "COUNTER.xml":
            soup = BeautifulSoup(att.payload, "xml")
            root = etree.fromstring(att.payload)
            counters = root.xpath(
                "./ChargeCounter/Counter",
            )
            for c in counters:
                df_counter.loc[len(df_counter)] = [
                    msg.date,
                    soup.MachineModel.string,
                    soup.SerialNumber.string,
                    soup.TotalCounter.string,
                    c.get("Mode"),
                    c.get("Color"),
                    c.get("Type"),
                    c.findtext("Large"),
                    c.findtext("Small"),
                    c.findtext("Total"),
                ]
            if OPT_TONER:
                toner_details = root.xpath("./TonerInformation/Details")
                for d in toner_details:
                    df_toner.loc[len(df_toner)] = [
                        msg.date,
                        soup.MachineModel.string,
                        soup.SerialNumber.string,
                        soup.TotalCounter.string,
                        d.findtext("ColorCode"),
                        d.findtext("RemainingQuantity"),
                    ]


mailbox.logout()

df_counter["Total"] = df_counter["Large"].fillna(0).astype(int) + df_counter[
    "Small"
].fillna(0).astype(int)

scan_total = (
    df_counter[df_counter["Mode"] == "SCAN"].groupby("SerialNumber")["Total"].sum()
)
df_counter["TotalScan"] = df_counter["SerialNumber"].map(scan_total).fillna(0)

black_total = (
    df_counter[(df_counter["Mode"] == "PRINT") & (df_counter["Color"] == "BLACK")]
    .groupby("SerialNumber")["Total"]
    .sum()
)
df_counter["TotalBlack"] = df_counter["SerialNumber"].map(black_total).fillna(0)

color_total = (
    df_counter[(df_counter["Mode"] == "PRINT") & (df_counter["Color"] == "FULL")]
    .groupby("SerialNumber")["Total"]
    .sum()
)
df_counter["TotalColor"] = df_counter["SerialNumber"].map(color_total).fillna(0)

color_low_total = (
    df_counter[
        (df_counter["Mode"] == "PRINT") & (df_counter["Color"].isin(["TWIN", "LOW"]))
    ]
    .groupby("SerialNumber")["Total"]
    .sum()
)
df_counter["TotalColorLow"] = df_counter["SerialNumber"].map(color_low_total).fillna(0)

df_counter_odoo = df_counter[
    ["SerialNumber", "Date", "TotalBlack", "TotalColor", "TotalColorLow", "TotalScan"]
].copy()
df_counter_odoo = df_counter_odoo.drop_duplicates()
df_counter_odoo = df_counter_odoo.rename(
    columns={
        "SerialNumber": "numero_serie",
        "Date": "fecha_lectura",
        "TotalBlack": "contador_bn",
        "TotalColor": "contador_color",
        "TotalColorLow": "contador_color_baja",
        "TotalScan": "contador_escaneo",
    }
)

df_counter_odoo["fecha_lectura"] = pd.to_datetime(
    df_counter_odoo["fecha_lectura"], utc=True
).dt.strftime("%Y-%m-%d")

timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
df_counter_odoo.to_csv(f"counter_{timestamp}.csv", index=False)
if OPT_RAW_COUNTER:
    df_counter.to_csv(f"raw_counter_{timestamp}.csv", index=False)
if OPT_TONER:
    df_toner.to_csv(f"toner_{timestamp}.csv", index=False)
