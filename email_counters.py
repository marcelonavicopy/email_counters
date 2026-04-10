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

OPT_TONER = True

df1 = pd.DataFrame(
    columns=[
        "Date",
        "MachineModel",
        "SerialNumber",
        "TotalCounter",
        "Model",
        "Color",
        "Type",
        "Large",
        "Small",
        "Total",
    ]
)

if OPT_TONER:
    df2 = pd.DataFrame(
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
                df1.loc[len(df1)] = [
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
                    df2.loc[len(df2)] = [
                        msg.date,
                        soup.MachineModel.string,
                        soup.SerialNumber.string,
                        soup.TotalCounter.string,
                        d.findtext("ColorCode"),
                        d.findtext("RemainingQuantity"),
                    ]


mailbox.logout()

timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
df1.to_csv(f"counter_{timestamp}.csv", index=False)
if OPT_TONER:
    df2.to_csv(f"toner_{timestamp}.csv", index=False)
