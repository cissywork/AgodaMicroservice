import imaplib
import email
from email.header import decode_header
from bs4 import BeautifulSoup
import re
from datetime import datetime, timedelta
import json
import os
from flask import Flask, request, jsonify,Response

'''
capture all status - amended, cancelled and confirmed 
"ehxe srtr fkuc iojk") #find in "app password" Security from Gmail Account
Check only upcoming reservations (checkin date > = today)
'''

''' FOR Future Enhancement ONLY
# Project: Agoda Email Reservation Extractor (Simplified)

This is a simplified Flask-based Web API that:
- Connects to Gmail
- Parses Agoda reservation confirmation emails
- Extracts structured reservation data
- Saves the result as JSON
- Tracks last scan time to avoid duplication

## Folder Structure (Simplified)

```
email_scraper/
├── main.py               # Entry point, contains Flask app & logic
├── data/
│   ├── last_scan_time.txt
│   └── reservations.json
├── credentials.json      # Gmail API credentials
├── token.json            # Gmail API token (generated after first login)
├── templates/            # (Optional) For future front-end
└── requirements.txt
```

## How to Use

1. Install requirements:
```bash
pip install -r requirements.txt
```

2. Run the app:
```bash
python main.py
```

3. Access API endpoints:
- `GET /reservations` → Incremental scan
- `GET /reservations?mode=reset` → Rescan 365 days
- `GET /reservations?mode=range&start=YYYY-MM-DD&end=YYYY-MM-DD` → Custom range

4. test
http://localhost:5000/scan-emails for incremental scan
http://localhost:5000/scan-emails?mode=reset to reset and scan all from 365 days
http://localhost:5000/scan-emails?mode=range&start=2024-12-01&end=2025-01-01 to define a custom scan window
'''

app = Flask(__name__)

@app.route("/scan-emails", methods=["POST","GET"])

def scan_emails():
    output_dir = "data"
    timestamp_file = os.path.join(output_dir, "last_scan_time.txt")
    now = datetime.now()
    previous_scan_time = now - timedelta(days=365)

    # Get platform, mode and optional parameters
    platform = request.args.get("platform").capitalize()
    mode = request.args.get("mode")
    start_date = request.args.get("start")
    end_date = request.args.get("end")

    try:
        if mode == "reset":
            previous_scan_time = now - timedelta(days=365)

        elif mode == "range" and start_date and end_date:
            previous_scan_time = datetime.strptime(start_date, "%Y-%m-%d")
            end_time = datetime.strptime(end_date, "%Y-%m-%d")
            if end_time < now:
                now = end_time  # override scan-to time

        elif os.path.exists(timestamp_file):
            with open(timestamp_file, "r") as f:
                last_scan_str = f.read().strip()
            previous_scan_time = datetime.strptime(last_scan_str, "%Y-%m-%d %H:%M:%S")

    except Exception as e:
        print("Error parsing dates:", e)
        previous_scan_time = now - timedelta(days=365)

    print(f"⏱️ Scanning emails from {previous_scan_time} to {now}")

    # Set up IMAP
    imap = imaplib.IMAP4_SSL("imap.gmail.com")
    imap.login("comfortopulence@gmail.com",  "ehxe srtr fkuc iojk") #find in "app password" Security from Gmail Account
    imap.select("INBOX")

    search_since = previous_scan_time.strftime("%d-%b-%Y")
    status, messages = imap.search(None, f'(SINCE {search_since} FROM "no-reply@agoda.com" SUBJECT "Agoda Booking ID")')
    email_ids = messages[0].split()

    translation_map = {
        '3 Bedroom Lakeview Panoramic Suite (665874002)': '4803',
        'Panoramic View Presidential Suite (667404026)': '4904',
        'Suite with Lake View (667303212)': '4201',
        'Suite Lake View (665872674)': '2004',
        'Suite with Kitchenette and Balcony (667292867)': '2611',
        'Suite with Balcony (667302451)': '1601'
    }

    reservation_dict = {}
    fetch_data = imap.fetch

    for num in reversed(email_ids):
        res, msg = fetch_data(num, "(RFC822)")
        for response in msg:
            if not isinstance(response, tuple):
                continue

            msg = email.message_from_bytes(response[1])
            subject, encoding = decode_header(msg["Subject"])[0]
            if isinstance(subject, bytes):
                subject = subject.decode(encoding or "utf-8")

            status = ""
            if "CONFIRMED" in subject.upper():
                status = "Confirmed"
            elif "CANCELLED" in subject.upper():
                status = "Cancelled"
            elif "AMENDED" in subject.upper():
                status = "Amended"

            email_date = msg["Date"]
            try:
                last_updated_dt = datetime.strptime(email_date[:31], "%a, %d %b %Y %H:%M:%S %z").astimezone().replace(tzinfo=None)
                last_updated = last_updated_dt.strftime("%Y-%m-%d %H:%M:%S")
            except Exception:
                last_updated_dt = now
                last_updated = now.strftime("%Y-%m-%d %H:%M:%S")

            html_content = msg.get_payload(decode=True).decode(errors="ignore")
            if not html_content:
                continue

            soup = BeautifulSoup(html_content, 'html.parser')

            booking_id = guest_name = total_price = rate_plan = room_type = rooms = occupancy = None
            adults = kids = 0
            checkin = checkout = nights = None
            unit_number = "N/A"
            country_of_residence = "N/A"

            for td in soup.find_all('td'):
                if td.get_text(strip=True) == 'Booking ID':
                    next_td = td.find_next('td')
                    if next_td:
                        booking_id = next_td.get_text(strip=True)
                        break

            for span in soup.find_all('span'):
                if 'Customer First Name' in span.get_text():
                    td = span.find_parent('td')
                    if td:
                        next_td = td.find_next_sibling('td')
                        if next_td:
                            first_name = next_td.get_text(strip=True)
                            last_name_td = td.find_next('tr').find_all('td')[1]
                            last_name = last_name_td.get_text(strip=True)
                            guest_name = f"{first_name} {last_name}"
                            break

            for span in soup.find_all("span"):
                if "Net rate (incl. taxes & fees)" in span.get_text():
                    parent_div = span.find_parent("div")
                    if parent_div:
                        next_div = parent_div.find_next_sibling("div")
                        if next_div:
                            total_price = next_div.get_text(strip=True)
                            break

            for span in soup.find_all("span"):
                if "Rate Plan name" in span.get_text():
                    rate_plan = span.get_text(strip=True).replace("Rate Plan name:", "").strip()
                    break

            all_trs = soup.find_all("tr")
            for tr in all_trs:
                tds = tr.find_all("td")
                if len(tds) == 4 and "Room Type" in tds[0].get_text():
                    try:
                        next_tds = all_trs[all_trs.index(tr)+1].find_all("td")
                        room_type = next_tds[0].get_text(strip=True)
                        rooms = next_tds[1].get_text(strip=True)
                        occupancy = next_tds[2].get_text(strip=True)
                    except Exception:
                        pass
                    break

            for full_name, unit in translation_map.items():
                base_name = full_name.split(' (')[0].strip()
                if room_type and base_name in room_type:
                    unit_number = unit
                    break

            if occupancy:
                occupancy = occupancy.strip()
                adult_match = re.search(r'(\d+)\s*Adult', occupancy, re.IGNORECASE)
                kid_match = re.search(r'(\d+)\s*Child', occupancy, re.IGNORECASE)
                if adult_match:
                    adults = int(adult_match.group(1))
                if kid_match:
                    kids = int(kid_match.group(1))

            for tr in all_trs:
                spans = tr.find_all("span")
                if len(spans) >= 2:
                    label = spans[0].get_text(strip=True)
                    value = spans[1].get_text(strip=True)
                    if label == "Check-in":
                        checkin = value.replace("\xa0", " ").strip()
                    elif label == "Check-out":
                        checkout = value.replace("\xa0", " ").strip()
                    elif label == "Country of Residence":
                        country_of_residence = value.strip()

            if checkin and checkout:
                try:
                    checkin_date = datetime.strptime(checkin, "%B %d, %Y")
                    checkout_date = datetime.strptime(checkout, "%B %d, %Y")
                    nights = (checkout_date - checkin_date).days
                except Exception:
                    nights = None
            else:
                continue

            if checkin_date < now:
                continue

            reservation = {
                "Platform": platform,
                "Booking ID": booking_id,
                "Guest Name": guest_name,
                "Total Price": total_price,
                "Rate Plan": rate_plan,
                "Room Type": room_type,
                "Unit Number": unit_number,
                "Number of Rooms": rooms,
                "Adults": adults,
                "Children": kids,
                "Check-in": checkin,
                "Check-out": checkout,
                "Duration of Stay": f"{nights} nights" if nights else "N/A",
                "Country of Residence": country_of_residence,
                "Last Updated Time": last_updated,
                "Status": status
            }

            reservation_dict[booking_id] = reservation
            #quite slow when comparing to retrieve the only the lastest change
            # if booking_id:
            #     existing = reservation_dict.get(booking_id)
            #     if not existing or last_updated_dt > datetime.strptime(existing["Last Updated Time"], "%Y-%m-%d %H:%M:%S"):
            #         reservation_dict[booking_id] = reservation

    imap.logout()

    with open(timestamp_file, "w") as f:
        f.write(now.strftime("%Y-%m-%d %H:%M:%S"))

    reservations = list(reservation_dict.values())
    os.makedirs(output_dir, exist_ok=True)
    with open(os.path.join(output_dir, "reservations.json"), "w", encoding="utf-8") as f:
        json.dump(reservations, f, indent=2, ensure_ascii=False)

    json_response = json.dumps({"reservations": reservations, "count": len(reservations)}, ensure_ascii=False)
    return Response(json_response, mimetype="application/json; charset=utf-8")

if __name__ == "__main__":
    app.run(debug=True)
