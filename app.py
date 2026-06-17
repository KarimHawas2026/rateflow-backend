import os
import json
import base64
from flask import Flask, request, jsonify
from flask_cors import CORS
import anthropic
from openpyxl import Workbook
import PyPDF2
import io

app = Flask(__name__)
CORS(app, origins="*")

client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))

def extract_text_from_pdf(pdf_file):
    reader = PyPDF2.PdfReader(pdf_file)
    text = ""
    for page in reader.pages:
        text += page.extract_text() + "\n"
    return text

def clean_json_response(text):
    text = text.strip()
    if "```" in text:
        parts = text.split("```")
        for part in parts:
            part = part.strip()
            if part.startswith("json"):
                part = part[4:].strip()
            if part.startswith("{") or part.startswith("["):
                return part.strip()
    return text.strip()

def generate_excel_from_data(rows, headers):
    wb = Workbook()
    ws = wb.active
    ws.append(headers)
    for row in rows:
        ws.append(row)
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return base64.b64encode(buffer.read()).decode("utf-8")

CONTRACT_HEADERS = [
    "Hotel", "Room", "Accommodation", "Meal",
    "Season begin", "Season end",
    "Reservation date from", "Reservation date till",
    "Nights", "Hotel net price", "Number of markups",
    "Currency (code)", "Currency", "Season type",
    "Market code", "Price type",
    "Staying nights from", "Staying nights till", "Booking code"
]

PROMOTION_HEADERS = [
    "SPO No", "Price type", "Hotel", "Room", "Accommodation", "Meal",
    "Hotel net price", "Currency (code)", "Market code",
    "Season begin", "Season end",
    "Days before check-in from", "Reservation date from", "Reservation date till",
    "Check-in from", "Check-in till", "Check-out from",
    "Staying nights from", "Check-out till", "Nights",
    "Nights from", "Nights till", "Number of markups",
    "Nights free", "Season type", "Days before check-in till",
    "Staying nights till", "Booking code"
]

OCCUPANCY_COMBINATIONS = [
    {"label": "1ADL", "adult_extra": 0, "child_free_sharing": 0, "child_paid_sharing": 0, "child_free_extra": 0, "child_paid_extra": 0},
    {"label": "2ADL", "adult_extra": 0, "child_free_sharing": 0, "child_paid_sharing": 0, "child_free_extra": 0, "child_paid_extra": 0},
    {"label": "2ADL+1- ADULT EXTRA BED", "adult_extra": 1, "child_free_sharing": 0, "child_paid_sharing": 0, "child_free_extra": 0, "child_paid_extra": 0},
    {"label": "3ADL", "adult_extra": 1, "child_free_sharing": 0, "child_paid_sharing": 0, "child_free_extra": 0, "child_paid_extra": 0},
    {"label": "2ADL+1-CHILD EXTRA BED (00 - 05.99)", "adult_extra": 0, "child_free_sharing": 0, "child_paid_sharing": 0, "child_free_extra": 1, "child_paid_extra": 0},
    {"label": "2ADL+1-CHILD EXTRA BED (06 - 11.99)", "adult_extra": 0, "child_free_sharing": 0, "child_paid_sharing": 0, "child_free_extra": 0, "child_paid_extra": 1},
    {"label": "2ADL+1-CHILD SHARING (00 - 05.99)", "adult_extra": 0, "child_free_sharing": 1, "child_paid_sharing": 0, "child_free_extra": 0, "child_paid_extra": 0},
    {"label": "2ADL+1-CHILD SHARING (06 - 11.99)", "adult_extra": 0, "child_free_sharing": 0, "child_paid_sharing": 1, "child_free_extra": 0, "child_paid_extra": 0},
    {"label": "2ADL+2-CHILD SHARING (00 - 05.99)", "adult_extra": 0, "child_free_sharing": 2, "child_paid_sharing": 0, "child_free_extra": 0, "child_paid_extra": 0},
    {"label": "2ADL+2-CHILD SHARING (06 - 11.99)", "adult_extra": 0, "child_free_sharing": 0, "child_paid_sharing": 2, "child_free_extra": 0, "child_paid_extra": 0},
    {"label": "2ADL+1-CHILD EXTRA BED (00 - 05.99)+1-CHILD SHARING (00 - 05.99)", "adult_extra": 0, "child_free_sharing": 1, "child_paid_sharing": 0, "child_free_extra": 1, "child_paid_extra": 0},
    {"label": "2ADL+1-CHILD EXTRA BED (06 - 11.99)+1-CHILD SHARING (00 - 05.99)", "adult_extra": 0, "child_free_sharing": 1, "child_paid_sharing": 0, "child_free_extra": 0, "child_paid_extra": 1},
    {"label": "2ADL+1-CHILD EXTRA BED (06 - 11.99)+1-CHILD SHARE  (06 - 11.99)", "adult_extra": 0, "child_free_sharing": 0, "child_paid_sharing": 1, "child_free_extra": 0, "child_paid_extra": 1},
    {"label": "2ADL+1-CHILD SHARING (00 - 05.99)+1-CHILD SHARING (06 - 11.99)", "adult_extra": 0, "child_free_sharing": 1, "child_paid_sharing": 1, "child_free_extra": 0, "child_paid_extra": 0},
    {"label": "1ADL+1-CHILD SHARING (00 - 05.99)", "adult_extra": 0, "child_free_sharing": 1, "child_paid_sharing": 0, "child_free_extra": 0, "child_paid_extra": 0},
    {"label": "1ADL+1-CHILD SHARING (06 - 11.99)", "adult_extra": 0, "child_free_sharing": 0, "child_paid_sharing": 1, "child_free_extra": 0, "child_paid_extra": 0},
    {"label": "1ADL+2-CHILD SHARING (00 - 05.99)", "adult_extra": 0, "child_free_sharing": 2, "child_paid_sharing": 0, "child_free_extra": 0, "child_paid_extra": 0},
    {"label": "1ADL+2-CHILD SHARING (06 - 11.99)", "adult_extra": 0, "child_free_sharing": 0, "child_paid_sharing": 2, "child_free_extra": 0, "child_paid_extra": 0},
    {"label": "1ADL+1-CHILD SHARING (00 - 05.99)+1-CHILD SHARING (06 - 11.99)", "adult_extra": 0, "child_free_sharing": 1, "child_paid_sharing": 1, "child_free_extra": 0, "child_paid_extra": 0},
]

MEAL_PLANS = ["Bed and Breakfast", "Half Board", "Full Board"]

def calculate_price(base_bb, meal, occ):
    if meal == "Bed and Breakfast":
        meal_add = 0
        child_paid_meal = 0
    elif meal == "Half Board":
        meal_add = 45
        child_paid_meal = 30
    else:
        meal_add = 90
        child_paid_meal = 60

    price = base_bb
    price += occ["adult_extra"] * 75
    price += meal_add
    price += occ["child_paid_sharing"] * child_paid_meal
    price += occ["child_paid_extra"] * (50 + child_paid_meal)

    return round(price)

def expand_rates(hotel_name, room_seasons, res_date_from, res_date_till, is_promotion=False, promo_code="", market_code=""):
    rows = []
    for season in room_seasons:
        room = season["room"]
        base_bb = season["base_bb"]
        season_begin = season["season_begin"]
        season_end = season["season_end"]
        season_type = season["season_type"]

        for meal in MEAL_PLANS:
            for occ in OCCUPANCY_COMBINATIONS:
                price = calculate_price(base_bb, meal, occ)

                if is_promotion:
                    row = {
                        "SPO No": "",
                        "Price type": "Standard",
                        "Hotel": hotel_name,
                        "Room": room,
                        "Accommodation": occ["label"],
                        "Meal": meal,
                        "Hotel net price": price,
                        "Currency (code)": "AED",
                        "Market code": market_code or "KPS",
                        "Season begin": season_begin,
                        "Season end": season_end,
                        "Days before check-in from": "",
                        "Reservation date from": res_date_from,
                        "Reservation date till": res_date_till,
                        "Check-in from": "",
                        "Check-in till": "",
                        "Check-out from": "",
                        "Staying nights from": 1,
                        "Check-out till": "",
                        "Nights": 1,
                        "Nights from": "",
                        "Nights till": "",
                        "Number of markups": 1,
                        "Nights free": "",
                        "Season type": season_type,
                        "Days before check-in till": "",
                        "Staying nights till": 366,
                        "Booking code": promo_code
                    }
                    rows.append([row.get(h, "") for h in PROMOTION_HEADERS])
                else:
                    row = {
                        "Hotel": hotel_name,
                        "Room": room,
                        "Accommodation": occ["label"],
                        "Meal": meal,
                        "Season begin": season_begin,
                        "Season end": season_end,
                        "Reservation date from": res_date_from,
                        "Reservation date till": res_date_till,
                        "Nights": 1,
                        "Hotel net price": price,
                        "Number of markups": 1,
                        "Currency (code)": "AED",
                        "Currency": "Dirham",
                        "Season type": season_type,
                        "Market code": "",
                        "Price type": "Standard",
                        "Staying nights from": 1,
                        "Staying nights till": 366,
                        "Booking code": ""
                    }
                    rows.append([row.get(h, "") for h in CONTRACT_HEADERS])

    return rows

CONTRACT_EXTRACTION_PROMPT = """
You are a hotel rate sheet expert. Extract the rate data from this hotel contract PDF.

Return ONLY a JSON object in this exact format with no markdown, no explanation:
{
  "hotel_name": "string",
  "reservation_date_from": number,
  "reservation_date_till": number,
  "room_seasons": [
    {
      "room": "string",
      "base_bb": number,
      "season_begin": number,
      "season_end": number,
      "season_type": "Low or Shoulder or High"
    }
  ]
}

RULES:
- Extract each room type for each season period as a separate entry
- base_bb is the BB SGL/DBL rate for that room and season
- Convert all dates to Excel serial numbers (days since Jan 1 1900)
- reservation_date_from and reservation_date_till are the contract validity period
- season_type must be exactly: "Low", "Shoulder", or "High"

COMMON DATE SERIALS:
3 Jan 2026 = 46025
16 Feb 2026 = 46069
17 Feb 2026 = 46070
18 Mar 2026 = 46099
19 Mar 2026 = 46100
30 Apr 2026 = 46142
1 May 2026 = 46143
14 Sep 2026 = 46279
15 Sep 2026 = 46280
30 Sep 2026 = 46295
1 Oct 2026 = 46296
2 Jan 2027 = 46384
3 Jan 2027 = 46385

Return ONLY raw JSON. Nothing else.
"""

PROMOTION_EXTRACTION_PROMPT = """
You are a hotel rate sheet expert. Extract the rate data from this hotel promotion PDF.

IMPORTANT: Extract the FINAL SELLING RATE as base_bb, not the contracted rate.

Return ONLY a JSON object in this exact format with no markdown, no explanation:
{
  "hotel_name": "string",
  "promo_code": "string",
  "reservation_date_from": number,
  "reservation_date_till": number,
  "room_seasons": [
    {
      "room": "string",
      "base_bb": number,
      "season_begin": number,
      "season_end": number,
      "season_type": "Low or Shoulder or High"
    }
  ]
}

RULES:
- Extract each room type for each season period as a separate entry
- base_bb is the FINAL SELLING RATE (not contracted rate) for BB SGL/DBL
- Convert all dates to Excel serial numbers
- reservation_date_from and reservation_date_till are the booking date window
- season_type must be exactly: "Low", "Shoulder", or "High"

COMMON DATE SERIALS:
1 Jun 2026 = 46174
30 Jun 2026 = 46203
14 Sep 2026 = 46279
15 Sep 2026 = 46280
30 Sep 2026 = 46295
1 Oct 2026 = 46296
2 Jan 2027 = 46384

Return ONLY raw JSON. Nothing else.
"""

@app.route("/api/process", methods=["POST"])
def process_pdfs():
    try:
        contract_file = request.files.get("contract")
        promotion_file = request.files.get("promotion")

        if not contract_file:
            return jsonify({"error": "Contract PDF is required"}), 400

        contract_text = extract_text_from_pdf(contract_file)

        contract_response = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=4000,
            system=CONTRACT_EXTRACTION_PROMPT,
            messages=[
                {
                    "role": "user",
                    "content": f"Extract the rate data from this contract:\n\n{contract_text}"
                }
            ]
        )

        raw_contract = clean_json_response(contract_response.content[0].text)
        contract_data = json.loads(raw_contract)

        contract_rows = expand_rates(
            hotel_name=contract_data["hotel_name"],
            room_seasons=contract_data["room_seasons"],
            res_date_from=contract_data["reservation_date_from"],
            res_date_till=contract_data["reservation_date_till"],
            is_promotion=False
        )

        contract_excel = generate_excel_from_data(contract_rows, CONTRACT_HEADERS)
        result = {"contract_excel": contract_excel}

        if promotion_file:
            promotion_text = extract_text_from_pdf(promotion_file)

            promotion_response = client.messages.create(
                model="claude-sonnet-4-6",
                max_tokens=4000,
                system=PROMOTION_EXTRACTION_PROMPT,
                messages=[
                    {
                        "role": "user",
                        "content": f"Extract the rate data from this promotion:\n\n{promotion_text}"
                    }
                ]
            )

            raw_promotion = clean_json_response(promotion_response.content[0].text)
            promotion_data = json.loads(raw_promotion)

            promotion_rows = expand_rates(
                hotel_name=promotion_data["hotel_name"],
                room_seasons=promotion_data["room_seasons"],
                res_date_from=promotion_data["reservation_date_from"],
                res_date_till=promotion_data["reservation_date_till"],
                is_promotion=True,
                promo_code=promotion_data.get("promo_code", ""),
                market_code="KPS"
            )

            promotion_excel = generate_excel_from_data(promotion_rows, PROMOTION_HEADERS)
            result["promotion_excel"] = promotion_excel

        return jsonify(result)

    except json.JSONDecodeError as e:
        return jsonify({"error": f"Failed to parse Claude response: {str(e)}"}), 500
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})

if __name__ == "__main__":
    app.run(debug=True)
