import os
import json
import base64
from flask import Flask, request, jsonify
from flask_cors import CORS
import anthropic
from openpyxl import Workbook
import PyPDF2
import io
from datetime import datetime, timedelta

app = Flask(__name__)
CORS(app, origins="*")

client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))

# ─────────────────────────────────────────────
# UTILITIES
# ─────────────────────────────────────────────

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

def date_to_excel_serial(date_str):
    """Convert DD/MM/YYYY string to Excel serial number."""
    try:
        d = datetime.strptime(date_str.strip(), "%d/%m/%Y")
        return (d - datetime(1899, 12, 30)).days
    except:
        return 0

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

# ─────────────────────────────────────────────
# HEADERS
# ─────────────────────────────────────────────

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

# ─────────────────────────────────────────────
# CLAUDE PROMPTS
# ─────────────────────────────────────────────

CONTRACT_EXTRACTION_PROMPT = """
You are a hotel rate sheet expert for Voyage Tours, a Dubai-based tour operator.

Extract rate data from this hotel contract PDF and return it as a single valid JSON object.

Return ONLY this JSON structure, no markdown, no explanation, no trailing commas:
{
  "hotel_name": "string",
  "reservation_date_from": "DD/MM/YYYY",
  "reservation_date_till": "DD/MM/YYYY",
  "meal_plans": ["Bed and Breakfast"],
  "supplement_rules": {
    "hb_per_adult": 0,
    "fb_per_adult": 0,
    "extra_bed_adult": 0,
    "extra_bed_child_under_6": 0,
    "extra_bed_child_6_to_12": 0,
    "child_meal_hb_under_6": 0,
    "child_meal_hb_6_to_12": 0,
    "child_meal_fb_under_6": 0,
    "child_meal_fb_6_to_12": 0
  },
  "room_seasons": [
    {
      "room": "string",
      "base_bb": 0,
      "season_begin": "DD/MM/YYYY",
      "season_end": "DD/MM/YYYY",
      "season_type": "Low",
      "res_date_from": "DD/MM/YYYY",
      "res_date_till": "DD/MM/YYYY"
    }
  ]
}

RATE EXTRACTION RULES:
- base_bb is the Double Occupancy (2A) BB rate per room per night
- If rates table has columns per occupancy (1A, 2A, 2A+1C, 3A), use the 2A column as base_bb
- One room_seasons entry per room type per season date range
- If a season has non-consecutive dates (e.g. High = Nov-May AND Sep-Oct), create a separate entry for each range

DATE RULES:
- reservation_date_from = contract start date (first valid stay date)
- reservation_date_till = contract end date (last valid stay date)
- res_date_from per season = same as top-level reservation_date_from
- res_date_till per season = same as top-level reservation_date_till — NEVER the season end date
- season_begin / season_end = actual stay check-in/check-out dates for that period
- All dates in DD/MM/YYYY format

SUPPLEMENT RULES:
- Add HB and FB to meal_plans only if explicitly offered in the contract
- Set any supplement to 0 if described as free or complimentary
- season_type must be exactly one of: "Low", "Shoulder", "High", "Peak"

OUTPUT: Valid JSON only. No trailing commas. No comments. No markdown fences.
"""


PROMOTION_EXTRACTION_PROMPT = """
You are a hotel rate sheet expert for Voyage Tours, a Dubai-based tour operator.

Extract all rate data from this hotel promotion/SPO PDF and return a single valid JSON object.

CRITICAL: Use the FINAL DISCOUNTED/PROMO RATE as base_rate — never the contracted rate.

Return ONLY this JSON structure, no markdown, no explanation, no trailing commas:
{
  "hotel_name": "string",
  "promo_code": "string",
  "reservation_date_from": "DD/MM/YYYY",
  "reservation_date_till": "DD/MM/YYYY",
  "mlos": 1,
  "mlos_till": 366,
  "meal_plans": ["Bed and Breakfast"],
  "supplement_rules": {
    "hb_per_adult": 0,
    "fb_per_adult": 0,
    "extra_bed_adult": 0,
    "extra_bed_child_under_6": 0,
    "extra_bed_child_6_to_12": 0,
    "child_meal_hb_under_6": 0,
    "child_meal_hb_6_to_12": 0,
    "child_meal_fb_under_6": 0,
    "child_meal_fb_6_to_12": 0
  },
  "room_seasons": [
    {
      "room": "string",
      "base_rate": 0,
      "meal_plan": "Bed and Breakfast",
      "season_begin": "DD/MM/YYYY",
      "season_end": "DD/MM/YYYY",
      "season_type": "Low",
      "res_date_from": "DD/MM/YYYY",
      "res_date_till": "DD/MM/YYYY"
    }
  ]
}

RATE EXTRACTION RULES:
- base_rate = the final selling rate after discount for that room and date range
- If the promo table shows many date rows per room type, create one room_seasons entry per room per date row
- meal_plan = the meal plan the base_rate is quoted at (usually "Bed and Breakfast")
- season_type must be exactly: "Low", "Shoulder", "High", or "Peak"

DATE RULES:
- reservation_date_from = booking open date for this promotion
- reservation_date_till = last date a booking can be made for this promotion
- res_date_from per season = same as top-level reservation_date_from
- res_date_till per season = same as top-level reservation_date_till — NOT the season_end date
- season_begin / season_end = the actual stay dates for that row
- All dates in DD/MM/YYYY format

OTHER RULES:
- mlos = minimum length of stay (default 1 if not stated)
- mlos_till = maximum length of stay (default 366 if not stated)
- Set supplements to 0 if free/complimentary or if not stated in promotion (fallback to contract context)
- OUTPUT: Valid JSON only. No trailing commas. No comments. No markdown.
"""


# ─────────────────────────────────────────────
# PRICE CALCULATION
# ─────────────────────────────────────────────

def calculate_price(base_bb, meal, occ, rules):
    """
    Calculate price for any occupancy and meal plan
    using supplement rules extracted from the PDF.
    """
    hb = rules.get("hb_per_adult", 45)
    fb = rules.get("fb_per_adult", 90)
    extra_bed_adult = rules.get("extra_bed_adult", 75)
    extra_bed_child_u6 = rules.get("extra_bed_child_under_6", 0)
    extra_bed_child_6 = rules.get("extra_bed_child_6_to_12", 50)
    child_hb_u6 = rules.get("child_meal_hb_under_6", 0)
    child_hb_6 = rules.get("child_meal_hb_6_to_12", 30)
    child_fb_u6 = rules.get("child_meal_fb_under_6", 0)
    child_fb_6 = rules.get("child_meal_fb_6_to_12", 60)

    if meal in ["Room Only", "RO"]:
        adult_meal = 0
        child_meal_u6 = 0
        child_meal_6 = 0
    elif meal in ["Bed and Breakfast", "BB"]:
        adult_meal = 0
        child_meal_u6 = 0
        child_meal_6 = 0
    elif meal in ["Half Board", "HB"]:
        adult_meal = hb
        child_meal_u6 = child_hb_u6
        child_meal_6 = child_hb_6
    elif meal in ["Full Board", "FB"]:
        adult_meal = fb
        child_meal_u6 = child_fb_u6
        child_meal_6 = child_fb_6
    else:
        adult_meal = 0
        child_meal_u6 = 0
        child_meal_6 = 0

    price = base_bb

    # Adult meal supplements
    total_adults = occ["adults"] + occ["adult_extra_beds"]
    price += total_adults * adult_meal

    # Adult extra bed fee
    price += occ["adult_extra_beds"] * extra_bed_adult

    # Child extra beds
    price += occ.get("child_paid_extra_under_6", 0) * (extra_bed_child_u6 + child_meal_u6)
    price += occ.get("child_paid_extra_6_to_12", 0) * (extra_bed_child_6 + child_meal_6)

    # Child sharing meal supplements
    price += occ.get("child_paid_sharing_under_6", 0) * child_meal_u6
    price += occ.get("child_paid_sharing_6_to_12", 0) * child_meal_6

    return round(price)

# ─────────────────────────────────────────────
# ROW EXPANSION
# ─────────────────────────────────────────────

def expand_contract_rates(hotel_name, room_seasons, meal_plans, occupancy_combinations, supplement_rules):
    rows = []
    for season in room_seasons:
        room = season["room"]
        base_bb = season["base_bb"]
        season_begin = date_to_excel_serial(season["season_begin"])
        season_end = date_to_excel_serial(season["season_end"])
        season_type = season["season_type"]
        res_date_from = date_to_excel_serial(season["res_date_from"])
        res_date_till = date_to_excel_serial(season["res_date_till"])

        for meal in meal_plans:
            for occ in occupancy_combinations:
                price = calculate_price(base_bb, meal, occ, supplement_rules)
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

def expand_promotion_rates(hotel_name, promo_code, room_seasons, meal_plans, occupancy_combinations, supplement_rules, mlos, mlos_till):
    rows = []
    for season in room_seasons:
        room = season["room"]
        base_rate = season["base_rate"]
        base_meal = season["meal_plan"]
        season_begin = date_to_excel_serial(season["season_begin"])
        season_end = date_to_excel_serial(season["season_end"])
        season_type = season["season_type"]
        res_date_from = date_to_excel_serial(season["res_date_from"])
        res_date_till = date_to_excel_serial(season["res_date_till"])

        for meal in meal_plans:
            for occ in occupancy_combinations:
                price = calculate_price(base_rate, meal, occ, supplement_rules)
                row = {
                    "SPO No": "",
                    "Price type": "Standard",
                    "Hotel": hotel_name,
                    "Room": room,
                    "Accommodation": occ["label"],
                    "Meal": meal,
                    "Hotel net price": price,
                    "Currency (code)": "AED",
                    "Market code": "",
                    "Season begin": season_begin,
                    "Season end": season_end,
                    "Days before check-in from": "",
                    "Reservation date from": res_date_from,
                    "Reservation date till": res_date_till,
                    "Check-in from": "",
                    "Check-in till": "",
                    "Check-out from": "",
                    "Staying nights from": mlos,
                    "Check-out till": "",
                    "Nights": 1,
                    "Nights from": "",
                    "Nights till": "",
                    "Number of markups": 1,
                    "Nights free": "",
                    "Season type": season_type,
                    "Days before check-in till": "",
                    "Staying nights till": mlos_till,
                    "Booking code": promo_code
                }
                rows.append([row.get(h, "") for h in PROMOTION_HEADERS])
    return rows

# ─────────────────────────────────────────────
# API ROUTES
# ─────────────────────────────────────────────

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
            max_tokens=8000,
            system=CONTRACT_EXTRACTION_PROMPT,
            messages=[
                {
                    "role": "user",
                    "content": f"Extract all rate data from this hotel contract:\n\n{contract_text}"
                }
            ]
        )

        raw_contract = clean_json_response(contract_response.content[0].text)
        contract_data = json.loads(raw_contract)

        contract_rows = expand_contract_rates(
            hotel_name=contract_data["hotel_name"],
            room_seasons=contract_data["room_seasons"],
            meal_plans=contract_data["meal_plans"],
            occupancy_combinations=contract_data["occupancy_combinations"],
            supplement_rules=contract_data["supplement_rules"]
        )

        contract_excel = generate_excel_from_data(contract_rows, CONTRACT_HEADERS)
        result = {"contract_excel": contract_excel}

        if promotion_file:
            promotion_text = extract_text_from_pdf(promotion_file)

            promotion_response = client.messages.create(
                model="claude-sonnet-4-6",
                max_tokens=8000,
                system=PROMOTION_EXTRACTION_PROMPT,
                messages=[
                    {
                        "role": "user",
                        "content": f"Extract all rate data from this promotion PDF. Contract is provided for context on supplement rules if not stated in promotion.\n\nPROMOTION:\n{promotion_text}\n\nCONTRACT CONTEXT:\n{contract_text}"
                    }
                ]
            )

            raw_promotion = clean_json_response(promotion_response.content[0].text)
            promotion_data = json.loads(raw_promotion)

            promotion_rows = expand_promotion_rates(
                hotel_name=promotion_data["hotel_name"],
                promo_code=promotion_data.get("promo_code", ""),
                room_seasons=promotion_data["room_seasons"],
                meal_plans=promotion_data["meal_plans"],
                occupancy_combinations=promotion_data["occupancy_combinations"],
                supplement_rules=promotion_data["supplement_rules"],
                mlos=promotion_data.get("mlos", 1),
                mlos_till=promotion_data.get("mlos_till", 366)
            )

            promotion_excel = generate_excel_from_data(promotion_rows, PROMOTION_HEADERS)
            result["promotion_excel"] = promotion_excel

        return jsonify(result)

    except json.JSONDecodeError as e:
        # Return the raw Claude response so the issue can be diagnosed
        raw = locals().get("raw_contract") or locals().get("raw_promotion") or "unavailable"
        return jsonify({
            "error": f"Failed to parse Claude response: {str(e)}",
            "raw_response": raw[:2000]  # first 2000 chars for debugging
        }), 500
    except Exception as e:
        import traceback
        return jsonify({"error": str(e), "trace": traceback.format_exc()[-1000:]}), 500

@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})

if __name__ == "__main__":
    app.run(debug=True)
