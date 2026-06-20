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
# STANDARD OCCUPANCY COMBINATIONS
# Hardcoded — Voyage Tours uses these labels across all hotels.
# Claude only extracts supplement rules; Python handles occupancy.
# ─────────────────────────────────────────────

STANDARD_OCCUPANCY_COMBINATIONS = [
    # Adults only
    {"label": "1ADL",         "adults": 1, "adult_extra_beds": 0, "child_free_sharing": 0, "child_paid_sharing_under_6": 0, "child_paid_sharing_6_to_12": 0, "child_free_extra": 0, "child_paid_extra_under_6": 0, "child_paid_extra_6_to_12": 0},
    {"label": "2ADL",         "adults": 2, "adult_extra_beds": 0, "child_free_sharing": 0, "child_paid_sharing_under_6": 0, "child_paid_sharing_6_to_12": 0, "child_free_extra": 0, "child_paid_extra_under_6": 0, "child_paid_extra_6_to_12": 0},
    {"label": "3ADL",         "adults": 2, "adult_extra_beds": 1, "child_free_sharing": 0, "child_paid_sharing_under_6": 0, "child_paid_sharing_6_to_12": 0, "child_free_extra": 0, "child_paid_extra_under_6": 0, "child_paid_extra_6_to_12": 0},
    # 2 Adults + children sharing
    {"label": "2ADL+1-CHILD SHARING",   "adults": 2, "adult_extra_beds": 0, "child_free_sharing": 1, "child_paid_sharing_under_6": 0, "child_paid_sharing_6_to_12": 0, "child_free_extra": 0, "child_paid_extra_under_6": 0, "child_paid_extra_6_to_12": 0},
    {"label": "2ADL+1-CHILD06 SHARING", "adults": 2, "adult_extra_beds": 0, "child_free_sharing": 0, "child_paid_sharing_under_6": 1, "child_paid_sharing_6_to_12": 0, "child_free_extra": 0, "child_paid_extra_under_6": 0, "child_paid_extra_6_to_12": 0},
    {"label": "2ADL+1-CHILD12 SHARING", "adults": 2, "adult_extra_beds": 0, "child_free_sharing": 0, "child_paid_sharing_under_6": 0, "child_paid_sharing_6_to_12": 1, "child_free_extra": 0, "child_paid_extra_under_6": 0, "child_paid_extra_6_to_12": 0},
    {"label": "2ADL+2-CHILD SHARING",   "adults": 2, "adult_extra_beds": 0, "child_free_sharing": 2, "child_paid_sharing_under_6": 0, "child_paid_sharing_6_to_12": 0, "child_free_extra": 0, "child_paid_extra_under_6": 0, "child_paid_extra_6_to_12": 0},
    {"label": "2ADL+1CH06+1CH12 SHARING","adults": 2, "adult_extra_beds": 0, "child_free_sharing": 0, "child_paid_sharing_under_6": 1, "child_paid_sharing_6_to_12": 1, "child_free_extra": 0, "child_paid_extra_under_6": 0, "child_paid_extra_6_to_12": 0},
    {"label": "2ADL+2-CHILD06 SHARING", "adults": 2, "adult_extra_beds": 0, "child_free_sharing": 0, "child_paid_sharing_under_6": 2, "child_paid_sharing_6_to_12": 0, "child_free_extra": 0, "child_paid_extra_under_6": 0, "child_paid_extra_6_to_12": 0},
    {"label": "2ADL+2-CHILD12 SHARING", "adults": 2, "adult_extra_beds": 0, "child_free_sharing": 0, "child_paid_sharing_under_6": 0, "child_paid_sharing_6_to_12": 2, "child_free_extra": 0, "child_paid_extra_under_6": 0, "child_paid_extra_6_to_12": 0},
    # 2 Adults + children on extra bed
    {"label": "2ADL+1-CHILD06 EXTRA",   "adults": 2, "adult_extra_beds": 0, "child_free_sharing": 0, "child_paid_sharing_under_6": 0, "child_paid_sharing_6_to_12": 0, "child_free_extra": 0, "child_paid_extra_under_6": 1, "child_paid_extra_6_to_12": 0},
    {"label": "2ADL+1-CHILD12 EXTRA",   "adults": 2, "adult_extra_beds": 0, "child_free_sharing": 0, "child_paid_sharing_under_6": 0, "child_paid_sharing_6_to_12": 0, "child_free_extra": 0, "child_paid_extra_under_6": 0, "child_paid_extra_6_to_12": 1},
    {"label": "2ADL+2-CHILD06 EXTRA",   "adults": 2, "adult_extra_beds": 0, "child_free_sharing": 0, "child_paid_sharing_under_6": 0, "child_paid_sharing_6_to_12": 0, "child_free_extra": 0, "child_paid_extra_under_6": 2, "child_paid_extra_6_to_12": 0},
    {"label": "2ADL+2-CHILD12 EXTRA",   "adults": 2, "adult_extra_beds": 0, "child_free_sharing": 0, "child_paid_sharing_under_6": 0, "child_paid_sharing_6_to_12": 0, "child_free_extra": 0, "child_paid_extra_under_6": 0, "child_paid_extra_6_to_12": 2},
    {"label": "2ADL+1CH06+1CH12 EXTRA", "adults": 2, "adult_extra_beds": 0, "child_free_sharing": 0, "child_paid_sharing_under_6": 0, "child_paid_sharing_6_to_12": 0, "child_free_extra": 0, "child_paid_extra_under_6": 1, "child_paid_extra_6_to_12": 1},
    # 1 Adult + children on extra bed
    {"label": "1ADL+1-CHILD06 EXTRA",   "adults": 1, "adult_extra_beds": 0, "child_free_sharing": 0, "child_paid_sharing_under_6": 0, "child_paid_sharing_6_to_12": 0, "child_free_extra": 0, "child_paid_extra_under_6": 1, "child_paid_extra_6_to_12": 0},
    {"label": "1ADL+1-CHILD12 EXTRA",   "adults": 1, "adult_extra_beds": 0, "child_free_sharing": 0, "child_paid_sharing_under_6": 0, "child_paid_sharing_6_to_12": 0, "child_free_extra": 0, "child_paid_extra_under_6": 0, "child_paid_extra_6_to_12": 1},
]

# ─────────────────────────────────────────────
# CLAUDE PROMPTS
# ─────────────────────────────────────────────

CONTRACT_EXTRACTION_PROMPT = """
You are a hotel rate sheet expert for Voyage Tours, a Dubai-based tour operator.

Your job is to read a hotel contract PDF and extract rate data needed to build a complete rate sheet.

You must extract:
1. Hotel name (single specific property only — not a combined name if multiple hotels are in the PDF)
2. Contract validity dates (when bookings open and the last travel/stay date)
3. For each room type and each season period: the BB DBL base rate (double room, bed & breakfast, per room per night)
4. The supplement rules: meal plan add-ons per person, extra bed charges, child policy
5. All available meal plans offered

Return ONLY this JSON structure, no markdown, no explanation:
{
  "hotel_name": "string",
  "reservation_date_from": "DD/MM/YYYY",
  "reservation_date_till": "DD/MM/YYYY",
  "meal_plans": ["Bed and Breakfast", "Half Board", "Full Board"],
  "supplement_rules": {
    "hb_per_adult": number,
    "fb_per_adult": number,
    "extra_bed_adult": number,
    "extra_bed_child_under_6": number,
    "extra_bed_child_6_to_12": number,
    "child_meal_hb_under_6": number,
    "child_meal_hb_6_to_12": number,
    "child_meal_fb_under_6": number,
    "child_meal_fb_6_to_12": number
  },
  "room_seasons": [
    {
      "room": "string",
      "base_bb": number,
      "season_begin": "DD/MM/YYYY",
      "season_end": "DD/MM/YYYY",
      "season_type": "Low or Shoulder or High",
      "res_date_from": "DD/MM/YYYY",
      "res_date_till": "DD/MM/YYYY"
    }
  ]
}

DATE RULES — read carefully:
- reservation_date_from (top level) = the date bookings open for the whole contract
- reservation_date_till (top level) = the LAST day of the entire contract (latest season_end date)
- res_date_from per season = same as top-level reservation_date_from (bookings open the same day for all seasons)
- res_date_till per season = ALWAYS equal to the top-level reservation_date_till (the full contract end date, NOT the season end date)
- season_begin / season_end = the actual stay dates for that season period
- All dates must be in DD/MM/YYYY format exactly

OTHER RULES:
- season_type must be exactly: "Low", "Shoulder", or "High"
- Only include "Room Only" in meal_plans if the contract explicitly offers it as a base plan
- Extract supplement rules exactly as stated in the contract; set to 0 if "free" or "complimentary"
- Return ONLY raw JSON. No explanation, no markdown.
"""

PROMOTION_EXTRACTION_PROMPT = """
You are a hotel rate sheet expert for Voyage Tours, a Dubai-based tour operator.

Your job is to read a hotel promotion/SPO PDF and extract rate data needed to build a complete promotion rate sheet.

CRITICAL: Use the FINAL SELLING RATE or PROMO RATE as base rate. Never use contracted rates.

Return ONLY this JSON structure, no markdown, no explanation:
{
  "hotel_name": "string",
  "promo_code": "string",
  "reservation_date_from": "DD/MM/YYYY",
  "reservation_date_till": "DD/MM/YYYY",
  "mlos": number,
  "mlos_till": number,
  "meal_plans": ["Bed and Breakfast", "Half Board"],
  "supplement_rules": {
    "hb_per_adult": number,
    "fb_per_adult": number,
    "extra_bed_adult": number,
    "extra_bed_child_under_6": number,
    "extra_bed_child_6_to_12": number,
    "child_meal_hb_under_6": number,
    "child_meal_hb_6_to_12": number,
    "child_meal_fb_under_6": number,
    "child_meal_fb_6_to_12": number
  },
  "room_seasons": [
    {
      "room": "string",
      "base_rate": number,
      "meal_plan": "string",
      "season_begin": "DD/MM/YYYY",
      "season_end": "DD/MM/YYYY",
      "season_type": "Low or Shoulder or High",
      "res_date_from": "DD/MM/YYYY",
      "res_date_till": "DD/MM/YYYY"
    }
  ]
}

DATE RULES — read carefully:
- reservation_date_from (top level) = the date bookings open for this promotion
- reservation_date_till (top level) = the LAST booking date for this promotion (last day to book, not last day to stay)
- res_date_from per season = same as top-level reservation_date_from
- res_date_till per season = ALWAYS equal to the top-level reservation_date_till (NOT the season end date)
- season_begin / season_end = the actual stay dates for that season/period
- All dates must be in DD/MM/YYYY format exactly

OTHER RULES:
- mlos = minimum length of stay nights (default 1 if not stated)
- mlos_till = maximum length of stay nights (default 366 if not stated)
- base_rate is the PROMO/FINAL SELLING RATE for that room and meal plan
- meal_plan per room_season is the base meal plan the promo rate is quoted at
- season_type must be exactly: "Low", "Shoulder", or "High"
- If supplement rules are not in the promotion, use the contract context provided
- Set any supplement to 0 if described as "free" or "complimentary"
- Return ONLY raw JSON. No explanation, no markdown.
"""

# ─────────────────────────────────────────────
# PRICE CALCULATION
# ─────────────────────────────────────────────

def validate_and_fix_dates(data, is_promotion=False):
    """
    Validate extracted dates and fix the most common Claude mistake:
    res_date_till per season being set to season_end instead of contract end.
    """
    contract_end = data.get("reservation_date_till", "")

    for season in data.get("room_seasons", []):
        season_end = season.get("season_end", "")
        res_till = season.get("res_date_till", "")

        # If res_date_till equals season_end (the common mistake), fix it to contract end
        if res_till == season_end and contract_end:
            season["res_date_till"] = contract_end

        # If res_date_till is missing or zero, use contract end
        if not res_till:
            season["res_date_till"] = contract_end

        # res_date_from should never be empty
        contract_start = data.get("reservation_date_from", "")
        if not season.get("res_date_from") and contract_start:
            season["res_date_from"] = contract_start

    return data


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

    # Adult meal supplements (all adults in room including extra bed adults)
    total_adults = occ["adults"] + occ["adult_extra_beds"]
    price += total_adults * adult_meal

    # Adult extra bed fee
    price += occ["adult_extra_beds"] * extra_bed_adult

    # Children on extra beds
    num_child_extra_u6 = occ.get("child_paid_extra_under_6", 0)
    num_child_extra_6 = occ.get("child_paid_extra_6_to_12", 0)

    # First child extra bed: extra_bed_fee + meal supplement
    # Second+ child extra bed: extra_bed_fee is charged again (second_paid_child_extra_bed)
    price += num_child_extra_u6 * (extra_bed_child_u6 + child_meal_u6)
    price += num_child_extra_6 * (extra_bed_child_6 + child_meal_6)

    # When 2 children aged 06-11.99 are on extra bed, the second child incurs an
    # additional extra bed fee on top of the one already counted above.
    if num_child_extra_6 >= 2:
        price += extra_bed_child_6  # second child extra bed surcharge

    # Children sharing (meal supplement only, no extra bed fee)
    price += occ.get("child_paid_sharing_under_6", 0) * child_meal_u6
    price += occ.get("child_paid_sharing_6_to_12", 0) * child_meal_6

    return round(price)

# ─────────────────────────────────────────────
# ROW EXPANSION
# ─────────────────────────────────────────────

def expand_contract_rates(hotel_name, room_seasons, meal_plans, supplement_rules):
    rows = []
    occupancy_combinations = STANDARD_OCCUPANCY_COMBINATIONS
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

def expand_promotion_rates(hotel_name, promo_code, room_seasons, meal_plans, supplement_rules, mlos, mlos_till):
    rows = []
    occupancy_combinations = STANDARD_OCCUPANCY_COMBINATIONS
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
            max_tokens=4000,
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
        contract_data = validate_and_fix_dates(contract_data)

        contract_rows = expand_contract_rates(
            hotel_name=contract_data["hotel_name"],
            room_seasons=contract_data["room_seasons"],
            meal_plans=contract_data["meal_plans"],
            supplement_rules=contract_data["supplement_rules"]
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
                        "content": f"Extract all rate data from this promotion PDF. Contract is provided for context on supplement rules if not stated in promotion.\n\nPROMOTION:\n{promotion_text}\n\nCONTRACT CONTEXT:\n{contract_text}"
                    }
                ]
            )

            raw_promotion = clean_json_response(promotion_response.content[0].text)
            promotion_data = json.loads(raw_promotion)
            promotion_data = validate_and_fix_dates(promotion_data, is_promotion=True)

            promotion_rows = expand_promotion_rates(
                hotel_name=promotion_data["hotel_name"],
                promo_code=promotion_data.get("promo_code", ""),
                room_seasons=promotion_data["room_seasons"],
                meal_plans=promotion_data["meal_plans"],
                supplement_rules=promotion_data["supplement_rules"],
                mlos=promotion_data.get("mlos", 1),
                mlos_till=promotion_data.get("mlos_till", 366)
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
