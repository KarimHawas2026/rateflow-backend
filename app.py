import os
import json
import base64
from flask import Flask, request, jsonify
from flask_cors import CORS
import anthropic
from openpyxl import Workbook
import PyPDF2
import fitz  # pymupdf
import io
from datetime import datetime, timedelta

app = Flask(__name__)
CORS(app, origins="*")

client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))

# ─────────────────────────────────────────────
# UTILITIES
# ─────────────────────────────────────────────

def extract_text_from_pdf(pdf_file):
    """Try text extraction first. If the PDF is image-based (scanned),
    return None so the caller knows to use vision instead."""
    pdf_bytes = pdf_file.read()
    pdf_file.seek(0)

    reader = PyPDF2.PdfReader(io.BytesIO(pdf_bytes))
    text = ""
    for page in reader.pages:
        text += (page.extract_text() or "") + "\n"

    if len(text.strip()) > 100:
        return text, None  # text OK, no images needed

    # Scanned PDF — render each page as an image for Claude vision
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    images = []
    for page in doc:
        mat = fitz.Matrix(2, 2)  # 2x zoom for readability
        pix = page.get_pixmap(matrix=mat)
        img_b64 = base64.standard_b64encode(pix.tobytes("png")).decode("utf-8")
        images.append(img_b64)
    doc.close()
    return None, images  # no text, images ready


def build_user_message_for_claude(text, images, prompt_suffix):
    """Build the user message content — text or vision depending on PDF type."""
    if text:
        return [{"type": "text", "text": f"{prompt_suffix}\n\n{text}"}]

    # Vision: send each page image + extraction instruction
    content = []
    for img_b64 in images:
        content.append({
            "type": "image",
            "source": {"type": "base64", "media_type": "image/png", "data": img_b64}
        })
    content.append({"type": "text", "text": prompt_suffix})
    return content

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

Your job is to read a hotel contract PDF and extract ALL rate data needed to build a complete rate sheet.

You must extract:
1. Hotel name
2. Contract validity dates (reservation_date_from and reservation_date_till)
3. For each room type and each season period: the BB SGL/DBL base rate
4. The supplement rules specific to this hotel (meal plan supplements per person, extra bed charges, child policy)
5. All available meal plans (Room Only, BB, HB, FB, etc.)
6. All valid occupancy combinations used by this hotel

Return ONLY this JSON structure, no markdown, no explanation:
{
  "hotel_name": "string",
  "reservation_date_from": "DD/MM/YYYY",
  "reservation_date_till": "DD/MM/YYYY",
  "meal_plans": ["Room Only", "Bed and Breakfast", "Half Board", "Full Board"],
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
      "season_type": "Low or Shoulder or High or Peak",
      "res_date_from": "DD/MM/YYYY",
      "res_date_till": "DD/MM/YYYY"
    }
  ],
  "occupancy_combinations": [
    {
      "label": "string",
      "adults": number,
      "adult_extra_beds": number,
      "child_free_sharing": number,
      "child_paid_sharing_under_6": number,
      "child_paid_sharing_6_to_12": number,
      "child_free_extra": number,
      "child_paid_extra_under_6": number,
      "child_paid_extra_6_to_12": number
    }
  ]
}

IMPORTANT RULES:
- All dates must be in DD/MM/YYYY format
- res_date_from per season = contract signing date or opening booking date
- res_date_till per season = that season's end date
- reservation_date_from = overall contract start (when bookings open)
- reservation_date_till = overall contract end date
- season_type must be exactly: "Low", "Shoulder", "High", or "Peak"
- If the hotel only has BB (no Room Only), do not include Room Only in meal_plans
- Extract supplement rules exactly as stated in the contract
- If a supplement is "free" set it to 0
- Return ONLY raw JSON. Nothing else.
"""

PROMOTION_EXTRACTION_PROMPT = """
You are a hotel rate sheet expert for Voyage Tours, a Dubai-based tour operator.

Your job is to read a hotel promotion/SPO PDF and extract ALL rate data needed to build a complete promotion rate sheet.

CRITICAL: Use the FINAL SELLING RATE or PROMO RATE as base rate. Never use contracted rates.

Return ONLY this JSON structure, no markdown, no explanation:
{
  "hotel_name": "string",
  "promo_code": "string",
  "reservation_date_from": "DD/MM/YYYY",
  "reservation_date_till": "DD/MM/YYYY",
  "mlos": number,
  "mlos_till": number,
  "meal_plans": ["Room Only", "Bed and Breakfast", "Half Board"],
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
      "season_type": "Low or Shoulder or High or Peak",
      "res_date_from": "DD/MM/YYYY",
      "res_date_till": "DD/MM/YYYY"
    }
  ],
  "occupancy_combinations": [
    {
      "label": "string",
      "adults": number,
      "adult_extra_beds": number,
      "child_free_sharing": number,
      "child_paid_sharing_under_6": number,
      "child_paid_sharing_6_to_12": number,
      "child_free_extra": number,
      "child_paid_extra_under_6": number,
      "child_paid_extra_6_to_12": number
    }
  ]
}

IMPORTANT RULES:
- All dates must be in DD/MM/YYYY format
- mlos = minimum length of stay (default 1 if not stated)
- mlos_till = maximum length of stay (default 366 if not stated)
- base_rate is the PROMO/FINAL SELLING RATE for that room and meal plan
- meal_plan per room_season is the base meal plan of the promo rate
- season_type must be exactly: "Low", "Shoulder", "High", or "Peak"
- Extract supplement rules exactly as stated in the promotion PDF
- Return ONLY raw JSON. Nothing else.
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

        contract_text, contract_images = extract_text_from_pdf(contract_file)

        contract_response = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=4000,
            system=CONTRACT_EXTRACTION_PROMPT,
            messages=[{
                "role": "user",
                "content": build_user_message_for_claude(
                    contract_text, contract_images,
                    "Extract all rate data from this hotel contract:"
                )
            }]
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
            promotion_text, promotion_images = extract_text_from_pdf(promotion_file)

            # Build context suffix — include contract text if available for supplement fallback
            context_note = "\n\nCONTRACT CONTEXT (for supplement rules if not in promotion):\n" + contract_text if contract_text else ""

            promotion_response = client.messages.create(
                model="claude-sonnet-4-6",
                max_tokens=4000,
                system=PROMOTION_EXTRACTION_PROMPT,
                messages=[{
                    "role": "user",
                    "content": build_user_message_for_claude(
                        promotion_text, promotion_images,
                        f"Extract all rate data from this promotion PDF.{context_note}"
                    )
                }]
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
        return jsonify({"error": f"Failed to parse Claude response: {str(e)}"}), 500
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})

if __name__ == "__main__":
    app.run(debug=True)
