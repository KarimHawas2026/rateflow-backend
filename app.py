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
    if text.startswith("```"):
        parts = text.split("```")
        if len(parts) >= 2:
            text = parts[1]
            if text.startswith("json"):
                text = text[4:]
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

CONTRACT_SYSTEM_PROMPT = """
You are a hotel rate sheet expert for Voyage Tours, a Dubai-based tour operator.

Your job is to read a hotel contract PDF and generate a complete rate sheet in Voyage Tours exact internal format.

SUPPLEMENT RULES (apply these exactly):
- Half Board (HB) = BB rate + AED 45 per adult per night
- Full Board (FB) = BB rate + AED 90 per adult per night
- 3rd Adult Extra Bed and Breakfast = base rate + AED 75
- Child (06-11.99) Extra Bed BB = base rate + AED 50
- Child (06-11.99) Extra Bed HB = base rate + AED 75
- Child (06-11.99) Extra Bed FB = base rate + AED 110
- Child (00-05.99) Extra Bed = FREE (same as base rate)
- Child (00-05.99) sharing = FREE (same as base rate)
- Child (06-11.99) sharing BB = same as base rate
- Child (06-11.99) sharing HB = base rate + AED 30
- Child (06-11.99) sharing FB = base rate + AED 60

OCCUPANCY COMBINATIONS to generate for EACH room type, EACH season, EACH meal plan (BB, HB, FB):
1ADL
2ADL
2ADL+1- ADULT EXTRA BED
3ADL
2ADL+1-CHILD EXTRA BED (00 - 05.99)
2ADL+1-CHILD EXTRA BED (06 - 11.99)
2ADL+1-CHILD SHARING (00 - 05.99)
2ADL+1-CHILD SHARING (06 - 11.99)
2ADL+2-CHILD SHARING (00 - 05.99)
2ADL+2-CHILD SHARING (06 - 11.99)
2ADL+1-CHILD EXTRA BED (00 - 05.99)+1-CHILD SHARING (00 - 05.99)
2ADL+1-CHILD EXTRA BED (06 - 11.99)+1-CHILD SHARING (00 - 05.99)
2ADL+1-CHILD EXTRA BED (06 - 11.99)+1-CHILD SHARE  (06 - 11.99)
2ADL+1-CHILD SHARING (00 - 05.99)+1-CHILD SHARING (06 - 11.99)
1ADL+1-CHILD SHARING (00 - 05.99)
1ADL+1-CHILD SHARING (06 - 11.99)
1ADL+2-CHILD SHARING (00 - 05.99)
1ADL+2-CHILD SHARING (06 - 11.99)
1ADL+1-CHILD SHARING (00 - 05.99)+1-CHILD SHARING (06 - 11.99)

PRICE CALCULATION RULES:
- 1ADL and 2ADL = same base BB rate
- 3ADL = same price as 2ADL+1 ADULT EXTRA BED = base + 75
- Two children 06-11.99 sharing = add 30 HB or 60 FB per child aged 06-11.99
- Mixed child ages: only add supplement for the 06-11.99 child

DATE SERIAL CONVERSION - use these exact values:
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

Use reservation date from = first date of contract validity
Use reservation date till = last date of contract validity

OUTPUT FORMAT - return ONLY this JSON, no markdown, no explanation:
{
  "hotel_name": "string",
  "rows": [
    {
      "Hotel": "string",
      "Room": "string",
      "Accommodation": "string",
      "Meal": "string",
      "Season begin": number,
      "Season end": number,
      "Reservation date from": number,
      "Reservation date till": number,
      "Nights": 1,
      "Hotel net price": number,
      "Number of markups": 1,
      "Currency (code)": "AED",
      "Currency": "Dirham",
      "Season type": "string",
      "Market code": "",
      "Price type": "Standard",
      "Staying nights from": 1,
      "Staying nights till": 366,
      "Booking code": ""
    }
  ]
}

Meal values must be exactly: "Bed and Breakfast", "Half Board", "Full Board"
Season type values must be exactly: "Low", "Shoulder", "High"
Return ONLY raw JSON. No markdown code blocks. No explanation. No extra text.
"""

PROMOTION_SYSTEM_PROMPT = """
You are a hotel rate sheet expert for Voyage Tours, a Dubai-based tour operator.

Your job is to read a hotel promotion PDF and generate a complete promotion rate sheet in Voyage Tours exact internal format.

IMPORTANT: Use the FINAL SELLING RATE from the promotion PDF as the base rate. Do not use contracted rates.

SUPPLEMENT RULES (apply these exactly):
- Half Board (HB) = Final Selling Rate + AED 45 per adult per night
- Full Board (FB) = Final Selling Rate + AED 90 per adult per night
- 3rd Adult Extra Bed and Breakfast = base rate + AED 75
- Child (06-11.99) Extra Bed BB = base rate + AED 50
- Child (06-11.99) Extra Bed HB = base rate + AED 75
- Child (06-11.99) Extra Bed FB = base rate + AED 110
- Child (00-05.99) Extra Bed = FREE (same as base rate)
- Child (00-05.99) sharing = FREE (same as base rate)
- Child (06-11.99) sharing BB = same as base rate
- Child (06-11.99) sharing HB = base rate + AED 30
- Child (06-11.99) sharing FB = base rate + AED 60

OCCUPANCY COMBINATIONS to generate for EACH room type, EACH season, EACH meal plan (BB, HB, FB):
1ADL
2ADL
2ADL+1- ADULT EXTRA BED
2ADL+1-CHILD EXTRA BED (00 - 05.99)
2ADL+1-CHILD EXTRA BED (06 - 11.99)
2ADL+1-CHILD SHARING (00 - 05.99)
2ADL+1-CHILD SHARING (06 - 11.99)
2ADL+2-CHILD SHARING (00 - 05.99)
2ADL+2-CHILD SHARING (06 - 11.99)
2ADL+1-CHILD EXTRA BED (00 - 05.99)+1-CHILD SHARING (00 - 05.99)
2ADL+1-CHILD EXTRA BED (06 - 11.99)+1-CHILD SHARING (00 - 05.99)
2ADL+1-CHILD EXTRA BED (06 - 11.99)+1-CHILD SHARE  (06 - 11.99)
2ADL+1-CHILD SHARING (00 - 05.99)+1-CHILD SHARING (06 - 11.99)
1ADL+1-CHILD SHARING (00 - 05.99)
1ADL+1-CHILD SHARING (06 - 11.99)
1ADL+2-CHILD SHARING (00 - 05.99)
1ADL+2-CHILD SHARING (06 - 11.99)
1ADL+1-CHILD SHARING (00 - 05.99)+1-CHILD SHARING (06 - 11.99)

DATE SERIAL CONVERSION - use these exact values:
1 Jun 2026 = 46174
14 Sep 2026 = 46279
15 Sep 2026 = 46280
30 Sep 2026 = 46295
1 Oct 2026 = 46296
2 Jan 2027 = 46384

Reservation date from = first booking date from the promotion
Reservation date till = last booking date from the promotion

OUTPUT FORMAT - return ONLY this JSON, no markdown, no explanation:
{
  "hotel_name": "string",
  "promo_code": "string",
  "rows": [
    {
      "SPO No": "",
      "Price type": "Standard",
      "Hotel": "string",
      "Room": "string",
      "Accommodation": "string",
      "Meal": "string",
      "Hotel net price": number,
      "Currency (code)": "AED",
      "Market code": "KPS",
      "Season begin": number,
      "Season end": number,
      "Days before check-in from": "",
      "Reservation date from": number,
      "Reservation date till": number,
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
      "Season type": "string",
      "Days before check-in till": "",
      "Staying nights till": 366,
      "Booking code": "string"
    }
  ]
}

Return ONLY raw JSON. No markdown code blocks. No explanation. No extra text.
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
            model="claude-opus-4-6",
            max_tokens=32000,
            system=CONTRACT_SYSTEM_PROMPT,
            messages=[
                {
                    "role": "user",
                    "content": f"Here is the hotel contract PDF text. Generate the complete rate sheet JSON:\n\n{contract_text}"
                }
            ]
        )

        raw_contract = clean_json_response(contract_response.content[0].text)
        contract_data = json.loads(raw_contract)
        contract_rows = []
        for row in contract_data["rows"]:
            contract_rows.append([row.get(h, "") for h in CONTRACT_HEADERS])

        contract_excel = generate_excel_from_data(contract_rows, CONTRACT_HEADERS)
        result = {"contract_excel": contract_excel}

        if promotion_file:
            promotion_text = extract_text_from_pdf(promotion_file)

            promotion_response = client.messages.create(
                model="claude-opus-4-6",
                max_tokens=32000,
                system=PROMOTION_SYSTEM_PROMPT,
                messages=[
                    {
                        "role": "user",
                        "content": f"Here is the hotel promotion PDF text. The contract is also provided for context.\n\nPROMOTION PDF:\n{promotion_text}\n\nCONTRACT CONTEXT:\n{contract_text}"
                    }
                ]
            )

            raw_promotion = clean_json_response(promotion_response.content[0].text)
            promotion_data = json.loads(raw_promotion)
            promotion_rows = []
            for row in promotion_data["rows"]:
                promotion_rows.append([row.get(h, "") for h in PROMOTION_HEADERS])

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
