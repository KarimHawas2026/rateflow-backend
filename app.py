import os
import json
import base64
from flask import Flask, request, jsonify
from flask_cors import CORS
import anthropic
import openpyxl
from openpyxl import Workbook
import PyPDF2
import io

app = Flask(__name__)
CORS(app)

client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))

def extract_text_from_pdf(pdf_file):
    reader = PyPDF2.PdfReader(pdf_file)
    text = ""
    for page in reader.pages:
        text += page.extract_text() + "\n"
    return text

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

Your job is to read a hotel contract PDF and generate a complete rate sheet in Voyage Tours' exact internal format.

SUPPLEMENT RULES (apply these exactly):
- Half Board (HB) = BB rate + AED 45 per adult per night
- Full Board (FB) = BB rate + AED 90 per adult per night
- 3rd Adult Extra Bed and Breakfast = base rate + AED 75
- Child (06-11.99) Extra Bed = base rate + AED 50
- Child (00-05.99) Extra Bed = FREE (same as base rate)
- Child sharing existing bed = FREE (same as base rate)
- Child (06-11.99) HB supplement = AED 30 per night
- Child (06-11.99) FB supplement = AED 60 per night

OCCUPANCY COMBINATIONS to generate for each room type, each season, each meal plan:
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
- 1ADL and 2ADL = same base rate (SGL/DBL)
- 3ADL = same as 2ADL+1 ADULT EXTRA BED
- Children 00-05.99 sharing = no extra charge
- Children 06-11.99 sharing = no extra charge for BB, add AED 30 for HB, add AED 60 for FB
- Child extra bed 00-05.99 = no extra charge
- Child extra bed 06-11.99 = add AED 50 for BB, add AED 75 for HB, add AED 110 for FB

DATE SERIAL CONVERSION:
Convert all dates to Excel serial numbers (days since 1 January 1900).
Example: 1 May 2026 = 46143, 14 Sep 2026 = 46279, 1 Oct 2026 = 46296, 2 Jan 2027 = 46384

OUTPUT FORMAT:
Return a JSON object with this exact structure:
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
Return ONLY the JSON. No explanation. No markdown. No extra text.
"""

PROMOTION_SYSTEM_PROMPT = """
You are a hotel rate sheet expert for Voyage Tours, a Dubai-based tour operator.

Your job is to read a hotel promotion PDF and generate a complete promotion rate sheet in Voyage Tours exact internal format.

IMPORTANT: Use the FINAL SELLING RATE from the promotion PDF, not the contracted rates.

SUPPLEMENT RULES (apply these exactly):
- Half Board (HB) = Final Selling Rate + AED 45 per adult per night
- Full Board (FB) = Final Selling Rate + AED 90 per adult per night
- 3rd Adult Extra Bed and Breakfast = base rate + AED 75
- Child (06-11.99) Extra Bed = base rate + AED 50
- Child (00-05.99) Extra Bed = FREE (same as base rate)
- Child sharing existing bed = FREE (same as base rate)
- Child (06-11.99) HB supplement = AED 30 per night
- Child (06-11.99) FB supplement = AED 60 per night

OCCUPANCY COMBINATIONS to generate for each room type, each season, each meal plan:
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

DATE SERIAL CONVERSION:
Convert all dates to Excel serial numbers.
Example: 1 June 2026 = 46174, 30 June 2026 = 46203, 14 Sep 2026 = 46279

OUTPUT FORMAT:
Return a JSON object with this exact structure:
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

Return ONLY the JSON. No explanation. No markdown. No extra text.
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
            max_tokens=8000,
            system=CONTRACT_SYSTEM_PROMPT,
            messages=[
                {
                    "role": "user",
                    "content": f"Here is the hotel contract PDF text. Generate the complete rate sheet JSON:\n\n{contract_text}"
                }
            ]
        )

        contract_data = json.loads(contract_response.content[0].text)
        contract_rows = []
        for row in contract_data["rows"]:
            contract_rows.append([row.get(h, "") for h in CONTRACT_HEADERS])

        contract_excel = generate_excel_from_data(contract_rows, CONTRACT_HEADERS)

        result = {"contract_excel": contract_excel}

        if promotion_file:
            promotion_text = extract_text_from_pdf(promotion_file)

            promotion_response = client.messages.create(
                model="claude-sonnet-4-6",
                max_tokens=8000,
                system=PROMOTION_SYSTEM_PROMPT,
                messages=[
                    {
                        "role": "user",
                        "content": f"Here is the hotel promotion PDF text. The contract context is also provided for supplement rules.\n\nPROMOTION PDF:\n{promotion_text}\n\nCONTRACT CONTEXT:\n{contract_text}"
                    }
                ]
            )

            promotion_data = json.loads(promotion_response.content[0].text)
            promotion_rows = []
            for row in promotion_data["rows"]:
                promotion_rows.append([row.get(h, "") for h in PROMOTION_HEADERS])

            promotion_excel = generate_excel_from_data(promotion_rows, PROMOTION_HEADERS)
            result["promotion_excel"] = promotion_excel

        return jsonify(result)

    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})

if __name__ == "__main__":
    app.run(debug=True)
