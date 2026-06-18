import os
import json
import io
import base64
import zipfile
from typing import Optional
from fastapi import FastAPI, File, UploadFile, Form, HTTPException
from fastapi.responses import StreamingResponse
import anthropic
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from pydantic import BaseModel
import uvicorn

# -------------------- Configuration --------------------
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY")
if not ANTHROPIC_API_KEY:
    raise ValueError("ANTHROPIC_API_KEY environment variable not set")

client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

app = FastAPI(title="Hotel Rate Sheet Processor")

from fastapi.middleware.cors import CORSMiddleware

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # For development; restrict to your Lovable domain in production
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# -------------------- Pydantic Schemas --------------------
class RateRow(BaseModel):
    hotel: str
    room: str
    accommodation: str
    meal: str
    season_begin: str
    season_end: str
    reservation_date_from: str
    reservation_date_till: str
    nights: int
    hotel_net_price: float
    number_of_markups: int
    currency_code: str
    currency: str
    season_type: str
    market_code: str
    price_type: str
    staying_nights_from: int
    staying_nights_till: int
    booking_code: Optional[str] = ""

class SPORow(BaseModel):
    spo_no: str
    price_type: str
    hotel: str
    room: str
    accommodation: str
    meal: str
    hotel_net_price: float
    currency_code: str
    market_code: str
    season_begin: str
    season_end: str
    days_before_checkin_from: Optional[int] = None
    reservation_date_from: str
    reservation_date_till: str
    check_in_from: str
    check_in_till: str
    check_out_from: Optional[str] = None
    staying_nights_from: int
    check_out_till: Optional[str] = None
    nights: int
    nights_from: Optional[int] = None
    nights_till: Optional[int] = None
    number_of_markups: int
    nights_free: Optional[int] = None
    season_type: str
    days_before_checkin_till: Optional[int] = None
    staying_nights_till: int
    booking_code: str

class OutputSchema(BaseModel):
    contract_rates: list[RateRow]
    promotion_rates: list[SPORow]

# -------------------- System Prompt --------------------
SYSTEM_PROMPT = """You are a Hotel Rate Standardisation Expert. You will receive two PDFs: a signed contract and an optional promotional rate sheet.

Your task is to extract all rates, rules, and supplements, then expand them into the structured JSON schemas provided.

**Rules:**
- Use the contract PDF as the source for base rates per room per season, extra bed/3rd adult charges, child policies, and meal supplements (HB, FB, AI) per season.
- If a promotion PDF is provided, extract its discount percentages, applicable stay/booking windows, promotional meal supplements (which override contract meal supplements for the promotion period), and promo code.
- For each valid occupancy combination (based on max occupancy rules), generate a rate row for each meal type (BB, HB, FB, AI where applicable).
- For the contract: use contract meal supplements and no discounts.
- For the promotion: apply discount only to the accommodation portion (base rate + extra bed charges). Meal supplements are NOT discounted and use promotion-specific supplement rates.
- Expand all possible occupancy combinations: 1ADL, 2ADL, 2ADL+1-ADULT EXTRA BED, 2ADL+1-CHILD SHARING (00-05.99), 2ADL+1-CHILD SHARING (06-11.99), 2ADL+1-CHILD EXTRA BED (00-05.99), 2ADL+1-CHILD EXTRA BED (06-11.99), 2ADL+1-CHILD SHARING (00-05.99)+1-CHILD SHARING (06-11.99), 2ADL+2-CHILD SHARING (00-05.99), 2ADL+2-CHILD SHARING (06-11.99), 1ADL+1-CHILD SHARING (00-05.99), etc. Include all valid combos per contract's max occupancy.
- For each row, output numeric fields as floats, dates as strings in YYYY-MM-DD format.

**Output format:** You MUST output a valid JSON object that matches the provided Pydantic schemas exactly.

**Important:** Do not invent rates. If a value is not stated, set it to null or 0. Include a "warnings" array in the JSON if any ambiguities exist.

Now process the uploaded files."""

# -------------------- Helper: Excel Generation --------------------
def create_contract_excel(rows: list[RateRow]) -> io.BytesIO:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Contract Rates"

    headers = ["Hotel", "Room", "Accommodation", "Meal", "Season begin", "Season end",
               "Reservation date from", "Reservation date till", "Nights", "Hotel net price",
               "Number of markups", "Currency (code)", "Currency", "Season type",
               "Market code", "Price type", "Staying nights from", "Staying nights till",
               "Booking code"]
    ws.append(headers)

    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
        cell.alignment = Alignment(horizontal="center")

    for r in rows:
        ws.append([
            r.hotel, r.room, r.accommodation, r.meal,
            r.season_begin, r.season_end,
            r.reservation_date_from, r.reservation_date_till,
            r.nights, r.hotel_net_price,
            r.number_of_markups, r.currency_code, r.currency,
            r.season_type, r.market_code, r.price_type,
            r.staying_nights_from, r.staying_nights_till,
            r.booking_code or ""
        ])

    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 30)
        ws.column_dimensions[column].width = adjusted_width

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out

def create_promotion_excel(rows: list[SPORow]) -> io.BytesIO:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Promotion Rates"

    headers = ["SPO No", "Price type", "Hotel", "Room", "Accommodation", "Meal",
               "Hotel net price", "Currency (code)", "Market code", "Season begin",
               "Season end", "Days before check-in from", "Reservation date from",
               "Reservation date till", "Check-in from", "Check-in till",
               "Check-out from", "Staying nights from", "Check-out till",
               "Nights", "Nights from", "Nights till", "Number of markups",
               "Nights free", "Season type", "Days before check-in till",
               "Staying nights till", "Booking code"]
    ws.append(headers)

    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
        cell.alignment = Alignment(horizontal="center")

    for r in rows:
        ws.append([
            r.spo_no, r.price_type, r.hotel, r.room, r.accommodation, r.meal,
            r.hotel_net_price, r.currency_code, r.market_code,
            r.season_begin, r.season_end,
            r.days_before_checkin_from,
            r.reservation_date_from, r.reservation_date_till,
            r.check_in_from, r.check_in_till,
            r.check_out_from, r.staying_nights_from,
            r.check_out_till, r.nights,
            r.nights_from, r.nights_till,
            r.number_of_markups, r.nights_free,
            r.season_type, r.days_before_checkin_till,
            r.staying_nights_till, r.booking_code
        ])

    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 30)
        ws.column_dimensions[column].width = adjusted_width

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out

# -------------------- Endpoint --------------------
@app.post("/process")
async def process_rates(
    contract: UploadFile = File(...),
    promotion: Optional[UploadFile] = File(None)
):
    if not contract.filename.endswith('.pdf'):
        raise HTTPException(status_code=400, detail="Contract file must be PDF")
    if promotion and not promotion.filename.endswith('.pdf'):
        raise HTTPException(status_code=400, detail="Promotion file must be PDF")

    # Read PDFs and encode as base64
    contract_bytes = await contract.read()
    contract_b64 = base64.b64encode(contract_bytes).decode("utf-8")

    messages = [
        {
            "role": "user",
            "content": [
                {
                    "type": "document",
                    "source": {
                        "type": "base64",
                        "media_type": "application/pdf",
                        "data": contract_b64
                    }
                }
            ]
        }
    ]

    if promotion:
        promo_bytes = await promotion.read()
        promo_b64 = base64.b64encode(promo_bytes).decode("utf-8")
        messages[0]["content"].append({
            "type": "document",
            "source": {
                "type": "base64",
                "media_type": "application/pdf",
                "data": promo_b64
            }
        })

    messages[0]["content"].append({
        "type": "text",
        "text": "Extract all rates, supplements, and promotions. Output the JSON exactly matching the provided schemas. Include both contract and promotion rates in the response."
    })

    tool_schema = {
        "name": "output_rate_sheet",
        "description": "Output the extracted and expanded rate sheets",
        "input_schema": OutputSchema.model_json_schema()
    }

    try:
        response = client.messages.create(
            model="claude-3-5-sonnet-20241022",
            max_tokens=8192,
            system=SYSTEM_PROMPT,
            messages=messages,
            tools=[tool_schema],
            tool_choice={"type": "tool", "name": "output_rate_sheet"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Claude API error: {str(e)}")

    try:
        tool_output = response.content[0]
        if tool_output.type != "tool_use":
            raise ValueError("Claude did not use the tool")
        raw_json = tool_output.input
        output = OutputSchema(**raw_json)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to parse Claude output: {str(e)}")

    contract_excel = create_contract_excel(output.contract_rates)
    promo_excel = create_promotion_excel(output.promotion_rates) if output.promotion_rates else None

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        zip_file.writestr("Contract_Rates.xlsx", contract_excel.getvalue())
        if promo_excel:
            zip_file.writestr("Promotion_Rates.xlsx", promo_excel.getvalue())

    zip_buffer.seek(0)

    return StreamingResponse(
        zip_buffer,
        media_type="application/zip",
        headers={"Content-Disposition": "attachment; filename=rate_sheets.zip"}
    )

@app.get("/health")
async def health():
    return {"status": "ok"}

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8080)
