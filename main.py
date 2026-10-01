from fastapi import FastAPI, HTTPException, Header
from pydantic import BaseModel, Field
from typing import Optional
import formulas
import os

app = FastAPI(
    title="NivrāaCare Dynamic Pricing API",
    description="CPQ calculation engine powered directly by the master CFO financial model."
)

EXCEL_FILE = "NivraaCare_Clean_API_Ready_Pricing_Finacial_Model_with Manual_Formula_v1.xlsx"
API_SECRET_KEY = os.getenv("API_KEY", "nc_live_secret_key_123")

# Load model into memory once on server boot
print("Compiling Excel calculation graph...")
xl_model = formulas.ExcelModel().loads(EXCEL_FILE).finish()
print("Pricing model compiled and ready.")

class QuoteRequest(BaseModel):
    service: str = Field(default="Hospital Visit", description="Exact service catalogue title")
    hours: float = Field(default=5.0, ge=1.0, le=24.0, description="Estimated service duration in hours")
    transport_required: bool = Field(default=True, description="Whether NivrāaCare coordinates transport")
    cab_fare: float = Field(default=500.0, ge=0.0, description="Estimated third-party cab fare")
    price_override: Optional[float] = Field(default=None, description="Optional founder price override")

@app.post("/api/v1/quote")
def calculate_quote(payload: QuoteRequest, x_api_key: Optional[str] = Header(None)):
    # 1. Authorization check
    if x_api_key != API_SECRET_KEY:
        raise HTTPException(status_code=401, detail="Unauthorized: Invalid or missing X-API-Key header")

    # 2. Map transport selection to the exact string expected by the model
    transport_str = "NivrāaCare Arranges Transport" if payload.transport_required else "Patient arranges transport"

    # 3. Inject inputs into the Pricing Engine tab
    inputs = {
        "'Pricing Engine'!B10": payload.service,
        "'Pricing Engine'!B11": payload.hours,
        "'Pricing Engine'!B12": transport_str,
        "'Pricing Engine'!B13": payload.cab_fare
    }
    
    if payload.price_override is not None:
        inputs["'Pricing Engine'!B14"] = payload.price_override

    try:
        # 4. Execute calculations across the full dependency tree
        xl_model.calculate(inputs=inputs)

        # 5. Extract calculated customer quote and internal CFO metrics
        return {
            "success": True,
            "data": {
                "service": str(xl_model.evaluate("'Pricing Engine'!D21")),
                "duration": str(xl_model.evaluate("'Pricing Engine'!D22")),
                "pricing": {
                    "careServiceCharge": float(xl_model.evaluate("'Pricing Engine'!D23")),
                    "estimatedCabFare": float(xl_model.evaluate("'Pricing Engine'!D24")),
                    "transportCoordinationFee": float(xl_model.evaluate("'Pricing Engine'!D25")),
                    "transportTotal": float(xl_model.evaluate("'Pricing Engine'!D26")),
                    "totalEstimatedCharge": float(xl_model.evaluate("'Pricing Engine'!D27"))
                },
                "breakup": {
                    "firstTierHours": float(xl_model.evaluate("'Pricing Engine'!B37")),
                    "firstTierRate": float(xl_model.evaluate("'Pricing Engine'!B38")),
                    "extendedHours": float(xl_model.evaluate("'Pricing Engine'!B40")),
                    "extendedRate": float(xl_model.evaluate("'Pricing Engine'!B41")),
                    "serviceMultiplier": float(xl_model.evaluate("'Pricing Engine'!B44")),
                    "careCoordinationFee": float(xl_model.evaluate("'Pricing Engine'!B46"))
                },
                "cfoEconomics": {
                    "netCareRevenue": round(float(xl_model.evaluate("'Pricing Engine'!B50")), 2),
                    "embeddedGST": round(float(xl_model.evaluate("'Pricing Engine'!B49")), 2),
                    "contribution": round(float(xl_model.evaluate("'Pricing Engine'!B71")), 2),
                    "marginPct": round(float(xl_model.evaluate("'Pricing Engine'!B72")) * 100, 2)
                }
            }
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Pricing calculation failed: {str(e)}")
