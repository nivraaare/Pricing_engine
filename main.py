from fastapi import FastAPI, HTTPException, Header
from pydantic import BaseModel, Field
from typing import Optional
import formulas
import os

app = FastAPI(
    title="NivrāaCare Dynamic Pricing API",
    description="CPQ calculation engine powered directly by the master CFO financial model."
)

EXCEL_FILE = "NivraaCare_Pricing_Engine.xlsx"
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

    # 3. Inject inputs into the Pricing Engine tab (Keys must be UPPERCASE for formulas)
    inputs = {
        "'PRICING ENGINE'!B10": payload.service,
        "'PRICING ENGINE'!B11": payload.hours,
        "'PRICING ENGINE'!B12": transport_str,
        "'PRICING ENGINE'!B13": payload.cab_fare
    }
    
    if payload.price_override is not None:
        inputs["'PRICING ENGINE'!B14"] = payload.price_override

    try:
        # 4. Execute calculations across the full dependency tree
        solution = xl_model.calculate(inputs=inputs)

        # Helper to extract a single cell's value from the 2D array returned by formulas
        def get_val(cell_ref):
            # formulas normalizes all internal keys to UPPERCASE
            return solution[cell_ref.upper()].value[0, 0]

        # 5. Extract calculated customer quote and internal CFO metrics
        return {
            "success": True,
            "data": {
                "service": str(get_val("'Pricing Engine'!D21")),
                "duration": str(get_val("'Pricing Engine'!D22")),
                "pricing": {
                    "careServiceCharge": float(get_val("'Pricing Engine'!D23")),
                    "estimatedCabFare": float(get_val("'Pricing Engine'!D24")),
                    "transportCoordinationFee": float(get_val("'Pricing Engine'!D25")),
                    "transportTotal": float(get_val("'Pricing Engine'!D26")),
                    "totalEstimatedCharge": float(get_val("'Pricing Engine'!D27"))
                },
                "breakup": {
                    "firstTierHours": float(get_val("'Pricing Engine'!B37")),
                    "firstTierRate": float(get_val("'Pricing Engine'!B38")),
                    "extendedHours": float(get_val("'Pricing Engine'!B40")),
                    "extendedRate": float(get_val("'Pricing Engine'!B41")),
                    "serviceMultiplier": float(get_val("'Pricing Engine'!B44")),
                    "careCoordinationFee": float(get_val("'Pricing Engine'!B46"))
                },
                "cfoEconomics": {
                    "netCareRevenue": round(float(get_val("'Pricing Engine'!B50")), 2),
                    "embeddedGST": round(float(get_val("'Pricing Engine'!B49")), 2),
                    "contribution": round(float(get_val("'Pricing Engine'!B71")), 2),
                    "marginPct": round(float(get_val("'Pricing Engine'!B72")) * 100, 2)
                }
            }
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Pricing calculation failed: {str(e)}")
