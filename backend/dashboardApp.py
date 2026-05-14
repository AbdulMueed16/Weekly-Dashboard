from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
import pandas as pd

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1bCiibJjqRGOFJrwpSvDO3yBRn_ThQ47WQ5wV3qyY8O8/export?format=csv"


@app.get("/api/dashboard")
def get_dashboard_data():

    df = pd.read_csv(GOOGLE_SHEET_URL)

    # Normalize columns
    df.columns = [str(col).strip().lower() for col in df.columns]

    print(df.columns)

    data = df.to_dict(orient="records")

    summary = {
        "totalEmployees": len(df),
        "totalTickets": int(df["total new tickets"].sum()),
        "totalClosed": int(df["total closed"].sum()),
        "totalBacklog": int(df["backlog"].sum()),
        "weeklyTargets": int(df["weekly targets"].sum())
    }

    response = JSONResponse(content={
        "summary": summary,
        "employees": data
    })

    response.headers["Cache-Control"] = "no-cache"

    return response