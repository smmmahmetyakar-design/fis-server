from fastapi import APIRouter, UploadFile, File, HTTPException
from pydantic import BaseModel
from typing import List, Dict, Optional
from fastapi.responses import HTMLResponse
import sys
sys.path.insert(0, '/srv/app')

from app.core.banks import parse_bank_statement, get_supported_banks
from app.core.learning_v2 import LearningModelV2

router = APIRouter(prefix="/api", tags=["enhanced"])
learning_models: Dict[str, LearningModelV2] = {}

class SuggestionRequest(BaseModel):
    transaction_desc: str
    firm_id: str = "default"

class FeedbackRequest(BaseModel):
    transaction_desc: str
    suggested_code: str
    correct_code: str
    firm_id: str = "default"
    user_id: str = "system"

def get_or_create_model(firm_id: str) -> LearningModelV2:
    if firm_id not in learning_models:
        learning_models[firm_id] = LearningModelV2(firm_id)
    return learning_models[firm_id]

@router.get("/banks/supported")
async def get_supported_banks_endpoint():
    banks = get_supported_banks()
    return {"status": "success", "count": len(banks), "banks": banks}

@router.post("/banks/parse")
async def parse_bank_statement_endpoint(file: UploadFile = File(...)):
    try:
        content = await file.read()
        text = content.decode('utf-8', errors='ignore')
        transactions, bank_type = parse_bank_statement(text)
        return {"status": "success", "bank_detected": bank_type.value if bank_type else None, "transaction_count": len(transactions), "transactions": [{"date": t.date, "description": t.description, "amount": t.amount, "balance": t.balance, "type": t.transaction_type} for t in transactions]}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

@router.post("/learning/suggest-account")
async def suggest_account_endpoint(request: SuggestionRequest):
    try:
        model = get_or_create_model(request.firm_id)
        result = model.suggest_account(request.transaction_desc)
        if result:
            return {"status": "success", "suggestion": {"account_code": result.account_code, "confidence": result.confidence, "layer": result.layer, "explanation": result.explanation}}
        else:
            return {"status": "no_suggestion"}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

@router.post("/learning/record-feedback")
async def record_feedback_endpoint(request: FeedbackRequest):
    try:
        model = get_or_create_model(request.firm_id)
        model.record_feedback(request.transaction_desc, request.suggested_code, request.correct_code, request.user_id)
        return {"status": "success", "message": "Feedback recorded"}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

@router.get("/learning/export-mappings")
async def export_mappings_endpoint(firm_id: str = "default"):
    try:
        model = get_or_create_model(firm_id)
        mappings = model.export_mappings()
        return {"status": "success", "mappings": mappings}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

@router.get("/", response_class=HTMLResponse)
async def serve_ui():
    return """<!DOCTYPE html>
<html><head><meta charset="UTF-8"><title>Fiş Aktarım Aracı</title><style>body{font-family:sans-serif;background:#f5f5f5;padding:20px}h1{color:#333;text-align:center}.card{background:white;padding:20px;margin:10px 0;border-radius:8px;box-shadow:0 2px 8px rgba(0,0,0,0.1)}.form-group{margin-bottom:15px}label{display:block;margin-bottom:5px;font-weight:500}input,button{padding:10px;border:1px solid #ddd;border-radius:4px;width:100%;font-size:14px}button{background:#2196F3;color:white;border:none;cursor:pointer}button:hover{background:#1976D2}.result{background:#f9f9f9;padding:15px;margin-top:15px;border-left:4px solid #2196F3;font-family:monospace;font-size:12px}</style></head><body><div style="max-width:800px;margin:0 auto"><h1>🏦 Fiş Aktarım Aracı</h1><div class="card"><h2>Desteklenen Bankalar</h2><button onclick="getBanks()">Bankaları Yükle</button><div id="bankResult" class="result" style="display:none"></div></div><div class="card"><h2>Hesap Kodu Önerisi</h2><div class="form-group"><label>İşlem Açıklaması</label><input type="text" id="suggestionDesc" value="İstanbul elektrik"></div><div class="form-group"><label>Firma</label><input type="text" id="suggestionFirm" value="umh"></div><button onclick="suggestAccount()">Önerisi Getir</button><div id="suggestionResult" class="result" style="display:none"></div></div><div class="card"><h2>Geri Bildirim</h2><div class="form-group"><label>İşlem</label><input type="text" id="feedbackDesc" value="İstanbul elektrik"></div><div class="form-group"><label>Hesap Kodu</label><input type="text" id="correctCode" value="6110"></div><button onclick="recordFeedback()">Kaydet</button><div id="feedbackResult" class="result" style="display:none"></div></div></div><script>const API_BASE='/api';async function getBanks(){const r=await fetch(API_BASE+'/banks/supported');const d=await r.json();document.getElementById('bankResult').innerHTML='<strong>Bankalar:</strong><br>'+d.banks.join('<br>');document.getElementById('bankResult').style.display='block'}async function suggestAccount(){const r=await fetch(API_BASE+'/learning/suggest-account',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({transaction_desc:document.getElementById('suggestionDesc').value,firm_id:document.getElementById('suggestionFirm').value})});const d=await r.json();const el=document.getElementById('suggestionResult');el.innerHTML=d.suggestion?'Hesap: '+d.suggestion.account_code+' (Güven: '+Math.round(d.suggestion.confidence*100)+'%)':'Önerisi yok';el.style.display='block'}async function recordFeedback(){const r=await fetch(API_BASE+'/learning/record-feedback',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({transaction_desc:document.getElementById('feedbackDesc').value,suggested_code:'',correct_code:document.getElementById('correctCode').value,firm_id:'umh'})});const d=await r.json();document.getElementById('feedbackResult').innerHTML=d.message||'Başarı';document.getElementById('feedbackResult').style.display='block'}</script></body></html>"""

@router.post("/banks/reload-config")
async def reload_banks_config():
    """Banka config'ini yeniden yükle (downtime yok)"""
    try:
        from app.core.banks import BankParserRegistry
        registry = BankParserRegistry()
        registry.reload_config()
        return {
            "status": "success",
            "message": "Bank configuration reloaded",
            "banks_count": len(registry._parsers)
        }
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))
