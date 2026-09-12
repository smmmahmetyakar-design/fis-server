from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple
from datetime import datetime
import json
import os
from pathlib import Path

@dataclass
class SuggestionResult:
    account_code: str
    confidence: float
    layer: int
    explanation: str

@dataclass
class Feedback:
    transaction_desc: str
    suggested_code: str
    correct_code: str
    user_id: str
    timestamp: str = field(default_factory=lambda: datetime.now().isoformat())

class LearningModelV2:
    """4-layer machine learning model for account code suggestions"""
    
    SUPPLIER_TYPES = {
        "fuel": ["yakit", "benzin", "dizel", "lpg"],
        "food": ["yemek", "kahve", "resepsiyon", "kafe"],
        "transport": ["taksi", "otobüs", "kargo", "kurye"],
        "office": ["ofis", "muhasebe", "yazilim"],
        "telecom": ["telefon", "internet", "hat"],
        "electricity": ["elektrik", "enerji"],
        "water": ["su", "kanalizasyon"],
        "insurance": ["sigorta", "prim"],
        "professional": ["danismanlik", "egitim"],
        "vehicle": ["oto", "araç", "makina"]
    }
    
    DOMAIN_RULES = {
        "vergi": (3201, 0.85),
        "kdv": (3202, 0.85),
        "dis_ticaret": (3203, 0.80),
        "kira": (6140, 0.90),
        "kur_farki": (7123, 0.75),
    }
    
    def __init__(self, firm_id: str, model_dir: str = "/data/learning_models"):
        self.firm_id = firm_id
        self.model_dir = Path(model_dir)
        self.model_dir.mkdir(parents=True, exist_ok=True)
        self.exact_matches: Dict[str, str] = {}
        self.history: List[Dict] = []
        self.load_model()
    
    def load_model(self):
        model_file = self.model_dir / f"{self.firm_id}_model.json"
        if model_file.exists():
            with open(model_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
                self.exact_matches = data.get('exact_matches', {})
                self.history = data.get('history', [])
    
    def save_model(self):
        model_file = self.model_dir / f"{self.firm_id}_model.json"
        with open(model_file, 'w', encoding='utf-8') as f:
            json.dump({
                'exact_matches': self.exact_matches,
                'history': self.history[-1000:]
            }, f, ensure_ascii=False, indent=2)
    
    def learn_from_history(self, history: List[Dict]):
        for entry in history:
            desc = entry.get('description', '').lower()
            code = entry.get('account_code', '')
            if desc and code:
                self.exact_matches[desc] = code
                self.history.append({'desc': desc, 'code': code, 'timestamp': datetime.now().isoformat()})
        self.save_model()
    
    def _levenshtein_distance(self, s1: str, s2: str) -> int:
        if len(s1) < len(s2):
            return self._levenshtein_distance(s2, s1)
        if len(s2) == 0:
            return len(s1)
        previous_row = range(len(s2) + 1)
        for i, c1 in enumerate(s1):
            current_row = [i + 1]
            for j, c2 in enumerate(s2):
                insertions = previous_row[j + 1] + 1
                deletions = current_row[j] + 1
                substitutions = previous_row[j] + (c1 != c2)
                current_row.append(min(insertions, deletions, substitutions))
            previous_row = current_row
        return previous_row[-1]
    
    def suggest_account(self, transaction_desc: str) -> Optional[SuggestionResult]:
        desc_lower = transaction_desc.lower()
        
        if desc_lower in self.exact_matches:
            return SuggestionResult(
                account_code=self.exact_matches[desc_lower],
                confidence=1.0,
                layer=1,
                explanation="Exact match"
            )
        
        for rule_keyword, (code, conf) in self.DOMAIN_RULES.items():
            if rule_keyword in desc_lower:
                return SuggestionResult(
                    account_code=str(code),
                    confidence=conf,
                    layer=4,
                    explanation=f"Domain rule: {rule_keyword}"
                )
        
        return None
    
    def record_feedback(self, transaction_desc: str, suggested_code: str, correct_code: str, user_id: str = "system"):
        desc_lower = transaction_desc.lower()
        self.exact_matches[desc_lower] = correct_code
        self.history.append({
            'desc': desc_lower,
            'code': correct_code,
            'timestamp': datetime.now().isoformat(),
            'user_id': user_id
        })
        self.save_model()
    
    def export_mappings(self) -> Dict:
        return {
            'firm_id': self.firm_id,
            'exact_matches': self.exact_matches,
            'history_size': len(self.history),
            'exported_at': datetime.now().isoformat()
        }
