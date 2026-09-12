from dataclasses import dataclass
from typing import Optional, List, Dict
from datetime import datetime
from enum import Enum
import json
from pathlib import Path

@dataclass
class BankTransaction:
    date: str
    description: str
    amount: float
    balance: float
    currency: str = "TRY"
    transaction_type: str = "unknown"
    reference: str = ""

class DynamicBankParser:
    """Dinamik banka parser - JSON config'den yüklenir"""
    
    def __init__(self, config: Dict):
        self.config = config
        self.id = config['id']
        self.name = config['name']
        self.detect_keywords = config.get('detect_keywords', [])
        self.separator = config.get('separator', '\t')
        self.columns = config.get('columns', [])
    
    def detect_format(self, content: str) -> bool:
        """İçerikte banka anahtar kelimeleri var mı?"""
        content_upper = content.upper()
        return any(keyword in content_upper for keyword in self.detect_keywords)
    
    def parse(self, content: str) -> List[BankTransaction]:
        """Yapılandırılmış metin dosyasını ayrıştır"""
        transactions = []
        lines = content.split('\n')
        
        for line in lines:
            if self.separator in line and any(c.isdigit() for c in line.split(self.separator)[0]):
                parts = [p.strip() for p in line.split(self.separator)]
                
                if len(parts) >= len(self.columns):
                    try:
                        record = {col: parts[i] for i, col in enumerate(self.columns)}
                        
                        date = record.get('date', '')
                        desc = record.get('desc', '')
                        amount = self.parse_turkish_amount(record.get('amount', '0'))
                        balance = self.parse_turkish_amount(record.get('balance', '0'))
                        
                        trans_type = "gelen" if amount > 0 else "giden"
                        
                        transactions.append(BankTransaction(
                            date=date,
                            description=desc,
                            amount=abs(amount),
                            balance=balance,
                            transaction_type=trans_type
                        ))
                    except (ValueError, IndexError, KeyError):
                        continue
        
        return transactions
    
    @staticmethod
    def parse_turkish_amount(amount_str: str) -> float:
        """Türkçe format (1.000,50) → float"""
        if not amount_str:
            return 0.0
        amount_str = amount_str.strip()
        amount_str = amount_str.replace('.', '')
        amount_str = amount_str.replace(',', '.')
        try:
            return float(amount_str)
        except ValueError:
            return 0.0

class BankParserRegistry:
    """Dinamik banka registry - JSON config'den yüklenir"""
    
    _instance = None
    _parsers: Dict[str, DynamicBankParser] = {}
    _config_path = Path('/srv/data/banks_config.json')
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            cls._instance._load_config()
        return cls._instance
    
    def _load_config(self):
        """JSON config dosyasından bankaları yükle"""
        if self._config_path.exists():
            with open(self._config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
                for bank_config in config.get('banks', []):
                    parser = DynamicBankParser(bank_config)
                    self._parsers[bank_config['id']] = parser
    
    def parse_bank_statement(self, content: str) -> tuple:
        """Otomatik banka tespit ve ayrıştırma"""
        for bank_id, parser in self._parsers.items():
            if parser.detect_format(content):
                transactions = parser.parse(content)
                return transactions, bank_id
        return [], None
    
    def get_supported_banks(self) -> List[str]:
        """Desteklenen bankaları listele"""
        return [parser.name for parser in self._parsers.values()]
    
    def add_bank(self, config: Dict):
        """Yeni banka ekle (runtime)"""
        parser = DynamicBankParser(config)
        self._parsers[config['id']] = parser
    
    def reload_config(self):
        """Config'i yeniden yükle"""
        self._parsers.clear()
        self._load_config()

def parse_bank_statement(content: str) -> tuple:
    """Banka dosyasını ayrıştır"""
    registry = BankParserRegistry()
    return registry.parse_bank_statement(content)

def get_supported_banks() -> List[str]:
    """Desteklenen bankaları getir"""
    registry = BankParserRegistry()
    return registry.get_supported_banks()
