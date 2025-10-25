"""
Pytest configuration and shared fixtures
"""

import pytest
import tempfile
from pathlib import Path


@pytest.fixture
def temp_dir():
    """Provide a temporary directory for tests"""
    with tempfile.TemporaryDirectory() as tmpdir:
        yield Path(tmpdir)


@pytest.fixture
def sample_config():
    """Provide a sample configuration dictionary"""
    return {
        "output_filename": "test_output.xlsx",
        "buildings": [
            {"name": "عمارة أ", "units": 5, "notes": ""},
            {"name": "عمارة ب", "units": 10, "notes": ""}
        ],
        "units": [
            {
                "unit_no": "101",
                "building": "عمارة أ",
                "type": "استديو",
                "rent": 2000,
                "status": "مُؤجّرة",
                "notes": ""
            }
        ],
        "tenants": [
            {
                "unit_no": "101",
                "name": "محمد أحمد",
                "id": "1234567890",
                "mobile": "0500000000",
                "start_date": "2024-01-01",
                "end_date": "2024-12-31",
                "rent": 2000,
                "email": "test@example.com",
                "notes": ""
            }
        ],
        "rents_paid": [
            {
                "unit_no": "101",
                "month": "يناير",
                "year": 2024,
                "amount": 2000,
                "date": "2024-01-05",
                "method": "تحويل بنكي",
                "status": "مدفوع",
                "notes": ""
            }
        ],
        "expenses": [
            {
                "building": "عمارة أ",
                "date": "2024-01-01",
                "type": "فاتورة كهرباء",
                "amount": 500,
                "category": "فواتير",
                "notes": ""
            }
        ]
    }
