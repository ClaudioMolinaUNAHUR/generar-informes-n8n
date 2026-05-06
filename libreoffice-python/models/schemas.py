from pydantic import BaseModel
from typing import Optional, List, Any

class GenerateRequest(BaseModel):
    # Definición mínima para permitir la importación en app.py
    data: dict