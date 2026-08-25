# -*- coding: utf-8 -*-
"""
Metadados de diagnóstico gravados em todos os PDFs gerados pelo ANEXT —
existem só para investigar erros/dúvidas futuras sobre como um PDF foi
montado (versão do app, quando, de onde), não como registro de auditoria
ou fiscalização de uso.
Visíveis em Arquivo → Propriedades no Acrobat Reader (aba Descrição):
  Produtor do PDF  → ANEXT vX.Y
  Aplicativo       → data_hora
Não aparecem no conteúdo do documento nem são lidos por OCR. Deliberadamente
não inclui usuário, máquina, pasta/arquivos de origem nem nomes de
segurados — informação de cliente ou de quem operou o programa não tem
por que ir pras propriedades do PDF.
"""

from datetime import datetime
from pathlib import Path


def _ler_versao() -> str:
    try:
        versao_txt = Path(__file__).parent.parent / "versao.txt"
        return versao_txt.read_text(encoding="utf-8").strip()
    except Exception:
        return "?"


def metadados_anext() -> dict:
    """Retorna dict de metadados para passar a fitz.Document.set_metadata()
    antes de salvar qualquer PDF gerado pelo ANEXT."""
    versao = _ler_versao()
    data_hora = datetime.now().strftime("%Y-%m-%d %H:%M")

    return {
        "producer": f"ANEXT v{versao}",
        "creator": data_hora,
    }
