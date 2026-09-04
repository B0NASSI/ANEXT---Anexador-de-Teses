# -*- coding: utf-8 -*-
"""
Algumas teses sempre levam um mesmo documento de referência como último
anexo (ex.: "CONTESTAÇÃO ADMINISTRATIVA, EFEITO SUSPENSIVO E PRAZO
PRESCRICIONAL" sempre fecha com a mesma decisão do STJ). Este módulo
identifica essas teses pelo título (tolerando maiúsculas/acentos/espaços
e pequenos erros de digitação) e devolve o caminho do PDF fixo a anexar.

Os PDFs ficam em `anexos_fixos/`, ao lado de `modelo/` — bundled no app
(ver `.spec`), chegam pra todo mundo a cada atualização.
"""

import sys
import unicodedata
from difflib import SequenceMatcher
from pathlib import Path

# limiar de semelhança (0-1) para tolerar erro de digitação no título sem
# disparar por engano numa tese completamente diferente
_LIMIAR_SIMILARIDADE = 0.90


def _pasta_recursos() -> Path:
    return Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent))


def _normalizar(texto: str) -> str:
    sem_acento = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("ascii")
    return " ".join(sem_acento.upper().split())


# título → arquivo dentro de anexos_fixos/. Cadastre aqui novas teses que
# sempre levem o mesmo documento de referência ao final.
_ANEXOS_POR_TITULO = {
    _normalizar("CONTESTAÇÃO ADMINISTRATIVA, EFEITO SUSPENSIVO E PRAZO PRESCRICIONAL"): "ACORDAO STJ.pdf",
}


def caminho_anexo_fixo(titulo: str) -> Path | None:
    """Se `titulo` corresponder a uma tese com anexo fixo cadastrado,
    devolve o caminho do PDF a anexar como último documento. Senão, None.
    Não confere se o arquivo existe — quem for usar o caminho deve conferir."""
    if not titulo or not titulo.strip():
        return None

    normalizado = _normalizar(titulo)
    melhor_arquivo = None
    melhor_razao = 0.0
    for titulo_cadastrado, arquivo in _ANEXOS_POR_TITULO.items():
        razao = SequenceMatcher(None, normalizado, titulo_cadastrado).ratio()
        if razao > melhor_razao:
            melhor_razao, melhor_arquivo = razao, arquivo

    if melhor_razao < _LIMIAR_SIMILARIDADE:
        return None
    return _pasta_recursos() / "anexos_fixos" / melhor_arquivo
