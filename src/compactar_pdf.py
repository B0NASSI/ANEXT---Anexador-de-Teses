"""
Lógica de compressão de PDFs.
Sem código de interface — só funções puras chamadas por app.py.
"""

import os


# ── Compressão sem perda ──────────────────────────────────────────────────────

def compress_lossless(src: str, dst: str) -> int:
    import fitz

    doc = fitz.open(src)
    try:
        doc.save(dst, garbage=4, deflate=True)
    finally:
        doc.close()
    return os.path.getsize(dst)


# ── Recompressão de imagens (PyMuPDF) ─────────────────────────────────────────

DPI_PADRAO = 120


def _recomprimir(src: str, dst: str, quality: int, dpi: int) -> int:
    import fitz

    doc = fitz.open(src)
    try:
        # threshold logo acima do alvo: qualquer imagem acima do alvo é
        # reduzida. Threshold alto demais faz scans de ~150 dpi serem
        # pulados e a compressão cai para quase nada.
        doc.rewrite_images(dpi_threshold=int(dpi * 1.1) + 1, dpi_target=dpi,
                           quality=quality)
        doc.save(dst, garbage=4, deflate=True)
    finally:
        doc.close()
    return os.path.getsize(dst)


def compress_images(src: str, dst: str, quality: int) -> int:
    return _recomprimir(src, dst, quality, DPI_PADRAO)


# ── Compactar até um tamanho alvo ─────────────────────────────────────────────

NIVEIS_ALVO = [
    ("qualidade alta — 200 dpi", 80, 200),
    ("qualidade boa — 150 dpi", 70, 150),
    ("equilibrado — 120 dpi", 60, 120),
    ("compacto — 96 dpi", 50, 96),
    ("mais compacto — 72 dpi", 40, 72),
    ("máxima compressão — 54 dpi", 25, 54),
]


def compress_to_target(src: str, dst: str, max_bytes: int,
                       cancel_event=None, status_cb=None) -> tuple[int, str, bool]:
    """
    Tenta níveis decrescentes de qualidade/resolução até o resultado
    caber em max_bytes. Retorna (tamanho final, nível usado, coube?).
    O arquivo de saída fica com o resultado do último nível tentado.
    """
    size, nome = 0, ""
    for i, (nome, quality, dpi) in enumerate(NIVEIS_ALVO, start=1):
        if cancel_event and cancel_event.is_set():
            raise InterruptedError("Cancelado pelo usuário.")
        if status_cb:
            status_cb(f"Tentativa {i} de {len(NIVEIS_ALVO)}  —  {nome}...")
        size = _recomprimir(src, dst, quality, dpi)
        if size <= max_bytes:
            return size, nome, True
    return size, nome, False


# ── Utilitário ────────────────────────────────────────────────────────────────

def fmt_tamanho(size: int) -> str:
    if size < 1024:
        return f"{size} B"
    if size < 1024 ** 2:
        return f"{size / 1024:.1f} KB"
    return f"{size / 1024 ** 2:.2f} MB"
