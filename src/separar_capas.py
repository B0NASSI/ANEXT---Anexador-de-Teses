"""
Divide um PDF de "capas" (Capa Geral + Benefícios Gerais + 1 capa por segurado)
em arquivos separados, um por segurado.

Regra:
- Página 1: Capa Geral
- Página 2 em diante: Benefícios Gerais (pode ocupar mais de 1 página)
- Depois disso: uma capa por segurado, sempre 1 página cada

- O 1º segurado recebe Capa Geral + Benefícios Gerais (todas as páginas) + a capa dele.
- Os demais segurados recebem apenas a própria capa.

A transição entre "Benefícios Gerais" e as capas individuais é detectada pelo texto:
a capa individual tem o nome do segurado isolado, direto acima da tabela "Item".

Se a pasta de saída já tiver uma subpasta numerada para cada segurado (ex: "1. NOME",
"2. NOME"...), cada capa é salva direto na subpasta correspondente, casando por ordem
(não por nome, já que o nome da pasta pode ter pequenas diferenças do nome extraído do PDF).
Senão, salva tudo direto na pasta de saída com prefixo sequencial para manter a ordem.
"""

import re
import sys
import time
from pathlib import Path

import fitz  # PyMuPDF

from limites_caminho import truncar_para_caminho

CARACTERES_INVALIDOS = r'[<>:"/\\|?*]'


def _com_repeticao(acao, tentativas: int = 4, espera: float = 0.5):
    # PDFs recem-criados podem ficar momentaneamente bloqueados por
    # antivirus/indexacao/sincronizacao (OneDrive) - tentar de novo depois
    # de uma pequena espera resolve a maioria desses casos transitorios.
    for tentativa in range(tentativas):
        try:
            return acao()
        except Exception:
            if tentativa == tentativas - 1:
                raise
            time.sleep(espera)


def nome_arquivo_seguro(texto: str) -> str:
    # a "/" some sem deixar rastro com um simples strip - problema pra CNPJ
    # (ex.: "85.782.878/0008-55" viraria "85.782.8780008-55", ilegível);
    # troca por "-" antes de remover o resto dos caracteres proibidos
    texto = texto.replace("/", "-")
    texto = re.sub(CARACTERES_INVALIDOS, "", texto).strip()
    return texto or "SEM NOME"


def linhas_relevantes(pagina: fitz.Page) -> list[str]:
    return [l.strip() for l in pagina.get_text().splitlines() if l.strip()]


def extrair_nome_segurado(pagina: fitz.Page) -> str:
    linhas = linhas_relevantes(pagina)
    return linhas[0] if linhas else "SEM NOME"


def extrair_numero_topico(pagina_capa_geral: fitz.Page) -> str:
    match = re.search(r"T[ÓO]PICO\s+([0-9]+(?:\.[0-9]+)?)", pagina_capa_geral.get_text(), re.IGNORECASE)
    return match.group(1) if match else "?"


def pastas_segurados_em_ordem(pasta_saida: Path) -> list[Path]:
    """Subpastas numeradas no início do nome (ex: "1. NOME"), ordenadas pelo número."""
    if not pasta_saida.is_dir():
        return []
    candidatas = []
    for item in pasta_saida.iterdir():
        if not item.is_dir():
            continue
        match = re.match(r"\s*(\d+)\s*\.", item.name)
        if match:
            candidatas.append((int(match.group(1)), item))
    candidatas.sort(key=lambda par: par[0])
    return [pasta for _, pasta in candidatas]


def eh_capa_individual(pagina: fitz.Page, titulo_beneficio: str) -> bool:
    """Capa individual: nome do segurado isolado, direto acima da tabela "Item".
    Páginas da planilha geral (e continuações) começam com "Item" ou repetem o título do benefício."""
    linhas = linhas_relevantes(pagina)
    if len(linhas) < 2:
        return False
    primeira = linhas[0].upper()
    if primeira.startswith("ITEM"):
        return False
    if titulo_beneficio and titulo_beneficio.upper() in primeira:
        return False
    return linhas[1].upper().startswith("ITEM")


def _montar_plano(doc: fitz.Document, pasta_saida: Path) -> tuple[list[tuple[Path, int | None]], int | None]:
    """Calcula, para cada segurado, o caminho de saída e o índice da página da capa.
    Retorna (plano, inicio_individuais) — não escreve nada em disco."""
    if doc.page_count < 2:
        raise ValueError("O PDF precisa ter ao menos 2 páginas (capa geral e a tabela de segurados).")

    numero_topico = extrair_numero_topico(doc[0])
    linhas_pagina_tabela = linhas_relevantes(doc[1])
    if not linhas_pagina_tabela:
        # a 2a página (índice 1) devia ser sempre a "tabela completa" — se
        # vier em branco (ex.: página em branco intrusa antes dela), o
        # título fica vazio e a detecção de onde começam as capas
        # individuais abaixo erra silenciosamente, embaralhando segurados.
        # Falha alto e claro em vez de gerar uma divisão errada.
        raise ValueError(
            "A segunda página do PDF de capas está em branco ou sem texto "
            "legível — não foi possível identificar a tabela completa para "
            "dividir as capas corretamente."
        )
    titulo_beneficio = linhas_pagina_tabela[0]

    inicio_individuais = next(
        (idx for idx in range(1, doc.page_count) if eh_capa_individual(doc[idx], titulo_beneficio)),
        None,
    )
    if inicio_individuais is None:
        # Só há 1 segurado na tabela: não existe página individual separada
        # (foi propositalmente omitida por repetir a tabela completa) - o
        # documento inteiro (capa geral + tabela) é a própria capa dele.
        pastas = pastas_segurados_em_ordem(pasta_saida)
        if len(pastas) == 1:
            destino = pastas[0]
            nome = re.sub(r"^\s*\d+\s*\.\s*", "", destino.name)  # nome do segurado, tirando o prefixo numérico da subpasta
        else:
            destino = pasta_saida
            nome = titulo_beneficio
        prefixo_fixo = f"0. Tópico {numero_topico} - "
        nome_seguro = truncar_para_caminho(destino, nome_arquivo_seguro(nome), len(prefixo_fixo) + len(".pdf"))
        nome_arquivo = f"{prefixo_fixo}{nome_seguro}.pdf"
        return [(destino / nome_arquivo, None)], None

    total_segurados = doc.page_count - inicio_individuais
    pastas = pastas_segurados_em_ordem(pasta_saida)
    usar_pastas = len(pastas) == total_segurados

    plano = []
    for posicao, pagina_idx in enumerate(range(inicio_individuais, doc.page_count), start=1):
        nome = extrair_nome_segurado(doc[pagina_idx])

        if usar_pastas:
            destino = pastas[posicao - 1]
            prefixo_fixo = f"0. Tópico {numero_topico} - "
        else:
            prefixo = "0." if posicao == 1 else f"0.{posicao - 1}"
            destino = pasta_saida
            prefixo_fixo = f"{prefixo} Tópico {numero_topico} - "

        # nome do segurado pode ser bem longo — trunca o quanto for preciso
        # pra "destino / nome_arquivo" não estourar o limite de caminho do
        # Windows (260 caracteres), que senão faz o salvamento falhar com
        # um erro de sistema incompreensível pra quem não é técnico
        nome_seguro = truncar_para_caminho(destino, nome_arquivo_seguro(nome), len(prefixo_fixo) + len(".pdf"))
        nome_arquivo = f"{prefixo_fixo}{nome_seguro}.pdf"

        plano.append((destino / nome_arquivo, pagina_idx))

    return plano, inicio_individuais


def calcular_destinos(caminho_entrada: Path, pasta_saida: Path) -> list[Path]:
    """Calcula os arquivos que seriam gerados, sem escrever nada — usado para avisar
    sobre sobrescrição antes de processar de fato."""
    with _com_repeticao(lambda: fitz.open(caminho_entrada)) as doc:
        plano, _ = _montar_plano(doc, pasta_saida)
    return [caminho for caminho, _ in plano]


def separar_capas(caminho_entrada: Path, pasta_saida: Path) -> list[Path]:
    pasta_saida.mkdir(parents=True, exist_ok=True)

    with _com_repeticao(lambda: fitz.open(caminho_entrada)) as doc:
        plano, inicio_individuais = _montar_plano(doc, pasta_saida)

        arquivos_gerados = []
        for posicao, (caminho_saida, pagina_idx) in enumerate(plano, start=1):
            novo = fitz.open()
            try:
                if pagina_idx is None:
                    novo.insert_pdf(doc)  # único segurado: documento inteiro é a capa dele
                else:
                    if posicao == 1:
                        novo.insert_pdf(doc, from_page=0, to_page=inicio_individuais - 1)  # capa geral + benefícios gerais (todas as páginas)
                    novo.insert_pdf(doc, from_page=pagina_idx, to_page=pagina_idx)
                _com_repeticao(lambda: novo.save(caminho_saida))
            finally:
                novo.close()
            arquivos_gerados.append(caminho_saida)

    return arquivos_gerados


if __name__ == "__main__":
    raiz_projeto = Path(__file__).resolve().parent.parent
    entrada = Path(sys.argv[1]) if len(sys.argv) > 1 else raiz_projeto / "modelos" / "0. Topico 4 - Trajeto - Capa.pdf"
    saida = Path(sys.argv[2]) if len(sys.argv) > 2 else raiz_projeto / "saida"

    arquivos = separar_capas(entrada, saida)
    print(f"{len(arquivos)} arquivo(s) gerado(s) em: {saida}")
    for arquivo in arquivos:
        print(f" - {arquivo.name}")
