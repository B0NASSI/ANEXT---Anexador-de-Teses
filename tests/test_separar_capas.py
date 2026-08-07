# -*- coding: utf-8 -*-
import fitz

from helpers import criar_pdf_paginas
from separar_capas import calcular_destinos, separar_capas

# "Capa Geral" (pág. 0) + "Benefícios Gerais" (pág. 1) + 1 capa por segurado
# (págs. 2+) — mesma estrutura descrita no docstring de separar_capas.py,
# com nomes fictícios no lugar dos de cliente real.
PAGINAS_TESE = [
    ["TÓPICO 4.1", "Capa Geral da tese"],
    ["ITEM", "BENEFÍCIO", "DATA"],  # "tabela completa" - não é capa individual
    ["FULANO DE TAL", "ITEM 1 - Benefício B91"],
    ["CICLANO DA SILVA", "ITEM 1 - Benefício B91"],
]


def test_calcular_destinos_nao_escreve_nada_em_disco(tmp_path):
    entrada = tmp_path / "entrada.pdf"
    criar_pdf_paginas(entrada, PAGINAS_TESE)
    pasta_saida = tmp_path / "saida"

    destinos = calcular_destinos(entrada, pasta_saida)

    assert [d.name for d in destinos] == [
        "0. Tópico 4.1 - FULANO DE TAL.pdf",
        "0.1 Tópico 4.1 - CICLANO DA SILVA.pdf",
    ]
    assert not pasta_saida.exists()


def test_separar_capas_primeiro_segurado_leva_capa_geral_e_beneficios(tmp_path):
    entrada = tmp_path / "entrada.pdf"
    criar_pdf_paginas(entrada, PAGINAS_TESE)
    pasta_saida = tmp_path / "saida"

    gerados = separar_capas(entrada, pasta_saida)

    assert len(gerados) == 2
    with fitz.open(gerados[0]) as doc:
        assert doc.page_count == 3  # capa geral + benefícios gerais + capa dele
    with fitz.open(gerados[1]) as doc:
        assert doc.page_count == 1  # só a própria capa


def test_separar_capas_usa_subpastas_numeradas_quando_ja_existem(tmp_path):
    entrada = tmp_path / "entrada.pdf"
    criar_pdf_paginas(entrada, PAGINAS_TESE)
    pasta_saida = tmp_path / "saida"
    (pasta_saida / "1. Qualquer nome").mkdir(parents=True)
    (pasta_saida / "2. Outro nome").mkdir(parents=True)

    gerados = separar_capas(entrada, pasta_saida)

    assert gerados[0].parent.name == "1. Qualquer nome"
    assert gerados[1].parent.name == "2. Outro nome"


def test_separar_capas_nao_quebra_com_nome_de_segurado_extremamente_longo(tmp_path):
    # nome de segurado gigante (ex.: várias partes/sobrenomes) não deve
    # travar o salvamento por estourar o limite de caminho do Windows
    nome_gigante = "FULANO " + "DE TAL " * 40  # bem além de qualquer limite razoável
    paginas = [
        ["TÓPICO 4.1", "Capa Geral da tese"],
        ["ITEM", "BENEFÍCIO", "DATA"],
        [nome_gigante, "ITEM 1 - Benefício B91"],
    ]
    entrada = tmp_path / "entrada.pdf"
    criar_pdf_paginas(entrada, paginas)
    pasta_saida = tmp_path / "saida"

    gerados = separar_capas(entrada, pasta_saida)

    assert len(gerados) == 1
    assert gerados[0].is_file()
    assert len(str(gerados[0])) <= 259
