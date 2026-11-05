# -*- coding: utf-8 -*-
"""Tabela colada com formatação: as larguras de coluna da tabela original
valem em TODAS as tabelas geradas (completa e de cada segurado) — nem o
padrão do programa nem o TABELAS.docx podem passar por cima."""
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Twips

import documento
from tabela import extrair_de_docx

RAIZ = Path(__file__).resolve().parent.parent
MODELO = RAIZ / "modelo" / "modelo.docx"
# mesmos cabeçalhos de uma tabela do TABELAS.docx (tipo catalogado) e com 7
# colunas (tem largura padrão no programa): as duas regras que antes
# sobrescreviam as larguras da tabela original
CABECALHOS = ["Item", "Vigência do FAP", "CNPJ", "Empregados", "NIT", "Tipo", "Benefício"]
LARGURAS_DA_EQUIPE = [600, 1700, 2500, 2600, 1900, 700, 2000]  # twips, fora de qualquer padrão


def _tabela_de_origem(caminho: Path) -> None:
    doc = Document()
    t = doc.add_table(rows=3, cols=len(CABECALHOS))
    linhas = [
        CABECALHOS,
        ["1", "2022", "00.000.000/0001-00", "FULANO DE TAL", "123", "B91", "6000000001"],
        ["2", "2022", "00.000.000/0001-00", "CICLANO DA SILVA", "456", "B91", "6000000002"],
    ]
    for linha, valores in zip(t.rows, linhas):
        for celula, valor, largura in zip(linha.cells, valores, LARGURAS_DA_EQUIPE):
            celula.text = valor
            celula.width = Twips(largura)
    for coluna, largura in zip(t._tbl.find(qn("w:tblGrid")).findall(qn("w:gridCol")), LARGURAS_DA_EQUIPE):
        coluna.set(qn("w:w"), str(largura))
    doc.save(str(caminho))


def _grade(doc) -> tuple[list[int], str | None]:
    tbl = doc.tables[-1]._tbl
    larguras = [int(c.get(qn("w:w"))) for c in tbl.find(qn("w:tblGrid")).findall(qn("w:gridCol"))]
    layout = tbl.tblPr.find(qn("w:tblLayout"))
    return larguras, layout.get(qn("w:type")) if layout is not None else None


def test_tabela_colada_com_formatacao_mantem_larguras_da_original(tmp_path):
    origem = tmp_path / "tabela.docx"
    _tabela_de_origem(origem)
    tabela, referencia = extrair_de_docx(str(origem))
    documento.carregar_catalogo_tabelas(RAIZ / "modelo" / "TABELAS.docx")
    assert documento._larguras_do_catalogo(tabela.cabecalhos), "o tipo devia estar no catálogo"
    padrao_do_programa = [1 / 7] * 7

    completa = documento.montar_tabela_completa(
        str(MODELO), "TESE", tabela.cabecalhos, tabela.linhas, tabela.indice_nome, referencia,
        padrao_do_programa,
    )
    assert _grade(completa) == (LARGURAS_DA_EQUIPE, "fixed")

    for nome, linhas in tabela.grupos:
        pagina = documento.montar_pagina_segurado(
            str(MODELO), nome, linhas, tabela.cabecalhos, tabela.indice_nome, referencia,
            padrao_do_programa,
        )
        assert _grade(pagina) == (LARGURAS_DA_EQUIPE, "fixed"), nome
