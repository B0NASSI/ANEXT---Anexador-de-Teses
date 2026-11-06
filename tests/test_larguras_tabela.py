# -*- coding: utf-8 -*-
"""Larguras e cores das tabelas geradas. As larguras salvas pela equipe em
"Padronizar tabela" valem em TODAS as tabelas (completa e de cada segurado)
— o catálogo do TABELAS.docx não pode passar por cima delas."""
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Twips

import documento
from tabela import extrair_de_docx

RAIZ = Path(__file__).resolve().parent.parent
MODELO = RAIZ / "modelo" / "modelo.docx"
# mesmos cabeçalhos de uma tabela do TABELAS.docx (tipo catalogado), com 7
# colunas (que também têm largura padrão no programa)
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


def _proporcoes(doc) -> list[float]:
    larguras, _ = _grade(doc)
    return [round(l / sum(larguras), 2) for l in larguras]


def test_larguras_do_padronizar_ganham_do_catalogo_em_todas_as_tabelas(tmp_path):
    origem = tmp_path / "tabela.docx"
    _tabela_de_origem(origem)
    tabela, referencia = extrair_de_docx(str(origem))
    documento.carregar_catalogo_tabelas(RAIZ / "modelo" / "TABELAS.docx")
    catalogo = documento._larguras_do_catalogo(tabela.cabecalhos)
    assert catalogo, "o tipo devia estar no catálogo"
    da_equipe = [0.05, 0.10, 0.20, 0.30, 0.15, 0.05, 0.15]
    esperado = [round(p, 2) for p in da_equipe]

    completa = documento.montar_tabela_completa(
        str(MODELO), "TESE", tabela.cabecalhos, tabela.linhas, tabela.indice_nome, referencia,
        documento.LargurasDoUsuario(da_equipe),
    )
    assert _proporcoes(completa) == esperado
    for nome, linhas in tabela.grupos:
        pagina = documento.montar_pagina_segurado(
            str(MODELO), nome, linhas, tabela.cabecalhos, tabela.indice_nome, referencia,
            documento.LargurasDoUsuario(da_equipe),
        )
        assert _proporcoes(pagina) == esperado, nome

    # sem larguras salvas (só o padrão do programa), o catálogo continua valendo
    sem_padronizar = documento.montar_tabela_completa(
        str(MODELO), "TESE", tabela.cabecalhos, tabela.linhas, tabela.indice_nome, referencia,
        [1 / 7] * 7,
    )
    assert _grade(sem_padronizar)[0] == catalogo


def _fundos(doc):
    """Cor de fundo de cada célula (linha a linha) da última tabela."""
    saida = []
    for tr in doc.tables[-1]._tbl.findall(qn("w:tr")):
        linha = []
        for tc in tr.findall(qn("w:tc")):
            shd = tc.find(qn("w:tcPr")).find(qn("w:shd"))
            linha.append(None if shd is None else (shd.get(qn("w:fill")), shd.get(qn("w:themeFill"))))
        saida.append(linha)
    return saida


def test_cabecalho_e_item_padronizados_corpo_mantem_destaques_da_origem(tmp_path):
    origem = tmp_path / "tabela.docx"
    _tabela_de_origem(origem)
    # cabeçalho em azul de tema (fora do padrão) e a 2ª linha de dados inteira em
    # azul-claro de tema (destaque proposital, ex.: o benefício contestado)
    doc = Document(str(origem))
    for tr in doc.tables[0]._tbl.findall(qn("w:tr"))[0:3:2]:
        for tc in tr.findall(qn("w:tc")):
            tcPr = tc.get_or_add_tcPr()
            shd = tcPr.makeelement(qn("w:shd"), {qn("w:val"): "clear", qn("w:fill"): "BDD7EE", qn("w:themeFill"): "accent1"})
            tcPr.append(shd)
    doc.save(str(origem))

    tabela, referencia = extrair_de_docx(str(origem))
    completa = documento.montar_tabela_completa(
        str(MODELO), "TESE", tabela.cabecalhos, tabela.linhas, tabela.indice_nome, referencia, None,
    )
    destaque = (documento.COR_TABELA_DESTAQUE, None)
    fundos = _fundos(completa)
    assert fundos[0] == [destaque] * len(fundos[0])  # cabeçalho no padrão
    assert fundos[1] == [destaque] + [None] * (len(fundos[1]) - 1)  # linha sem cor: só o Item
    destacada = (documento.COR_LINHA_DESTACADA, None)
    assert fundos[2] == [destaque] + [destacada] * (len(fundos[2]) - 1)  # linha destacada
