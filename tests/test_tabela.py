# -*- coding: utf-8 -*-
import pytest

from tabela import extrair_de_texto


def test_extrair_de_texto_agrupa_por_nome():
    texto = (
        "Nome\tBenefício\tData\n"
        "FULANO DE TAL\tB91\t01/01/2020\n"
        "\tB94\t01/02/2021\n"
        "CICLANO DA SILVA\tB91\t05/05/2019\n"
    )
    tabela = extrair_de_texto(texto)

    assert tabela.indice_nome == 0
    assert [nome for nome, _ in tabela.grupos] == ["FULANO DE TAL", "CICLANO DA SILVA"]
    assert len(tabela.grupos[0][1]) == 2  # os 2 benefícios de FULANO
    assert len(tabela.grupos[1][1]) == 1


def test_extrair_de_texto_agrupa_por_cnpj_quando_nao_ha_coluna_de_nome():
    texto = (
        "CNPJ\tVigência\tValor\n"
        "11.111.111/0001-11\t2021\t100\n"
        "11.111.111/0001-11\t2021\t200\n"
        "22.222.222/0001-22\t2022\t300\n"
    )
    tabela = extrair_de_texto(texto)

    rotulos = [rotulo for rotulo, _ in tabela.grupos]
    assert rotulos == [
        "VIGÊNCIA 2021 – CNPJ Nº 11.111.111/0001-11",
        "VIGÊNCIA 2022 – CNPJ Nº 22.222.222/0001-22",
    ]
    assert len(tabela.grupos[0][1]) == 2
    assert len(tabela.grupos[1][1]) == 1


def test_extrair_de_texto_sem_coluna_reconhecida_da_erro_claro():
    texto = "Coluna A\tColuna B\nvalor 1\tvalor 2\n"
    with pytest.raises(ValueError, match="agrupamento"):
        extrair_de_texto(texto)
