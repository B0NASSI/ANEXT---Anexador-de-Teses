import datetime

import pytest

import tema


@pytest.fixture(autouse=True)
def _sem_override(monkeypatch):
    monkeypatch.delenv("ANEXT_TEMA", raising=False)


@pytest.mark.parametrize("dia", [
    datetime.date(2026, 10, 28),  # lançamento
    datetime.date(2026, 10, 31),
    datetime.date(2026, 11, 1),
    datetime.date(2026, 11, 2),   # segunda, último dia
    datetime.date(2027, 10, 31),
])
def test_halloween_de_28_10_ate_02_11(dia):
    assert tema.halloween_ativo(dia)


@pytest.mark.parametrize("dia", [
    datetime.date(2026, 10, 27),
    datetime.date(2026, 11, 3),
    datetime.date(2026, 3, 31),
    datetime.date(2026, 12, 31),
])
def test_fora_do_periodo_e_tema_normal(dia):
    assert not tema.halloween_ativo(dia)


def test_variavel_de_ambiente_forca_halloween(monkeypatch):
    monkeypatch.setenv("ANEXT_TEMA", "halloween")
    assert tema.halloween_ativo(datetime.date(2026, 5, 10))


def test_variavel_de_ambiente_forca_normal_mesmo_em_31_10(monkeypatch):
    monkeypatch.setenv("ANEXT_TEMA", "normal")
    assert not tema.halloween_ativo(datetime.date(2026, 10, 31))


def test_paletas_tem_as_mesmas_chaves():
    assert tema.PALETA_NORMAL.keys() == tema.PALETA_HALLOWEEN.keys()
