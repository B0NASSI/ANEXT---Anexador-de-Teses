# -*- coding: utf-8 -*-
from app import _limpar_zips_temporarios_orfaos


def test_remove_zips_temporarios_orfaos(tmp_path):
    orfao1 = tmp_path / "tmp19ln8yl9.zip"
    orfao2 = tmp_path / "tmpcv_ux0yh.zip"
    orfao1.write_bytes(b"lixo")
    orfao2.write_bytes(b"lixo")

    _limpar_zips_temporarios_orfaos(tmp_path)

    assert not orfao1.exists()
    assert not orfao2.exists()


def test_nao_remove_outros_arquivos(tmp_path):
    mantido = tmp_path / "ANEXT.exe"
    mantido.write_bytes(b"nao e lixo")
    outro_zip = tmp_path / "modelo.zip"  # não começa com "tmp"
    outro_zip.write_bytes(b"nao e lixo")

    _limpar_zips_temporarios_orfaos(tmp_path)

    assert mantido.exists()
    assert outro_zip.exists()


def test_pasta_sem_zips_nao_da_erro(tmp_path):
    _limpar_zips_temporarios_orfaos(tmp_path)  # não deve levantar exceção
