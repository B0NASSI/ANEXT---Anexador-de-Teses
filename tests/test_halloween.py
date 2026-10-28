import halloween


def test_estrelas_ficam_dentro_da_area_e_sao_deterministicas():
    a = halloween.estrelas(232, 600)
    assert a == halloween.estrelas(232, 600)
    assert len(a) > 20
    for x, y, raio, opacidade in a:
        assert 0 < x < 232 and 0 < y < 600
        assert raio > 0 and 0 <= opacidade <= 255


def test_misturar_extremos_e_meio():
    assert halloween.misturar("#2B1646", (255, 255, 255), 0) == "#2b1646"
    assert halloween.misturar("#2B1646", (255, 255, 255), 255) == "#ffffff"
    assert halloween.misturar("#000000", (200, 100, 50), 128) == "#643219"


def test_logo_com_gosma_mantem_o_logo_intacto_e_ganha_espaco_pras_gotas():
    from PIL import Image, ImageDraw
    logo = Image.new("RGBA", (120, 40), (0, 0, 0, 0))
    ImageDraw.Draw(logo).rectangle([10, 5, 110, 38], fill=(255, 255, 255, 255))
    img = halloween.logo_com_gosma(logo, escorrimento=24)
    assert img.size == (120, 64)
    # o miolo das letras não é coberto pela gosma
    assert img.getpixel((60, 20)) == (255, 255, 255, 255)
    # e tem gosma pingando abaixo do logo
    assert any(img.getpixel((x, y))[3] > 0 for x in range(120) for y in range(42, 64))


def test_silhueta_morcego_simetrica():
    for erguidas in (False, True):
        pontos = halloween.silhueta_morcego(1.0, asas_erguidas=erguidas)
        xs = sorted(round(x, 6) for x, _ in pontos)
        assert xs == sorted(-x for x in xs)


def test_icones_e_fantasma_tem_o_tamanho_pedido_e_fundo_transparente():
    for img, altura in ((halloween.icone_bruxa_vassoura(20), 20), (halloween.icone_caldeirao(20), 20)):
        assert img.height == altura and img.mode == "RGBA"
        assert img.getpixel((0, 0))[3] == 0
    fantasma = halloween.fantasma_fap(34)
    assert fantasma.width == 34 + 12 and fantasma.height > fantasma.width  # mais alto que largo
    assert fantasma.getpixel((0, 0))[3] == 0
    # corpo sólido no centro, sem semitransparência
    assert fantasma.getpixel((fantasma.width // 2, fantasma.height // 2))[3] == 255


def test_piadas_do_fantasma():
    assert len(halloween.PIADAS_FAP) >= 8
    assert len(set(halloween.PIADAS_FAP)) == len(halloween.PIADAS_FAP)
    for piada in halloween.PIADAS_FAP:
        assert piada.startswith("Buuu!") and len(piada) <= 75


def test_icones_dos_botoes_de_acao():
    for desenho in (halloween.icone_tumulo, halloween.icone_machado, halloween.icone_pocao):
        icone = desenho(20)
        assert icone.height == 20 and icone.mode == "RGBA"
        assert icone.getpixel((0, 0))[3] == 0  # canto transparente


def test_frases_de_sucesso_e_erro():
    assert len(halloween.MENSAGENS_SUCESSO) >= 3 and len(halloween.PIADAS_ERRO) >= 3
    assert all(p.startswith("Buuu!") for p in halloween.PIADAS_ERRO)
