# -*- coding: utf-8 -*-
"""
Decorações da Edição Halloween (ver tema.halloween_ativo): teia de aranha com
aranha pendurada e uma cena com lua, morcegos, abóbora e lápides.

Tudo desenhado com Pillow, em RGBA, num canvas 4x maior e reduzido no final
(suaviza as bordas) — sem arquivos de imagem soltos pra empacotar.
"""

import math
import random

from PIL import Image, ImageDraw, ImageFilter

_SS = 4  # fator de super-amostragem

_MORCEGO = [  # metade direita da silhueta; a esquerda é o espelho
    (0, -8), (3, -14), (5, -8), (12, -10), (22, -15), (31, -12), (27, -5),
    (23, -6), (19, 0), (15, -2), (11, 4), (6, 2), (3, 9), (0, 12),
]


def _rgba(cor: str, alpha: int = 255) -> tuple[int, int, int, int]:
    cor = cor.lstrip("#")
    return (int(cor[0:2], 16), int(cor[2:4], 16), int(cor[4:6], 16), alpha)


def _novo(largura: int, altura: int) -> Image.Image:
    return Image.new("RGBA", (largura * _SS, altura * _SS), (0, 0, 0, 0))


def _final(img: Image.Image, largura: int, altura: int) -> Image.Image:
    return img.resize((largura, altura), Image.LANCZOS)


def _brilho(img: Image.Image, centro, raio, cor, alpha, desfoque) -> Image.Image:
    camada = Image.new("RGBA", img.size, (0, 0, 0, 0))
    cx, cy = centro[0] * _SS, centro[1] * _SS
    r = raio * _SS
    ImageDraw.Draw(camada).ellipse([cx - r, cy - r, cx + r, cy + r], fill=(*cor[:3], alpha))
    camada = camada.filter(ImageFilter.GaussianBlur(desfoque * _SS))
    return Image.alpha_composite(img, camada)


def _bezier(p0, p1, p2, passos: int = 14) -> list[tuple[float, float]]:
    pontos = []
    for i in range(passos + 1):
        t = i / passos
        x = (1 - t) ** 2 * p0[0] + 2 * (1 - t) * t * p1[0] + t ** 2 * p2[0]
        y = (1 - t) ** 2 * p0[1] + 2 * (1 - t) * t * p1[1] + t ** 2 * p2[1]
        pontos.append((x, y))
    return pontos


def _morcego(draw: ImageDraw.ImageDraw, cx: float, cy: float, escala: float, cor) -> None:
    direita = [(cx + x * escala, cy + y * escala) for x, y in _MORCEGO]
    esquerda = [(cx - x * escala, cy + y * escala) for x, y in reversed(_MORCEGO)]
    draw.polygon([(x * _SS, y * _SS) for x, y in direita + esquerda], fill=cor)


def _aranha(draw: ImageDraw.ImageDraw, x: float, y: float, cor, fio, escala: float = 1.0) -> None:
    """Aranha com o centro do corpo em (x, y), pendurada por um fio que sobe até y=0."""
    S = _SS
    draw.line([(x * S, 0), (x * S, (y - 4 * escala) * S)], fill=fio, width=max(1, int(1.1 * S)))
    for lado in (-1, 1):
        for i in range(4):
            ky = y - 5 * escala + i * 4.6 * escala
            joelho = (x + lado * (11 + (2 if i in (1, 2) else 0)) * escala, ky - 7 * escala)
            pe = (x + lado * (17 + (1 if i in (1, 2) else 0)) * escala, ky + 6 * escala)
            inicio = (x + lado * 2.5 * escala, y - 1 * escala)
            draw.line(
                [(inicio[0] * S, inicio[1] * S), (joelho[0] * S, joelho[1] * S), (pe[0] * S, pe[1] * S)],
                fill=cor, width=int(1.7 * S), joint="curve",
            )
    draw.ellipse(  # abdômen
        [(x - 7 * escala) * S, (y - 1 * escala) * S, (x + 7 * escala) * S, (y + 15 * escala) * S], fill=cor,
    )
    draw.ellipse(  # cabeça
        [(x - 4.6 * escala) * S, (y - 8 * escala) * S, (x + 4.6 * escala) * S, (y + 1.5 * escala) * S], fill=cor,
    )
    draw.ellipse(  # ampulheta vermelha
        [(x - 2 * escala) * S, (y + 5 * escala) * S, (x + 2 * escala) * S, (y + 10 * escala) * S],
        fill=(200, 30, 30, 255),
    )
    for lado in (-1, 1):  # olhinhos
        ex = x + lado * 2 * escala
        draw.ellipse(
            [(ex - 1.1 * escala) * S, (y - 5 * escala) * S, (ex + 1.1 * escala) * S, (y - 2.8 * escala) * S],
            fill=(255, 255, 255, 255),
        )


def teia_com_aranha(largura: int, altura: int, cor_fio: str, cor_aranha: str = "#1A0F2B") -> Image.Image:
    """Teia no canto superior direito (a origem é o canto) com uma aranha
    pendurada num fio. Fundo transparente."""
    S = _SS
    img = _novo(largura, altura)
    draw = ImageDraw.Draw(img)
    origem = (largura * S, 0.0)
    comprimento = min(largura, altura) * 0.98 * S
    n_raios = 6
    angulos = [math.radians(90 * i / (n_raios - 1)) for i in range(n_raios)]
    direcoes = [(-math.cos(a), math.sin(a)) for a in angulos]  # 0° = esquerda, 90° = baixo
    fio = _rgba(cor_fio, 215)
    espessura = max(1, int(1.15 * S))

    for dx, dy in direcoes:
        draw.line([origem, (origem[0] + dx * comprimento, origem[1] + dy * comprimento)], fill=fio, width=espessura)

    aneis = 6
    for k in range(1, aneis + 1):
        raio = comprimento * k / aneis
        for i in range(n_raios - 1):
            p0 = (origem[0] + direcoes[i][0] * raio, origem[1] + direcoes[i][1] * raio)
            p1 = (origem[0] + direcoes[i + 1][0] * raio, origem[1] + direcoes[i + 1][1] * raio)
            meio = ((p0[0] + p1[0]) / 2, (p0[1] + p1[1]) / 2)
            controle = (origem[0] + (meio[0] - origem[0]) * 0.80, origem[1] + (meio[1] - origem[1]) * 0.80)
            draw.line(_bezier(p0, controle, p1), fill=fio, width=espessura, joint="curve")

    _aranha(draw, largura - 40, altura * 0.72, _rgba(cor_aranha), _rgba(cor_fio, 235), escala=0.85)
    return _final(img, largura, altura)


def _camada_abobora() -> Image.Image:
    """Abóbora (corpo de 3 gomos, cabinho e rosto aceso) numa camada
    transparente de 100x100 (em unidades de desenho, já super-amostrada),
    com o centro do corpo em (50, 52)."""
    S = _SS
    camada = _novo(100, 100)
    draw = ImageDraw.Draw(camada)
    px, py = 50, 52
    laranja, sombra = _rgba("#F28A1C"), _rgba("#C24E00")
    for dx, w in ((-15, 40), (15, 40), (0, 44)):
        draw.ellipse([(px + dx - w / 2) * S, (py - 27) * S, (px + dx + w / 2) * S, (py + 27) * S], fill=laranja, outline=sombra, width=int(1.3 * S))
    draw.polygon(  # cabinho
        [((px - 3) * S, (py - 26) * S), ((px + 5) * S, (py - 26) * S), ((px + 8) * S, (py - 37) * S), ((px + 1) * S, (py - 34) * S)],
        fill=_rgba("#4F7A22"),
    )
    brilho_rosto = _rgba("#FFE27A")
    for lado in (-1, 1):  # olhos triangulares
        ex = px + lado * 12
        draw.polygon([((ex - 5) * S, (py - 3) * S), ((ex + 5) * S, (py - 3) * S), (ex * S, (py - 14) * S)], fill=brilho_rosto)
    boca = [  # sorriso de dentes
        (-20, 4), (-14, 14), (-9, 8), (-4, 17), (0, 9), (4, 17), (9, 8), (14, 14), (20, 4),
        (14, 22), (0, 25), (-14, 22),
    ]
    draw.polygon([((px + x) * S, (py + y) * S) for x, y in boca], fill=brilho_rosto)
    return camada


def icone_abobora(tamanho: int = 256) -> Image.Image:
    """Ícone quadrado da abóbora (fundo transparente) pra janela/barra de
    tarefas na Edição Halloween."""
    S = _SS
    return _camada_abobora().crop((10 * S, 8 * S, 90 * S, 88 * S)).resize((tamanho, tamanho), Image.LANCZOS)


def cena_sidebar(largura: int = 232, altura: int = 230, ceu: bool = True) -> Image.Image:
    """Cena da barra lateral: lua cheia, morcegos, abóbora e lápides num
    morro. Fundo transparente (aparece o roxo da barra por trás).
    `ceu=False` desenha só o morro (abóbora e lápides), pra janelas baixas.

    O brilho da lua é um degradê suave: se encostar na borda de cima da
    imagem ele é cortado reto e vira uma linha horizontal bem visível contra o
    céu estrelado que vem logo acima. Por isso a lua só ganha o brilho grande
    em imagens altas, onde sobra espaço em cima pra ele sumir sozinho; nas
    mais baixas o brilho é menor e a lua fica mais pra baixo."""
    S = _SS
    img = _novo(largura, altura)
    draw = ImageDraw.Draw(img)

    if ceu:
        alta = altura >= 330
        lua = (158, 112 if alta else 88)
        raio_brilho, desfoque = (56, 16) if alta else (40, 12)

        # estrelinhas
        sorteio = random.Random(31)
        for _ in range(int(altura / 14)):
            x, y = sorteio.uniform(6, largura - 6), sorteio.uniform(4, altura * 0.6)
            r = sorteio.choice((0.8, 1.0, 1.3))
            draw.ellipse([(x - r) * S, (y - r) * S, (x + r) * S, (y + r) * S], fill=(255, 245, 210, sorteio.choice((120, 170, 220))))

        # lua com brilho
        img = _brilho(img, lua, raio_brilho, (255, 233, 168), 70, desfoque)
        draw = ImageDraw.Draw(img)
        raio = 33
        draw.ellipse([(lua[0] - raio) * S, (lua[1] - raio) * S, (lua[0] + raio) * S, (lua[1] + raio) * S], fill=_rgba("#FFE9A8"))
        for dx, dy, r in ((-11, -9, 6), (12, 7, 8), (-4, 15, 4), (15, -15, 4), (-18, 9, 3.5)):
            draw.ellipse(
                [(lua[0] + dx - r) * S, (lua[1] + dy - r) * S, (lua[0] + dx + r) * S, (lua[1] + dy + r) * S],
                fill=_rgba("#EBCB85"),
            )

        # morcegos (posições relativas à lua, pra acompanhar ela)
        morcego = _rgba("#0E0619")
        morcegos = [(-6, -2, 1.15), (-92, -30, 0.8), (-54, 36, 0.6), (52, 54, 0.55)]
        if alta:  # mais espaço: mais morcegos descendo em direção à abóbora
            morcegos += [(-122, 118, 0.7), (28, 168, 0.5), (-40, 196, 0.45)]
        for dx, dy, escala in morcegos:
            _morcego(draw, lua[0] + dx, lua[1] + dy, escala, morcego)

    # morro
    draw.ellipse([(-30) * S, (altura - 34) * S, (largura + 30) * S, (altura + 60) * S], fill=_rgba("#170A28"))

    # lápides
    lapide, borda = _rgba("#3B2263"), _rgba("#5A3A86")
    for cx, topo, w, h in ((186, altura - 74, 28, 56), (216, altura - 56, 22, 40)):
        draw.rectangle([(cx - w / 2) * S, (topo + w / 2) * S, (cx + w / 2) * S, (topo + h + 6) * S], fill=lapide)
        draw.ellipse([(cx - w / 2) * S, topo * S, (cx + w / 2) * S, (topo + w) * S], fill=lapide)
        draw.arc([(cx - w / 2) * S, topo * S, (cx + w / 2) * S, (topo + w) * S], 180, 360, fill=borda, width=int(1.4 * S))
    cx, cy = 186, altura - 52  # cruz na lápide maior
    draw.rectangle([(cx - 1.4) * S, (cy - 8) * S, (cx + 1.4) * S, (cy + 8) * S], fill=borda)
    draw.rectangle([(cx - 6) * S, (cy - 3) * S, (cx + 6) * S, (cy + 0.2) * S], fill=borda)

    # abóbora (com brilho laranja por trás)
    px, py = 62, altura - 40
    img = _brilho(img, (px, py), 44, (255, 140, 30), 90, 14)
    img.alpha_composite(_camada_abobora(), ((px - 50) * S, (py - 52) * S))

    return _final(img, largura, altura)


def estrelas(largura: int, altura: int, semente: int = 31) -> list[tuple[float, float, float, int]]:
    """Estrelas espalhadas numa área largura x altura: (x, y, raio, opacidade 0-255).
    Determinístico (mesma semente = mesmo céu), pra não "mudar de lugar" a cada
    abertura do programa."""
    sorteio = random.Random(semente)
    quantidade = int(largura * altura / 2200)
    return [
        (
            sorteio.uniform(5, largura - 5), sorteio.uniform(3, altura - 3),
            sorteio.choice((0.8, 1.0, 1.0, 1.3, 1.8)), sorteio.choice((90, 130, 170, 220, 255)),
        )
        for _ in range(quantidade)
    ]


def misturar(cor_fundo: str, cor_frente: tuple[int, int, int], alpha: int) -> str:
    """Cor (hex) de `cor_frente` com opacidade `alpha` sobre `cor_fundo` — o
    Canvas do Tk não tem transparência, então o "brilho fraco" das estrelas
    vem de misturar a cor delas com a do fundo."""
    fundo = _rgba(cor_fundo)[:3]
    mistura = [round(f * alpha / 255 + b * (1 - alpha / 255)) for f, b in zip(cor_frente, fundo)]
    return "#%02x%02x%02x" % tuple(mistura)


def logo_com_gosma(logo: Image.Image, escorrimento: int = 24, semente: int = 13) -> Image.Image:
    """Logo com gosma verde escorrendo por BAIXO das letras, em gotas. As
    letras ficam intactas (o nome do escritório tem que continuar legível):
    a gosma só nasce na base das letras que chegam ao pé do logo (o "RS" e a
    linha "& SOUSA") e pinga pra baixo. Devolve uma imagem `escorrimento` px
    mais alta que o logo, com o logo no topo."""
    S = _SS
    largura, altura = logo.size
    px = logo.getchannel("A").load()

    base = {}
    for x in range(largura):
        ys = [y for y in range(altura) if px[x, y] > 110]
        if ys and ys[-1] >= altura * 0.75:  # só letras que chegam ao pé do logo
            base[x] = ys[-1]

    sorteio = random.Random(semente)
    colunas = sorted(base)
    gotas = []
    if colunas:
        n = 7
        faixa = (colunas[-1] - colunas[0]) / n
        for i in range(n):
            ini_faixa = colunas[0] + faixa * i
            candidatas = [x for x in colunas if ini_faixa + faixa * 0.2 <= x <= ini_faixa + faixa * 0.8]
            if candidatas:
                x = sorteio.choice(candidatas)
                comprimento = sorteio.uniform(5, escorrimento - 5)
                gotas.append((x + 0.5, base[x], base[x] + comprimento, sorteio.uniform(0.9, 1.3)))

    def _desenhar(draw, cor, folga):
        for x, yb, fim, grossura in gotas:
            for xc in range(int(x) - 4, int(x) + 5):  # pocinha na base da letra, de onde a gota cai
                if xc in base:
                    afinando = 1 - abs(xc - x) / 5
                    draw.rectangle(
                        [xc * S, (base[xc] - 0.6 - folga) * S, (xc + 1) * S, (base[xc] + 0.4 + 1.2 * afinando + folga) * S],
                        fill=cor,
                    )
            g = grossura + folga
            draw.rounded_rectangle([(x - g) * S, yb * S, (x + g) * S, fim * S], radius=g * S, fill=cor)
            r = grossura * 1.8 + folga
            draw.ellipse([(x - r) * S, (fim - r) * S, (x + r) * S, (fim + r) * S], fill=cor)

    gosma = _novo(largura, altura + escorrimento)
    draw = ImageDraw.Draw(gosma)
    _desenhar(draw, _rgba("#2F8F1F"), 0.6)   # contorno escuro
    _desenhar(draw, _rgba("#7EDE3C"), 0.0)   # gosma
    for x, _yb, fim, grossura in gotas:      # brilho nas gotas
        r = grossura * 0.6
        cx, cy = x - grossura * 0.7, fim - grossura * 0.8
        draw.ellipse([(cx - r) * S, (cy - r) * S, (cx + r) * S, (cy + r) * S], fill=(235, 255, 220, 230))

    resultado = Image.new("RGBA", (largura, altura + escorrimento), (0, 0, 0, 0))
    resultado.alpha_composite(logo, (0, 0))
    resultado.alpha_composite(_final(gosma, largura, altura + escorrimento))
    return resultado


def silhueta_morcego(escala: float = 1.0, asas_erguidas: bool = False) -> list[tuple[float, float]]:
    """Pontos (relativos ao centro) da silhueta do morcego, pra animar num
    Canvas do Tk. Com `asas_erguidas`, as pontas das asas sobem — alternar
    entre as duas dá o bater de asas."""
    direita = []
    for x, y in _MORCEGO:
        if asas_erguidas and x > 6:
            y = -y * 0.6 - (x - 6) * 0.35
        direita.append((x * escala, y * escala))
    esquerda = [(-x, y) for x, y in reversed(direita)]
    return direita + esquerda


_FONTE_NEGRITO = r"C:\Windows\Fonts\segoeuib.ttf"


def fantasma_fap(largura: int = 34, cor: str = "#4D69A3") -> Image.Image:
    """Fantasma clássico de lençol no azul do selo do FAP, com "FAP" escrito
    no corpo: mais alto que largo, olhos ocos, bracinhos erguidos de "buuu" e
    o pé terminando em pontas balançando pro lado. Corpo sólido com contorno
    fino, sem brilho em volta (fica sobre o fundo branco do conteúdo).
    Fundo transparente, com uma pequena margem pros braços."""
    from PIL import ImageFont

    S = _SS
    margem = 6
    altura_corpo = round(largura * 1.45)
    W, H = largura + 2 * margem, altura_corpo + 4
    img = _novo(W, H)
    draw = ImageDraw.Draw(img)

    raio = largura / 2
    cx, topo = W / 2, 2
    pe = topo + altura_corpo  # ponta mais baixa do lençol
    contorno, corpo = _rgba("#33497A"), _rgba(cor)

    def _silhueta(folga: float) -> list[tuple[float, float]]:
        pontos = []
        for i in range(19):  # cabeça arredondada (meia-volta de cima)
            a = math.pi + math.pi * i / 18
            pontos.append((cx + (raio + folga) * math.cos(a), topo + raio + (raio + folga) * math.sin(a)))
        esq, dir_ = cx - raio - folga, cx + raio + folga
        pontos += [(dir_ + 1.2, pe - 10)]  # lateral direita, alargando um pouco
        # pé em 3 pontas, todas puxadas pra direita (como se flutuasse pra esquerda)
        pontos += [
            (dir_ + 3.0, pe + folga), (dir_ - largura * 0.26, pe - 7),
            (cx + largura * 0.12, pe - 1 + folga), (cx - largura * 0.08, pe - 8),
            (esq + largura * 0.10, pe - 2 + folga), (esq - 0.6, pe - 9),
        ]
        return [(x * S, y * S) for x, y in pontos]

    # bracinhos erguidos de "buuu"
    for lado in (-1, 1):
        base_x = cx + lado * (raio - 1)
        braco = [(base_x, topo + raio + 6), (base_x + lado * 7.5, topo + raio - 1.5), (base_x + lado * 5.5, topo + raio + 2.5), (base_x, topo + raio + 11)]
        draw.polygon([(x * S, y * S) for x, y in braco], fill=corpo, outline=contorno, width=max(1, int(0.9 * S)))

    draw.polygon(_silhueta(0.9), fill=contorno)
    draw.polygon(_silhueta(0.0), fill=corpo)

    # olhos ocos, levemente puxados pra fora, e boca de "uuu"
    olho_y = topo + raio * 0.80
    for lado in (-1, 1):
        ox = cx + lado * largura * 0.19
        draw.ellipse([(ox - 2.6) * S, (olho_y - 3.8) * S, (ox + 2.6) * S, (olho_y + 3.8) * S], fill=_rgba("#120A24"))
    draw.ellipse([(cx - 1.8) * S, (olho_y + 3.6) * S, (cx + 1.8) * S, (olho_y + 7.6) * S], fill=_rgba("#120A24"))

    # "FAP" em branco no corpo, como no selo
    fonte = ImageFont.truetype(_FONTE_NEGRITO, int(largura * 0.36 * S))
    x0, y0, x1, y1 = draw.textbbox((0, 0), "FAP", font=fonte)
    tx = cx * S - (x1 - x0) / 2 - x0
    ty = (topo + altura_corpo * 0.64) * S - (y1 - y0) / 2 - y0
    draw.text((tx, ty), "FAP", font=fonte, fill=(255, 255, 255, 255))

    return _final(img, W, H)


# Piadinhas do fantasma do FAP: os "sustos" de quem trabalha com FAP. Cada uma
# presa a uma regra real da metodologia (Resolução CNPS 1.347/2021, que
# consolidou as Res. 1.316/2010, 1.329/2017 e 1.335/2018) — revisar se a
# metodologia mudar.
PIADAS_FAP = (
    # trava de morte/invalidez: sem bonificação (FAP não fica abaixo de 1,0000)
    "Buuu! Uma invalidez no período-base travou sua bonificação!",
    # trava de rotatividade: taxa média acima de 75% → FAP 1,0000
    "Buuu! A rotatividade passou de 75%... e a bonificação foi junto!",
    # teto do intervalo (0,5000 a 2,0000)
    "Buuu! Seu FAP acordou valendo 2,0000!",
    # CAT não emitida constatada em fiscalização → FAP 2,0000
    "Buuu! Acharam uma CAT que não foi emitida... FAP direto pra 2,0000!",
    # benefício acidentário sem CAT vinculada também entra na frequência
    "Buuu! Um B91 sem CAT entrou no seu índice de frequência!",
    # NTEP: auxílio-doença comum reconhecido como acidentário
    "Buuu! O NTEP transformou seu B31 em B91!",
    # acidente de trajeto é excluído do cálculo (desde a Res. 1.329/2017)
    "Buuu! Um acidente de trajeto escapou da exclusão e entrou no cálculo!",
    # peso da pensão por morte (B93) no índice de gravidade
    "Buuu! Uma pensão por morte pesa 0,50 na sua gravidade!",
    # benefícios contados no CNPJ completo (14 dígitos), por estabelecimento
    "Buuu! Caiu no seu CNPJ um benefício de empregado de outra empresa!",
    "Buuu! Aquela filial esquecida também tem FAP próprio!",
    # elementos contestáveis: benefícios, massa salarial, vínculos, rotatividade
    "Buuu! O mesmo benefício foi contado duas vezes!",
    "Buuu! A massa salarial do FAP não bate com a sua folha!",
    "Buuu! O número médio de vínculos veio errado!",
    # malus: IC acima de 1 eleva o FAP acima de 1,0000
    "Buuu! Seu índice composto passou de 1... bem-vindo ao malus!",
    # contestação
    "Buuu! Contestação indeferida sem análise do mérito!",
    "Buuu! Novembro chegou... e o prazo da contestação também!",
)


def icone_bruxa_vassoura(altura: int = 20) -> Image.Image:
    """Bruxinha voando de vassoura (pros botões "Limpar"). Fundo transparente,
    proporção 3:2 (mais larga que alta). Escura o bastante pra aparecer tanto
    no branco quanto no vermelho do botão com o mouse em cima."""
    S = _SS
    W, H = 30, 20
    img = _novo(W, H)
    d = ImageDraw.Draw(img)

    def p(*pts):
        return [(x * S, y * S) for x, y in pts]

    # vassoura: cabo e palha (a palha fica atrás, à esquerda: ela voa pra direita)
    d.line(p((6, 15.2), (28.5, 11.0)), fill=_rgba("#6B3F17"), width=int(1.7 * S))
    d.polygon(p((7.5, 14.9), (0.6, 11.6), (0.2, 15.6), (1.0, 19.0), (7.8, 16.4)), fill=_rgba("#D9A441"))
    for y0, y1 in ((12.6, 14.9), (15.4, 15.6), (18.0, 16.2)):  # fios da palha
        d.line(p((1.0, y0), (7.2, y1)), fill=_rgba("#A5752A"), width=max(1, int(0.6 * S)))
    d.line(p((7.6, 14.8), (7.9, 16.5)), fill=_rgba("#7A3E1A"), width=int(1.2 * S))  # amarração

    # capa esvoaçando pra trás e corpo sentado no cabo
    d.polygon(p((15.5, 13.8), (9.0, 9.8), (11.0, 12.6), (8.6, 13.6), (13.2, 15.6), (20.0, 14.4), (19.0, 8.2), (16.0, 8.0)), fill=_rgba("#2A1446"))
    d.polygon(p((19.6, 13.6), (22.6, 16.6), (24.2, 16.2), (21.0, 13.0)), fill=_rgba("#2A1446"))  # perna
    # rosto verde, cabelo laranja, chapéu pontudo com fita laranja
    d.ellipse(p((16.2, 4.4), (20.4, 8.6)), fill=_rgba("#7EC850"))
    d.polygon(p((16.4, 6.0), (13.6, 8.6), (15.8, 8.8), (16.6, 7.6)), fill=_rgba("#E8650A"))
    d.polygon(p((14.4, 5.4), (22.6, 4.8), (21.8, 5.9), (15.0, 6.4)), fill=_rgba("#1A0F2B"))  # aba
    d.polygon(p((16.2, 5.2), (21.2, 4.9), (17.2, -0.2), (14.2, 0.4)), fill=_rgba("#1A0F2B"))  # copa
    d.line(p((16.4, 4.6), (21.0, 4.3)), fill=_rgba("#E8650A"), width=max(1, int(0.9 * S)))
    largura = round(altura * W / H)
    return img.resize((largura, altura), Image.LANCZOS)


def icone_caldeirao(altura: int = 20) -> Image.Image:
    """Caldeirão borbulhando de poção verde (pro botão "Gerar anexo completo").
    Fundo transparente; escuro e verde, pra aparecer no laranja do botão."""
    S = _SS
    W, H = 22, 22
    img = _novo(W, H)
    d = ImageDraw.Draw(img)

    def p(*pts):
        return [(x * S, y * S) for x, y in pts]

    # bolhas e respingos subindo
    for cx, cy, r in ((8.0, 4.2, 1.6), (13.2, 2.6, 1.2), (15.6, 5.4, 1.4), (10.8, 0.9, 0.8)):
        d.ellipse(p((cx - r, cy - r), (cx + r, cy + r)), fill=_rgba("#7EDE3C"), outline=_rgba("#2F8F1F"), width=max(1, int(0.5 * S)))
    # perninhas
    for x in (5.2, 16.8):
        d.polygon(p((x - 1.3, 18.0), (x + 1.3, 18.0), (x + 0.6, 21.4), (x - 0.6, 21.4)), fill=_rgba("#14091F"))
    # corpo do caldeirão
    d.ellipse(p((1.8, 7.4), (20.2, 20.0)), fill=_rgba("#1E1426"))
    d.ellipse(p((4.0, 9.4), (8.6, 13.4)), fill=_rgba("#3D3150"))  # reflexo (opaco: ver fantasma_fap)
    # borda e poção
    d.ellipse(p((1.0, 6.2), (21.0, 10.6)), fill=_rgba("#3A3046"))
    d.ellipse(p((2.6, 6.9), (19.4, 9.9)), fill=_rgba("#7EDE3C"))
    d.ellipse(p((6.0, 7.3), (10.0, 8.6)), fill=_rgba("#C9F7A6"))  # brilho na poção
    largura = round(altura * W / H)
    return img.resize((largura, altura), Image.LANCZOS)


def icone_tumulo(altura: int = 20) -> Image.Image:
    """Lápide com cruz e grama (pro botão "Gerar capas"). Pedra clara com
    contorno escuro, pra aparecer no laranja do botão."""
    S = _SS
    W, H = 18, 22
    img = _novo(W, H)
    d = ImageDraw.Draw(img)

    def p(*pts):
        return [(x * S, y * S) for x, y in pts]

    contorno = _rgba("#2B2638")
    # pedra: retângulo com topo arredondado
    d.rectangle(p((2.4, 8.0), (15.6, 19.6)), fill=_rgba("#A8A4B8"), outline=contorno, width=int(1.0 * S))
    d.pieslice(p((2.4, 1.6), (15.6, 14.4)), 180, 360, fill=_rgba("#A8A4B8"), outline=contorno, width=int(1.0 * S))
    d.rectangle(p((3.4, 7.6), (14.6, 9.0)), fill=_rgba("#A8A4B8"))  # apaga a emenda do arco
    # cruz gravada e uma rachadura
    d.line(p((9.0, 5.0), (9.0, 13.6)), fill=contorno, width=int(1.4 * S))
    d.line(p((6.4, 7.6), (11.6, 7.6)), fill=contorno, width=int(1.4 * S))
    d.line(p((12.8, 12.0), (11.6, 14.2), (12.6, 16.0)), fill=_rgba("#5E596E"), width=max(1, int(0.6 * S)))
    # grama na base
    d.polygon(p((0.4, 21.6), (1.6, 18.4), (3.0, 20.2), (4.6, 17.6), (6.4, 20.0), (8.6, 18.0), (10.6, 20.2),
                (12.8, 17.8), (14.6, 20.0), (16.4, 18.2), (17.6, 21.6)), fill=_rgba("#2F7A1F"))
    largura = round(altura * W / H)
    return img.resize((largura, altura), Image.LANCZOS)


def icone_machado(altura: int = 20) -> Image.Image:
    """Machado de lenhador com uma gotinha vermelha (pro botão "Dividir PDF")."""
    S = _SS
    W, H = 22, 22
    img = _novo(W, H)
    d = ImageDraw.Draw(img)

    def p(*pts):
        return [(x * S, y * S) for x, y in pts]

    contorno = _rgba("#2B2638")
    # cabo na diagonal
    d.line(p((4.0, 20.4), (15.6, 4.6)), fill=contorno, width=int(3.0 * S))
    d.line(p((4.0, 20.4), (15.6, 4.6)), fill=_rgba("#8A5426"), width=int(1.8 * S))
    # lâmina (meia-lua) presa no topo do cabo
    lamina = p((12.4, 3.2), (16.0, 0.6), (21.0, 4.4), (21.4, 10.4), (18.2, 11.2), (14.6, 7.8))
    d.polygon(lamina, fill=_rgba("#C9CED8"), outline=contorno, width=int(1.0 * S))
    d.line(p((20.4, 4.6), (20.8, 10.0)), fill=_rgba("#FFFFFF"), width=max(1, int(0.8 * S)))  # fio brilhando
    # gota pingando do fio
    d.ellipse(p((19.2, 13.0), (21.6, 15.6)), fill=_rgba("#C0141C"))
    d.polygon(p((19.4, 13.8), (20.4, 11.6), (21.4, 13.8)), fill=_rgba("#C0141C"))
    largura = round(altura * W / H)
    return img.resize((largura, altura), Image.LANCZOS)


def icone_pocao(altura: int = 20) -> Image.Image:
    """Frasco redondo de poção roxa borbulhando (pro botão "Juntar PDFs")."""
    S = _SS
    W, H = 18, 22
    img = _novo(W, H)
    d = ImageDraw.Draw(img)

    def p(*pts):
        return [(x * S, y * S) for x, y in pts]

    contorno = _rgba("#1E1426")
    # bolhas saindo
    for cx, cy, r in ((11.6, 1.4, 0.9), (13.6, 3.4, 0.7)):
        d.ellipse(p((cx - r, cy - r), (cx + r, cy + r)), fill=_rgba("#C792FF"), outline=contorno, width=max(1, int(0.4 * S)))
    # gargalo e rolha
    d.rectangle(p((6.4, 4.6), (11.6, 10.0)), fill=_rgba("#DCD6EA"), outline=contorno, width=int(0.9 * S))
    d.rectangle(p((6.0, 2.6), (12.0, 5.2)), fill=_rgba("#8A5426"), outline=contorno, width=int(0.8 * S))
    # bojo de vidro com a poção
    d.ellipse(p((1.0, 7.8), (17.0, 21.6)), fill=_rgba("#DCD6EA"), outline=contorno, width=int(1.0 * S))
    d.chord(p((2.0, 8.8), (16.0, 20.6)), 10, 170, fill=_rgba("#8E3FE8"))
    d.ellipse(p((4.2, 10.6), (6.6, 13.4)), fill=_rgba("#FFFFFF"))  # reflexo (opaco)
    for cx, cy, r in ((7.4, 16.6, 0.9), (11.0, 18.0, 0.7), (10.0, 15.2, 0.5)):
        d.ellipse(p((cx - r, cy - r), (cx + r, cy + r)), fill=_rgba("#D8B4FF"))
    largura = round(altura * W / H)
    return img.resize((largura, altura), Image.LANCZOS)


# Frase que abre o resultado quando uma operação termina bem
MENSAGENS_SUCESSO = (
    "🎃 Feitiço concluído com sucesso!",
    "🧙 Anexo conjurado! Nenhum espírito foi ferido.",
    "🦇 Pronto! Nem o fantasma do FAP achou erro.",
    "🕯️ O ritual terminou — os PDFs saíram do caldeirão.",
    "👻 Tudo certo! Pode respirar, não foi susto dessa vez.",
)

# O fantasma consolando quando dá erro em alguma operação
PIADAS_ERRO = (
    "Buuu! Até eu levei susto com esse erro!",
    "Buuu! Esse erro não fui eu, juro!",
    "Buuu! Calma, é só um erro... não é o FAP 2,0000.",
    "Buuu! Que assombração, hein? Tenta de novo!",
)


def balao_fala(largura: int, altura: int, cor_fundo: str, cor_borda: str, cor_externa: str,
               cor_chave: str, cauda: int = 15) -> Image.Image:
    """Balão de fala com cantos arredondados e a pontinha (cauda) do lado
    direito, apontando pro fantasma. `largura` inclui a cauda.

    Sai em RGB, com o lado de fora pintado de `cor_chave` — a janela do balão
    usa essa cor como transparente (-transparentcolor do Windows), que não tem
    meio-termo: por isso as bordas suavizadas são misturadas com `cor_externa`
    (o branco do conteúdo, que fica atrás do balão) em vez de semitransparentes."""
    S = _SS
    corpo = largura - cauda - 1
    forma = Image.new("L", (largura * S, altura * S), 0)
    miolo = Image.new("L", forma.size, 0)
    meio = min(17, altura / 2)  # altura da cauda: perto do topo, na linha do rosto do fantasma
    for camada, recuo in ((forma, 0.0), (miolo, 1.4)):
        d = ImageDraw.Draw(camada)
        r = (recuo) * S
        d.rounded_rectangle([r, r, corpo * S - r, (altura - 1) * S - r], radius=10 * S - r, fill=255)
        d.polygon([
            ((corpo - 4) * S, (meio - 9) * S + r),
            ((largura - 1) * S - r * 1.8, (meio + 3) * S),
            ((corpo - 4) * S, (meio + 6) * S - r),
        ], fill=255)
    img = Image.new("RGB", forma.size, cor_externa)
    img.paste(Image.new("RGB", forma.size, cor_borda), mask=forma)
    img.paste(Image.new("RGB", forma.size, cor_fundo), mask=miolo)
    img = img.resize((largura, altura), Image.LANCZOS)
    alfa = forma.resize((largura, altura), Image.LANCZOS)
    fora = alfa.point(lambda a: 255 if a < 110 else 0)
    img.paste(Image.new("RGB", img.size, cor_chave), mask=fora)
    return img
