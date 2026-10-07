# -*- coding: utf-8 -*-
"""Gera os SVGs ANIMADOS (arte original inspirada em Operação Big Hero) em docs/.
Animações via SMIL/CSS dentro do SVG: funcionam quando o SVG é usado com <img> no README do GitHub."""
import math
import os
import random
import sys
from xml.etree import ElementTree as ET

OUT = sys.argv[1]
os.makedirs(OUT, exist_ok=True)

RED = "#CA1F26"       # vermelho do logo (referência: SchemeColor)
RED_L = "#ff5a60"     # vermelho claro para textos sobre fundo escuro
WHITE = "#F6FFFF"     # branco do logo
INK = "#1b1030"       # roxo-noite (texto sobre balão claro, fundo dos cartões)
NS = 'xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink"'
FONT = "'Segoe UI', 'Helvetica Neue', Arial, sans-serif"
MONO = "Consolas, 'Courier New', monospace"


def save(name, svg):
    ET.fromstring(svg)                       # garante XML bem formado
    with open(os.path.join(OUT, name), "w", encoding="utf-8", newline="\n") as f:
        f.write(svg)
    print(f"{name}: {len(svg)/1024:.1f} KB")


def f1(x):
    return f"{x:.1f}"


# ------------------------------------------------------------------ peças reutilizáveis
def baymax(prefix, scan=True):
    """Baymax estilizado (arte própria): corpo/cabeça brancos, olhos-pontos ligados por uma linha,
    cruz vermelha no peito (assistente de saúde). Coordenadas centradas em (0,0)."""
    g = f'''
  <defs>
    <linearGradient id="{prefix}bg" x1="0" y1="0" x2="0" y2="1">
      <stop offset="0" stop-color="#ffffff"/><stop offset="1" stop-color="#d6e3ec"/>
    </linearGradient>
    <clipPath id="{prefix}clip"><ellipse cx="0" cy="40" rx="58" ry="70"/></clipPath>
  </defs>
  <ellipse cx="0" cy="118" rx="46" ry="8" fill="#000" opacity=".33">
    <animate attributeName="rx" values="46;38;46" dur="3.2s" repeatCount="indefinite" calcMode="spline" keyTimes="0;.5;1" keySplines=".45 0 .55 1;.45 0 .55 1"/>
    <animate attributeName="opacity" values=".33;.2;.33" dur="3.2s" repeatCount="indefinite"/>
  </ellipse>
  <g>
    <ellipse cx="-24" cy="102" rx="20" ry="16" fill="url(#{prefix}bg)" stroke="#c4d3de" stroke-width="1.5"/>
    <ellipse cx="24" cy="102" rx="20" ry="16" fill="url(#{prefix}bg)" stroke="#c4d3de" stroke-width="1.5"/>
    <g>
      <animateTransform attributeName="transform" type="rotate" values="0 -56 8;5 -56 8;0 -56 8" dur="3.2s" repeatCount="indefinite"/>
      <ellipse cx="-66" cy="38" rx="16" ry="38" transform="rotate(10 -66 38)" fill="url(#{prefix}bg)" stroke="#c4d3de" stroke-width="1.5"/>
    </g>
    <g>
      <animateTransform attributeName="transform" type="rotate" values="0 56 8;-5 56 8;0 56 8" dur="3.2s" repeatCount="indefinite"/>
      <ellipse cx="66" cy="38" rx="16" ry="38" transform="rotate(-10 66 38)" fill="url(#{prefix}bg)" stroke="#c4d3de" stroke-width="1.5"/>
    </g>
    <ellipse cx="0" cy="40" rx="58" ry="70" fill="url(#{prefix}bg)" stroke="#c4d3de" stroke-width="1.5"/>'''
    if scan:
        g += f'''
    <g clip-path="url(#{prefix}clip)">
      <rect x="-60" y="-40" width="120" height="6" fill="#7fe3ff" opacity=".55">
        <animate attributeName="y" values="-34;108;-34" dur="4.5s" repeatCount="indefinite"/>
        <animate attributeName="opacity" values="0;.6;.6;0" keyTimes="0;.1;.9;1" dur="4.5s" repeatCount="indefinite"/>
      </rect>
    </g>'''
    g += f'''
    <g transform="translate(0 34)">
      <animateTransform attributeName="transform" type="translate" values="0 34;0 34" dur="1s" repeatCount="indefinite"/>
      <g>
        <animateTransform attributeName="transform" type="scale" values="1;1.16;1;1.1;1" keyTimes="0;.12;.25;.37;1" dur="1.6s" repeatCount="indefinite"/>
        <rect x="-4.5" y="-13" width="9" height="26" rx="2.5" fill="{RED}"/>
        <rect x="-13" y="-4.5" width="26" height="9" rx="2.5" fill="{RED}"/>
      </g>
    </g>
    <ellipse cx="0" cy="-38" rx="44" ry="34" fill="url(#{prefix}bg)" stroke="#c4d3de" stroke-width="1.5"/>
    <line x1="-17" y1="-38" x2="17" y2="-38" stroke="#1a1a22" stroke-width="2.2" stroke-linecap="round"/>
    <ellipse cx="-17" cy="-38" rx="5.6" ry="5.6" fill="#1a1a22">
      <animate attributeName="ry" values="5.6;5.6;.6;5.6;5.6" keyTimes="0;.9;.935;.97;1" dur="5s" repeatCount="indefinite"/>
    </ellipse>
    <ellipse cx="17" cy="-38" rx="5.6" ry="5.6" fill="#1a1a22">
      <animate attributeName="ry" values="5.6;5.6;.6;5.6;5.6" keyTimes="0;.9;.935;.97;1" dur="5s" repeatCount="indefinite"/>
    </ellipse>
  </g>'''
    return g


def buildings(rng, x0, x1, base, hmin, hmax, fill, windows=True, wcolor="#ffd98a", pagodas=True):
    """Prédios + janelas piscando. Retorna (svg, lista de topos)."""
    out, x = [], x0
    while x < x1:
        w = rng.randint(34, 78)
        h = rng.randint(hmin, hmax)
        y = base - h
        out.append(f'<rect x="{x}" y="{y}" width="{w}" height="{h}" fill="{fill}"/>')
        if pagodas and rng.random() < .22 and w >= 44:   # telhado de pagode: toque "Tokyo"
            out.append(f'<polygon points="{x-4},{y} {x+w//2},{y-16} {x+w+4},{y}" fill="{fill}"/>')
        elif rng.random() < .25:
            out.append(f'<rect x="{x+w//2}" y="{y-14}" width="2" height="14" fill="{fill}"/>')
        if windows:
            cols, rows = (w - 8) // 9, (h - 10) // 12
            for r in range(rows):
                for c in range(cols):
                    if rng.random() < .30:
                        wx, wy = x + 5 + 9 * c, y + 6 + 12 * r
                        if rng.random() < .35:
                            d = rng.uniform(2.5, 8)
                            out.append(f'<rect x="{wx}" y="{wy}" width="4" height="6" fill="{wcolor}"><animate attributeName="opacity" values="1;.15;1" dur="{d:.1f}s" begin="-{rng.uniform(0, 8):.1f}s" repeatCount="indefinite"/></rect>')
                        else:
                            out.append(f'<rect x="{wx}" y="{wy}" width="4" height="6" fill="{wcolor}" opacity=".85"/>')
        x += w + rng.randint(0, 6)
    return "\n".join(out)


def torii(cx, deck_y):
    return f'''
  <g transform="translate({cx} {deck_y})">
    <rect x="-30" y="-120" width="11" height="120" fill="{RED}"/><rect x="19" y="-120" width="11" height="120" fill="{RED}"/>
    <rect x="-44" y="-92" width="88" height="9" fill="{RED}"/><rect x="-44" y="-83" width="88" height="3" fill="#7d1015"/>
    <path d="M-64 -126 Q0 -146 64 -126 L57 -111 Q0 -125 -57 -111 Z" fill="{RED}"/>
    <path d="M-57 -111 Q0 -125 57 -111 L57 -108 Q0 -121 -57 -108 Z" fill="#7d1015"/>
    <circle cx="0" cy="-136" r="3" fill="#fff"><animate attributeName="opacity" values="1;.1;1" dur="1.8s" repeatCount="indefinite"/></circle>
  </g>'''


def qpt(p0, p1, p2, t):
    x = (1 - t) ** 2 * p0[0] + 2 * (1 - t) * t * p1[0] + t ** 2 * p2[0]
    y = (1 - t) ** 2 * p0[1] + 2 * (1 - t) * t * p1[1] + t ** 2 * p2[1]
    return x, y


def cable(p0, p1, p2, deck, hangers=True, pulses=0, dur=6):
    d = f"M{p0[0]} {p0[1]} Q{p1[0]} {p1[1]} {p2[0]} {p2[1]}"
    out = [f'<path d="{d}" fill="none" stroke="#e8c9d6" stroke-width="2" opacity=".85"/>']
    if hangers:
        for i in range(1, 15):
            x, y = qpt(p0, p1, p2, i / 15)
            if 0 < x < 1200 and y < deck - 6:
                out.append(f'<line x1="{f1(x)}" y1="{f1(y)}" x2="{f1(x)}" y2="{deck}" stroke="#e8c9d6" stroke-width="1" opacity=".5"/>')
    for k in range(pulses):
        out.append(f'<circle r="3.2" fill="{RED_L}"><animateMotion dur="{dur}s" begin="-{k*dur/pulses:.1f}s" repeatCount="indefinite" path="{d}"/></circle>')
    return "\n".join(out), d


# ================================================================== 1) BANNER
def banner():
    rng = random.Random(7)
    W, H, SEA, DECK = 1200, 400, 359, 352
    stars = []
    for _ in range(48):
        stars.append(f'<circle cx="{rng.randint(6,1194)}" cy="{rng.randint(6,190)}" r="{rng.uniform(.6,1.8):.1f}" fill="#fff"><animate attributeName="opacity" values=".1;1;.1" dur="{rng.uniform(2,5):.1f}s" begin="-{rng.uniform(0,5):.1f}s" repeatCount="indefinite"/></circle>')
    back = buildings(rng, -10, 1210, SEA, 50, 100, "#3a1b5e", windows=False)
    front = buildings(rng, -10, 1210, SEA, 28, 84, "#1c0e38", windows=True)
    # cabos
    T1, T2, TOP = 330, 870, 236
    main, _ = cable((T1, TOP), (600, 424), (T2, TOP), DECK, pulses=5, dur=7)
    left, _ = cable((T1, TOP), (150, 330), (-20, 346), DECK, pulses=2, dur=5)
    right, _ = cable((T2, TOP), (1050, 330), (1220, 346), DECK, pulses=2, dur=5)
    # lanternas penduradas no cabo principal
    lant = []
    for i, t in enumerate((.22, .5, .78)):
        x, y = qpt((T1, TOP), (600, 424), (T2, TOP), t)
        lant.append(f'''<g transform="translate({f1(x)} {f1(y)})"><g><animateTransform attributeName="transform" type="rotate" values="-7;7;-7" dur="{3+i*.7:.1f}s" begin="-{i}s" repeatCount="indefinite" calcMode="spline" keyTimes="0;.5;1" keySplines=".45 0 .55 1;.45 0 .55 1"/>
          <line x1="0" y1="0" x2="0" y2="14" stroke="#e8c9d6" stroke-width="1"/>
          <ellipse cx="0" cy="22" rx="7" ry="9" fill="{RED}"/><ellipse cx="0" cy="22" rx="12" ry="14" fill="url(#glow)"/>
          <rect x="-4" y="12" width="8" height="2.5" fill="#3b0d10"/><rect x="-4" y="30" width="8" height="2.5" fill="#3b0d10"/></g></g>''')
    # mar com brilhos
    sea = []
    for _ in range(26):
        x, y, w = rng.randint(0, 1160), rng.randint(SEA + 6, H - 4), rng.randint(18, 70)
        sea.append(f'<rect x="{x}" y="{y}" width="{w}" height="1.6" rx=".8" fill="#ffb36b" opacity=".4"><animate attributeName="opacity" values=".08;.55;.08" dur="{rng.uniform(2,5):.1f}s" begin="-{rng.uniform(0,5):.1f}s" repeatCount="indefinite"/></rect>')
    # enxame de microbots (fita seguindo um caminho até o Baymax)
    swarm = "M-20 292 C250 262 470 340 700 276 S930 198 1004 212"
    bots = []
    N, D = 70, 9
    for i in range(N):
        off = rng.uniform(-16, 16)
        sz = rng.uniform(4, 7)
        bots.append(f'''<g opacity="0"><rect x="{-sz/2:.1f}" y="{f1(off - sz/2)}" width="{sz:.1f}" height="{sz:.1f}" rx="1.3" fill="#15121f" stroke="#8fe8ff" stroke-width=".7"/>
        <animateMotion dur="{D}s" begin="-{i*D/N + rng.uniform(0,.08):.2f}s" repeatCount="indefinite" rotate="auto"><mpath xlink:href="#swarm"/></animateMotion>
        <animate attributeName="opacity" values="0;1;1;0" keyTimes="0;.07;.9;1" dur="{D}s" begin="-{i*D/N:.2f}s" repeatCount="indefinite"/></g>''')
    clouds = []
    for i, (y, s) in enumerate(((70, 1), (130, .7), (40, .5))):
        clouds.append(f'<g opacity=".16"><animateTransform attributeName="transform" type="translate" from="-340 0" to="1400 0" dur="{80+i*25}s" begin="-{i*31}s" repeatCount="indefinite"/><ellipse cx="0" cy="{y}" rx="{150*s:.0f}" ry="{16*s:.0f}" fill="#ffd5e8"/><ellipse cx="{60*s:.0f}" cy="{y-9}" rx="{90*s:.0f}" ry="{13*s:.0f}" fill="#ffd5e8"/></g>')

    svg = f'''<svg {NS} viewBox="0 0 {W} {H}" width="{W}" height="{H}" role="img" aria-labelledby="t d">
<title id="t">Painel de Montagem SoBot</title>
<desc id="d">Banner animado: San Fransokyo ao entardecer, ponte com torii, microbots e um assistente branco com cruz vermelha.</desc>
<defs>
  <linearGradient id="sky" x1="0" y1="0" x2="0" y2="1">
    <stop offset="0"><animate attributeName="stop-color" values="#170d33;#231245;#170d33" dur="14s" repeatCount="indefinite"/></stop>
    <stop offset=".45"><animate attributeName="stop-color" values="#5b1f6b;#7a2a72;#5b1f6b" dur="14s" repeatCount="indefinite"/></stop>
    <stop offset=".78"><animate attributeName="stop-color" values="#e0467c;#f0597a;#e0467c" dur="14s" repeatCount="indefinite"/></stop>
    <stop offset="1"><animate attributeName="stop-color" values="#ffb36b;#ffc98f;#ffb36b" dur="14s" repeatCount="indefinite"/></stop>
  </linearGradient>
  <radialGradient id="sun" cx=".5" cy=".5" r=".5"><stop offset="0" stop-color="#ffd9a0" stop-opacity=".9"/><stop offset="1" stop-color="#ff8a5c" stop-opacity="0"/></radialGradient>
  <radialGradient id="glow" cx=".5" cy=".5" r=".5"><stop offset="0" stop-color="{RED_L}" stop-opacity=".55"/><stop offset="1" stop-color="{RED_L}" stop-opacity="0"/></radialGradient>
  <linearGradient id="seaG" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#3a1b5e"/><stop offset="1" stop-color="#120a28"/></linearGradient>
  <path id="swarm" d="{swarm}" fill="none"/>
</defs>
<rect width="{W}" height="{H}" fill="url(#sky)"/>
{chr(10).join(stars)}
<circle cx="600" cy="{SEA}" r="300" fill="url(#sun)"><animate attributeName="opacity" values=".55;.9;.55" dur="6s" repeatCount="indefinite"/></circle>
{chr(10).join(clouds)}
{back}
{front}
<rect x="0" y="{SEA}" width="{W}" height="{H-SEA}" fill="url(#seaG)"/>
{chr(10).join(sea)}
{main}
{left}
{right}
{torii(T1, DECK)}
{torii(T2, DECK)}
{chr(10).join(lant)}
<rect x="-10" y="{DECK}" width="1220" height="7" fill="#120a28"/><rect x="-10" y="{DECK}" width="1220" height="1.2" fill="{RED_L}" opacity=".7"/>
{chr(10).join(bots)}
<g>
  <animateTransform attributeName="transform" type="translate" values="1010 214;1010 200;1010 214" dur="3.2s" repeatCount="indefinite" calcMode="spline" keyTimes="0;.5;1" keySplines=".45 0 .55 1;.45 0 .55 1"/>
  <g transform="scale(1.02)">{baymax("b")}</g>
</g>
<g font-family="{FONT}">
  <g opacity="0">
    <animate attributeName="opacity" values="0;1" dur=".9s" fill="freeze"/>
    <animateTransform attributeName="transform" type="translate" values="0 14;0 0" dur=".9s" fill="freeze"/>
    <text x="62" y="62" font-size="15" font-weight="700" letter-spacing="7" fill="{RED_L}">OPERAÇÃO · MONTAGEM</text>
    <text x="60" y="124" font-size="58" font-weight="800" fill="{WHITE}">Painel de Montagem</text>
    <text x="60" y="184" font-size="58" font-weight="800" fill="{RED_L}">SoBot</text>
    <rect x="62" y="196" width="0" height="5" rx="2.5" fill="{RED}"><animate attributeName="width" values="0;250" begin=".6s" dur=".8s" fill="freeze" calcMode="spline" keyTimes="0;1" keySplines=".2 .8 .2 1"/></rect>
  </g>
  <clipPath id="tw"><rect x="60" y="205" width="0" height="30"><animate attributeName="width" values="0;0;560;560;0" keyTimes="0;.06;.45;.92;1" dur="10s" repeatCount="indefinite"/></rect></clipPath>
  <text x="62" y="226" font-family="{MONO}" font-size="20" fill="#ffe3c2" clip-path="url(#tw)">Olá! Sou o seu assistente pessoal de montagem.</text>
  <rect y="208" width="10" height="22" fill="{RED_L}"><animate attributeName="x" values="62;62;622;622;62" keyTimes="0;.06;.45;.92;1" dur="10s" repeatCount="indefinite"/><animate attributeName="opacity" values="1;0;1" dur=".9s" repeatCount="indefinite"/></rect>
</g>
</svg>'''
    save("banner.svg", svg)


# ================================================================== 2) DIVISOR
def divider():
    rng = random.Random(3)
    bots = []
    for i in range(14):
        bots.append(f'<rect x="-3" y="-3" width="6" height="6" rx="1.3" fill="#5b6b8c" stroke="#8fe8ff" stroke-width=".6" opacity="0"><animateMotion dur="7s" begin="-{i*0.5:.1f}s" repeatCount="indefinite" rotate="auto" path="M0 {14+rng.uniform(-3,3):.1f} L900 {14+rng.uniform(-3,3):.1f}"/><animate attributeName="opacity" values="0;1;1;0" keyTimes="0;.1;.9;1" dur="7s" begin="-{i*0.5:.1f}s" repeatCount="indefinite"/></rect>')
    svg = f'''<svg {NS} viewBox="0 0 900 28" width="900" height="28" role="img" aria-label="divisor animado">
<defs><linearGradient id="l" x1="0" x2="1"><stop offset="0" stop-color="{RED}" stop-opacity="0"/><stop offset=".5" stop-color="{RED}"/><stop offset="1" stop-color="{RED}" stop-opacity="0"/></linearGradient></defs>
<rect x="0" y="13" width="900" height="2" fill="url(#l)" opacity=".7"/>
{chr(10).join(bots)}
<g transform="translate(450 14)"><circle r="12" fill="{WHITE}" stroke="{RED}" stroke-width="1.5"/>
  <g><animateTransform attributeName="transform" type="scale" values="1;1.18;1;1.12;1" keyTimes="0;.12;.25;.37;1" dur="1.6s" repeatCount="indefinite"/>
  <rect x="-2.2" y="-6.5" width="4.4" height="13" rx="1.3" fill="{RED}"/><rect x="-6.5" y="-2.2" width="13" height="4.4" rx="1.3" fill="{RED}"/></g></g>
</svg>'''
    save("divider.svg", svg)


# ================================================================== 3) BAYMAX DIZ
def baymax_diz():
    falas = [
        ("Olá! Sou o assistente pessoal de montagem.", "Vim cuidar do seu progresso."),
        ("Em uma escala de 1 a 10, quanto dói", "refazer a mesma conta em duas planilhas?"),
        ("Detectei etapas sem tempo ideal.", "Recomendo preenchê-las para eu calcular os pesos."),
        ("Progresso salvo na planilha.", "Você está satisfeito com o seu cuidado?"),
    ]
    grupos = []
    for i, (a, b) in enumerate(falas):
        grupos.append(f'''<g opacity="0" font-family="{FONT}" font-size="19" fill="{INK}">
    <animate attributeName="opacity" values="0;1;1;0;0" keyTimes="0;.03;.22;.25;1" dur="18s" begin="{i*4.5}s" repeatCount="indefinite"/>
    <text x="206" y="62" font-weight="700">{a}</text><text x="206" y="92">{b}</text></g>''')
    svg = f'''<svg {NS} viewBox="0 0 900 170" width="900" height="170" role="img" aria-label="Assistente animado com falas que se alternam">
<g>
  <animateTransform attributeName="transform" type="translate" values="92 74;92 66;92 74" dur="3.2s" repeatCount="indefinite" calcMode="spline" keyTimes="0;.5;1" keySplines=".45 0 .55 1;.45 0 .55 1"/>
  <g transform="scale(.62)">{baymax("m", scan=False)}</g>
</g>
<path d="M176 60 L196 48 L196 74 Z" fill="{WHITE}" stroke="{RED}" stroke-width="2" stroke-linejoin="round"/>
<rect x="190" y="20" width="690" height="104" rx="18" fill="{WHITE}" stroke="{RED}" stroke-width="2.5"/>
<rect x="191.5" y="48.5" width="4" height="26" fill="{WHITE}"/>
{chr(10).join(grupos)}
</svg>'''
    save("baymax-diz.svg", svg)


# ================================================================== 4) ARQUITETURA ANIMADA
def arquitetura():
    def card(x, titulo, arquivo, icone, cor):
        return f'''<g transform="translate({x} 40)">
  <rect width="230" height="170" rx="16" fill="url(#cg)" stroke="{cor}" stroke-width="2"/>
  <rect width="230" height="8" rx="4" fill="{cor}"/>
  {icone}
  <text x="115" y="124" text-anchor="middle" font-family="{FONT}" font-size="20" font-weight="800" fill="{WHITE}">{titulo}</text>
  <text x="115" y="148" text-anchor="middle" font-family="{MONO}" font-size="12.5" fill="#cbb8ee">{arquivo}</text></g>'''
    ic_browser = '<g transform="translate(75 30)"><rect width="80" height="62" rx="8" fill="#2a1a55" stroke="#8fe8ff" stroke-width="2"/><rect width="80" height="14" rx="8" fill="#8fe8ff" opacity=".9"/><circle cx="9" cy="7" r="2.4" fill="#1b1030"/><circle cx="17" cy="7" r="2.4" fill="#1b1030"/><rect x="10" y="24" width="26" height="30" rx="3" fill="#8fe8ff" opacity=".6"><animate attributeName="height" values="30;18;30" dur="3s" repeatCount="indefinite"/><animate attributeName="y" values="24;36;24" dur="3s" repeatCount="indefinite"/></rect><rect x="44" y="24" width="26" height="30" rx="3" fill="#ff5a60" opacity=".75"><animate attributeName="height" values="18;30;18" dur="3s" repeatCount="indefinite"/><animate attributeName="y" values="36;24;36" dur="3s" repeatCount="indefinite"/></rect></g>'
    ic_server = '<g transform="translate(75 28)"><rect width="80" height="20" rx="5" fill="#2a1a55" stroke="#ff5a60" stroke-width="2"/><rect y="25" width="80" height="20" rx="5" fill="#2a1a55" stroke="#ff5a60" stroke-width="2"/><rect y="50" width="80" height="20" rx="5" fill="#2a1a55" stroke="#ff5a60" stroke-width="2"/>' + "".join(f'<circle cx="{12}" cy="{10+25*i}" r="3.2" fill="#7CFF9B"><animate attributeName="opacity" values="1;.2;1" dur="{1.2+i*.5:.1f}s" repeatCount="indefinite"/></circle>' for i in range(3)) + '</g>'
    ic_xlsx = '<g transform="translate(75 28)"><rect width="80" height="66" rx="6" fill="#0f5132" stroke="#7CFF9B" stroke-width="2"/>' + "".join(f'<line x1="0" y1="{16*i}" x2="80" y2="{16*i}" stroke="#7CFF9B" stroke-width="1" opacity=".6"/>' for i in range(1, 4)) + "".join(f'<line x1="{20*i}" y1="0" x2="{20*i}" y2="66" stroke="#7CFF9B" stroke-width="1" opacity=".6"/>' for i in range(1, 4)) + '<rect x="21" y="17" width="19" height="15" fill="#7CFF9B" opacity=".5"><animate attributeName="opacity" values=".1;.8;.1" dur="2.4s" repeatCount="indefinite"/></rect></g>'

    def lane(x1, x2, y, rot, label, cor, dur):
        d = f"M{x1} {y} L{x2} {y}" if rot else f"M{x2} {y} L{x1} {y}"
        dots = "".join(f'<circle r="4.2" fill="{cor}"><animateMotion dur="{dur}s" begin="-{k*dur/4:.2f}s" repeatCount="indefinite" path="{d}"/></circle>' for k in range(4))
        head = f'<polygon points="{x2},{y} {x2-9},{y-5} {x2-9},{y+5}" fill="{cor}"/>' if rot else f'<polygon points="{x1},{y} {x1+9},{y-5} {x1+9},{y+5}" fill="{cor}"/>'
        return f'<line x1="{x1}" y1="{y}" x2="{x2}" y2="{y}" stroke="{cor}" stroke-width="2" stroke-dasharray="5 5" opacity=".55"/>{head}{dots}<text x="{(x1+x2)//2}" y="{y-10 if rot else y+22}" text-anchor="middle" font-family="{FONT}" font-size="12.5" fill="{cor}" font-weight="600">{label}</text>'

    svg = f'''<svg {NS} viewBox="0 0 1000 260" width="1000" height="260" role="img" aria-label="Arquitetura: navegador, servidor e planilha trocando dados">
<defs><linearGradient id="cg" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#2b1759"/><stop offset="1" stop-color="{INK}"/></linearGradient></defs>
{card(10, "Navegador", "sistema/index.html", ic_browser, "#8fe8ff")}
{card(385, "Servidor", "sistema/server.py", ic_server, "#ff5a60")}
{card(760, "Planilha", "acompanhamento_montagem_vN.xlsx", ic_xlsx, "#7CFF9B")}
{lane(250, 380, 100, True, "grava (POST)", "#ff9aa0", 2.6)}
{lane(250, 380, 150, False, "dados prontos (GET, 2 s)", "#8fe8ff", 2.6)}
{lane(625, 755, 100, True, "escreve nas células amarelas", "#ff9aa0", 2.6)}
{lane(625, 755, 150, False, "lê o que as fórmulas calcularam", "#7CFF9B", 2.6)}
<text x="500" y="240" text-anchor="middle" font-family="{FONT}" font-size="14" fill="#9aa3b8">Nenhuma conta é feita no painel: todas as contas ficam nas fórmulas da planilha.</text>
</svg>'''
    save("arquitetura.svg", svg)


# ================================================================== 5) RODAPÉ
def rodape():
    rng = random.Random(11)
    svg = f'''<svg {NS} viewBox="0 0 1200 130" width="1200" height="130" role="img" aria-label="Rodapé: skyline de San Fransokyo">
<defs><linearGradient id="fg" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#5b1f6b"/><stop offset="1" stop-color="#e0467c"/></linearGradient></defs>
<rect width="1200" height="130" rx="14" fill="url(#fg)"/>
{buildings(rng, -10, 1210, 130, 30, 80, "#3a1b5e", windows=False)}
{buildings(rng, -10, 1210, 130, 18, 56, "#1c0e38", windows=True)}
<text x="600" y="46" text-anchor="middle" font-family="{FONT}" font-size="19" font-weight="700" fill="{WHITE}">Seu cuidado com a montagem foi salvo na planilha.
<animate attributeName="opacity" values=".75;1;.75" dur="4s" repeatCount="indefinite"/></text>
<text x="600" y="68" text-anchor="middle" font-family="{FONT}" font-size="13" fill="#ffe3c2">Projeto não oficial · arte original inspirada em Operação Big Hero</text>
</svg>'''
    save("rodape.svg", svg)


banner(); divider(); baymax_diz(); arquitetura(); rodape()
