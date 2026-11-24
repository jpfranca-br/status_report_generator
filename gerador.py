"""Gera o status report one-page (.pptx) a partir da planilha modelo.

Uso pela linha de comando:
    python gerador.py projeto.xlsx                 -> projeto.pptx
    python gerador.py a.xlsx b.xlsx -o deck.pptx   -> um slide por planilha

O desenho é feito numa grade de 1920 x 1080 "px" (16:9) e convertido para
unidades do PowerPoint. Tudo sai como formas nativas: dá para ajustar à mão depois.
"""
from __future__ import annotations

import argparse
import math
import os
import shutil
import subprocess
import tempfile
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from pathlib import Path

from openpyxl import load_workbook
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

# ----------------------------------------------------------------------------
# Aparência
# ----------------------------------------------------------------------------
FONTE = "Segoe UI"          # troque por "Calibri" / "Arial" se preferir
CHAR = 0.53                 # largura média de um caractere, em fração do tamanho da fonte

COR = {
    "ink": "1F2A27", "muted": "5B6662", "brand": "0E3B34", "bg": "F6F7F4",
    "card": "FFFFFF", "line": "DCE2DE", "grid": "E3E8E5",
    "ok": "1E7A47", "ok_bg": "DFF0E5", "ok_ink": "155233",
    "warn": "E0A100", "warn_bg": "FCEFD2", "warn_ink": "6E4600",
    "bad": "C62828", "bad_bg": "FBE0DE", "bad_ink": "8E1A1A",
    "neutral_bg": "E6ECEA", "neutral_ink": "2F3B37",
    "plan": "9AA5A1", "dash": "7D8884", "diamond": "3E4A46",
    "yellow": "F2C94C", "note_bg": "FFF6D6", "note_ink": "3A2A00",
}
FAROL_COR = {"verde": ("ok", "ok_ink"), "amarelo": ("warn", "warn_ink"), "vermelho": ("bad", "bad_ink")}
MESES = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]
STATUS_VALIDOS = {"concluido", "realizado", "sem_sucesso", "critico", "planejado", "janela"}

W, H = 1920, 1080
PX = 6350  # EMU por px (13,333 pol / 1920)


# ----------------------------------------------------------------------------
# Dados
# ----------------------------------------------------------------------------
@dataclass
class Entrega:
    titulo: str
    inicio: date | None
    fim: date | None
    prazo: date | None
    status: str
    observacao: str = ""


@dataclass
class Relatorio:
    projeto: dict
    cronograma: list[Entrega]
    destravadores: list[tuple[str, str, str]]
    progresso: list[tuple[str, str]]
    passos: list[tuple[str, str]]
    ridac: list[tuple[str, str]]
    avisos: list[str] = field(default_factory=list)
    origem: str = ""


def _data(v, avisos, onde):
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = str(v).strip()
    for fmt in ("%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    avisos.append(f"{onde}: data inválida '{s}' (use dd/mm/aaaa)")
    return None


def _txt(v):
    if v is None:
        return ""
    if isinstance(v, (datetime, date)):
        return v.strftime("%d/%m")
    return str(v).strip()


def _linhas(ws, ncols):
    out = []
    for row in ws.iter_rows(min_row=2, max_col=ncols, values_only=True):
        if row and any(c not in (None, "") for c in row):
            out.append(list(row) + [None] * (ncols - len(row)))
    return out


def _aba(wb, *nomes):
    for n in nomes:
        for ws in wb.worksheets:
            if ws.title.strip().lower() == n.lower():
                return ws
    return None


def ler_planilha(arquivo, nome="") -> Relatorio:
    wb = load_workbook(arquivo, data_only=True)
    avisos: list[str] = []

    ws = _aba(wb, "Projeto")
    if ws is None:
        raise ValueError("A planilha precisa da aba 'Projeto'.")
    projeto = {}
    for campo, valor, *_ in _linhas(ws, 2):
        if campo:
            projeto[str(campo).strip().lower()] = valor
    for k in ("inicio", "fim", "data_referencia"):
        projeto[k] = _data(projeto.get(k), avisos, f"Projeto.{k}")

    cron = []
    ws = _aba(wb, "Cronograma")
    for i, (tit, ini, fim, prazo, st, obs) in enumerate(_linhas(ws, 6) if ws else [], start=2):
        if not tit:
            continue
        st = _txt(st).lower().replace(" ", "_").replace("í", "i").replace("ç", "c")
        if st not in STATUS_VALIDOS:
            avisos.append(f"Cronograma linha {i}: status '{st}' desconhecido; usei 'planejado'")
            st = "planejado"
        e = Entrega(_txt(tit), _data(ini, avisos, f"Cronograma linha {i}"),
                    _data(fim, avisos, f"Cronograma linha {i}"),
                    _data(prazo, avisos, f"Cronograma linha {i}"), st, _txt(obs))
        if e.fim is None and e.prazo is None:
            avisos.append(f"Cronograma linha {i} ('{e.titulo}'): sem fim nem prazo; ficou sem barra")
        cron.append(e)

    ws = _aba(wb, "Destravadores")
    destr = [(_txt(g), _txt(it), _txt(s)) for g, it, s in (_linhas(ws, 3) if ws else []) if it]
    ws = _aba(wb, "Progresso")
    prog = [(_txt(d), _txt(t)) for d, t in (_linhas(ws, 2) if ws else []) if t]
    ws = _aba(wb, "Proximos passos", "Próximos passos")
    passos = [(_txt(a), _txt(d)) for a, d in (_linhas(ws, 2) if ws else []) if a]
    ws = _aba(wb, "RIDAC")
    ridac = [(_txt(t), _txt(d)) for t, d in (_linhas(ws, 2) if ws else []) if d]

    return Relatorio(projeto, cron, destr, prog, passos, ridac, avisos, nome)


# ----------------------------------------------------------------------------
# Utilidades de desenho
# ----------------------------------------------------------------------------
def E(v):
    return Emu(int(round(v * PX)))


def rgb(k):
    return RGBColor.from_string(COR.get(k, k))


def largura_txt(texto, px, bold=False, folga=0.0):
    return len(texto) * px * (CHAR + folga + (0.03 if bold else 0))


def n_linhas(texto, largura, px, bold=False, folga=0.0):
    """Estimativa de quebras de linha (quebra só em espaços)."""
    if not texto:
        return 0
    cpl = max(1, int(largura / (px * (CHAR + folga + (0.03 if bold else 0)))))
    linhas, atual = 1, 0
    for palavra in texto.split():
        n = len(palavra)
        if atual == 0:
            atual = n
        elif atual + 1 + n <= cpl:
            atual += 1 + n
        else:
            linhas += 1
            atual = n
    return linhas


def fmt(d: date, ano=True):
    return d.strftime("%d/%m/%y") if ano else d.strftime("%d/%m")


class Tela:
    def __init__(self, slide):
        self.s = slide.shapes

    def caixa(self, x, y, w, h, fill=None, line=None, raio=0, forma=None, dash=False, lw=2):
        tipo = forma or (MSO_SHAPE.ROUNDED_RECTANGLE if raio else MSO_SHAPE.RECTANGLE)
        sh = self.s.add_shape(tipo, E(x), E(y), E(w), E(h))
        if raio and tipo == MSO_SHAPE.ROUNDED_RECTANGLE:
            sh.adjustments[0] = min(0.5, raio / max(1, min(w, h)))
        if fill:
            sh.fill.solid()
            sh.fill.fore_color.rgb = rgb(fill)
        else:
            sh.fill.background()
        if line:
            sh.line.color.rgb = rgb(line)
            sh.line.width = Pt(lw * 0.5)
            if dash:
                sh.line.dash_style = MSO_LINE_DASH_STYLE.DASH
        else:
            sh.line.fill.background()
        sh.shadow.inherit = False
        st = sh._element.find(qn("p:style"))
        if st is not None:  # sem estilo do tema: evita sombra e cor de texto herdadas
            sh._element.remove(st)
        sh.text_frame.text = ""
        return sh

    def texto(self, x, y, w, h, partes, px=24, cor="ink", bold=False, align="l",
              anchor="t", sh=None, linha=1.2, quebra=True):
        """partes: str ou lista de (texto, {cor, bold})."""
        if sh is None:
            sh = self.s.add_textbox(E(x), E(y), E(w), E(h))
        tf = sh.text_frame
        tf.word_wrap = quebra
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}[anchor]
        p = tf.paragraphs[0]
        p.alignment = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}[align]
        p.line_spacing = linha
        if isinstance(partes, str):
            partes = [(partes, {})]
        for t, o in partes:
            r = p.add_run()
            r.text = t
            f = r.font
            f.name = FONTE
            f.size = Pt(o.get("px", px) * 0.5)
            f.bold = o.get("bold", bold)
            f.color.rgb = rgb(o.get("cor", cor))
        return sh

    def linha(self, x1, y1, x2, y2, cor, lw=2, dash=False):
        c = self.s.add_connector(MSO_CONNECTOR.STRAIGHT, E(x1), E(y1), E(x2), E(y2))
        st = c._element.find(qn("p:style"))
        if st is not None:
            c._element.remove(st)
        c.line.color.rgb = rgb(cor)
        c.line.width = Pt(lw * 0.5)
        if dash:
            c.line.dash_style = MSO_LINE_DASH_STYLE.DASH
        return c

    def chip(self, x, y, w, texto, tipo, px=24):
        bg, ink = {"ok": ("ok_bg", "ok_ink"), "warn": ("warn_bg", "warn_ink"),
                   "bad": ("bad_bg", "bad_ink")}.get(tipo, ("neutral_bg", "neutral_ink"))
        h = px * 1.45
        sh = self.caixa(x, y, w, h, fill=bg, raio=h / 2)
        self.texto(0, 0, 0, 0, texto, px=px, cor=ink, bold=True, align="c", anchor="m", sh=sh,
                   quebra=False)
        return h

    def cartao(self, x, y, w, h, titulo, sub=""):
        self.caixa(x, y, w, h, fill="card", line="line", raio=14)
        partes = [(titulo, {"bold": True, "cor": "brand", "px": 28})]
        if sub:
            partes.append(("  · " + sub, {"cor": "muted", "px": 22}))
        self.texto(x + 24, y + 16, w - 48, 40, partes)


def tipo_status(s):
    s = s.lower()
    if s.startswith(("conclu", "realiz", "ok", "feito")):
        return "ok"
    if s.startswith(("bloq", "atras", "issue", "crít", "crit")):
        return "bad"
    return "warn"


# ----------------------------------------------------------------------------
# Blocos do slide
# ----------------------------------------------------------------------------
def cabecalho(t: Tela, p, avisos):
    t.texto(52, 30, 1150, 64, p.get("nome") or "Projeto", px=48, cor="brand", bold=True)
    desc = _txt(p.get("descricao"))
    if n_linhas(desc, 1150, 24, folga=-0.04) > 1:
        avisos.append("Descrição passa de uma linha; encurte para não encostar na mensagem-chave")
    t.texto(52, 98, 1150, 34, desc, px=24, cor="muted")

    farol = _txt(p.get("farol")).lower() or "verde"
    cor, ink = FAROL_COR.get(farol, FAROL_COR["verde"])
    rot = f"Farol {farol}"
    lw = largura_txt(rot, 32, True) + 8
    lx = W - 52 - lw
    t.caixa(lx - 54, 38, 40, 40, fill=cor, forma=MSO_SHAPE.OVAL)
    t.texto(lx, 30, lw, 52, rot, px=32, cor=ink, bold=True, anchor="m", quebra=False)

    dims = [("Escopo", "farol_escopo"), ("Tempo", "farol_tempo"), ("Custo", "farol_custo"),
            ("Recursos", "farol_recursos"), ("Qualidade", "farol_qualidade")]
    x = W - 52
    for rotulo, k in reversed(dims):
        v = _txt(p.get(k)).lower()
        if not v:
            continue
        lw = largura_txt(rotulo, 24) + 12
        x -= lw
        t.texto(x, 84, lw, 32, rotulo, px=24, cor="muted", anchor="m", quebra=False)
        x -= 22
        t.caixa(x, 92, 16, 16, fill=FAROL_COR.get(v, FAROL_COR["verde"])[0], forma=MSO_SHAPE.OVAL)
        x -= 18

    partes = []
    if p.get("inicio") and p.get("fim"):
        partes.append(f"{p['inicio']:%d/%m/%Y} → {p['fim']:%d/%m/%Y}")
    if p.get("classificacao"):
        partes.append(_txt(p["classificacao"]))
    t.texto(1100, 120, 768, 34, " · ".join(partes), px=24, cor="muted", align="r")


def mensagem(t: Tela, p, avisos):
    y, h = 162, 100
    t.caixa(52, y, 1816, h, fill="brand", raio=14)
    tag = t.caixa(80, y + 30, 250, 40, fill="yellow", raio=8)
    t.texto(0, 0, 0, 0, "Mensagem-chave", px=24, cor="brand", bold=True, align="c", anchor="m", sh=tag)
    d, c = _txt(p.get("mensagem_destaque")), _txt(p.get("mensagem_complemento"))
    partes = []
    if d:
        partes.append((d, {"bold": True, "cor": "yellow"}))
    if c:
        partes.append(((" " if d else "") + c, {"cor": "FFFFFF"}))
    if n_linhas(d + " " + c, 1480, 26) > 2:
        avisos.append("Mensagem-chave passa de duas linhas; encurte o texto")
    t.texto(356, y + 8, 1484, h - 16, partes, px=26, anchor="m", linha=1.15)


def _trimestre(d: date):
    return date(d.year, 3 * ((d.month - 1) // 3) + 1, 1)


def _prox_tri(d: date):
    m = d.month + 3
    return date(d.year + (m - 1) // 12, (m - 1) % 12 + 1, 1)


def cronograma(t: Tela, r: Relatorio, hoje: date, x0, y0, w, h, faixa_nota=False, avisos=None):
    """Desenha o cronograma. Retorna False se a nota não coube no gráfico."""
    p = r.projeto
    avisos = r.avisos if avisos is None else avisos
    t.cartao(x0, y0, w, h, "Macrocronograma", _txt(p.get("subtitulo_cronograma")))
    ents = r.cronograma
    if not ents:
        avisos.append("Cronograma vazio")
        return True
    if len(ents) > 8:
        avisos.append(f"Cronograma com {len(ents)} entregas; o slide comporta bem até 8")

    cx, cy, cw, ch = x0 + 26, y0 + 64, w - 52, h - 80
    LW = 330
    tx0, tx1 = cx + LW, cx + cw - 12

    datas = [d for e in ents for d in (e.inicio, e.fim, e.prazo) if d] + [hoje]
    if p.get("fim"):
        datas.append(p["fim"])
    ini, fim = _trimestre(min(datas)), _prox_tri(_trimestre(max(datas)))
    tris = []
    q = ini
    while q < fim:
        tris.append(q)
        q = _prox_tri(q)
    if len(tris) > 8:
        avisos.append(f"Cronograma cobre {len(tris)} trimestres; fica apertado (ideal até 6)")
    span = (fim - ini).days

    def X(d):
        return tx0 + (d - ini).days / span * (tx1 - tx0)

    qw = (tx1 - tx0) / len(tris)
    for i, q in enumerate(tris):
        fimq = _prox_tri(q) - timedelta(days=1)
        rot = f"{MESES[q.month - 1]}–{MESES[fimq.month - 1]} {q.year % 100:02d}"
        b = t.caixa(tx0 + i * qw + 1, cy, qw - 2, 30, fill="brand", raio=6)
        t.texto(0, 0, 0, 0, rot, px=17, cor="FFFFFF", bold=True, align="c", anchor="m", sh=b)

    top = cy + 36
    legenda_h = 30 + (40 if faixa_nota else 0)
    rh = min(36, (ch - 36 - legenda_h - 6) / len(ents))
    bottom = top + rh * len(ents)

    for i in range(len(tris) + 1):
        t.linha(tx0 + i * qw, top, tx0 + i * qw, bottom, "grid")
    for i in range(1, len(ents) + 1):
        t.linha(cx, top + i * rh, tx1, top + i * rh, "grid")
    if p.get("fim") and ini <= p["fim"] <= fim:
        t.linha(X(p["fim"]), top, X(p["fim"]), bottom, "brand", lw=2.5)

    ocupado = []  # retângulos do cronograma, para posicionar a nota
    usados = set()
    for i, e in enumerate(ents):
        ym = top + i * rh + rh / 2
        partes = [(e.titulo, {})]
        if e.prazo:
            partes.append((f" · {fmt(e.prazo)}", {}))
        if e.status == "sem_sucesso":
            partes.append((" · sem sucesso", {"cor": "bad"}))
        total = "".join(p[0] for p in partes)
        fpx = 17
        while fpx > 13 and largura_txt(total, fpx, True) > LW - 10:
            fpx -= 1
        if largura_txt(total, fpx, True) > LW - 10:
            avisos.append(f"Cronograma: título '{e.titulo}' é longo demais para a coluna; encurte")
        partes = [(tx_, dict(o, px=fpx)) for tx_, o in partes]
        t.texto(cx, ym - rh / 2, LW - 8, rh, partes, px=fpx, bold=True, anchor="m")

        ini_b = e.inicio or (hoje if e.status == "janela" else None)
        fim_b = e.fim or e.prazo
        fim_x = None
        if ini_b and fim_b and fim_b > ini_b:
            bx, bw = X(ini_b), X(fim_b) - X(ini_b)
            if e.status == "janela":
                t.caixa(bx, ym - 6, bw, 12, line="dash", raio=6, dash=True)
            else:
                cor = {"critico": "bad", "planejado": "plan"}.get(e.status, "ok")
                t.caixa(bx, ym - 6, bw, 12, fill=cor, raio=6)
            usados.add(e.status)
            ocupado.append((bx, ym - 10, bx + bw, ym + 10))
            fim_x = bx + bw
            if e.status in ("concluido", "sem_sucesso"):
                cor, sim = ("ok", "✓") if e.status == "concluido" else ("bad", "✕")
                c = t.caixa(fim_x - 11, ym - 11, 22, 22, fill=cor, forma=MSO_SHAPE.OVAL)
                t.texto(0, 0, 0, 0, sim, px=15, cor="FFFFFF", bold=True, align="c", anchor="m", sh=c)
                fim_x += 11
        elif fim_b and ini_b and fim_b <= ini_b:
            r.avisos.append(f"'{e.titulo}': fim antes do início; barra omitida")
        if e.prazo:
            dx = X(e.prazo)
            t.caixa(dx - 10, ym - 10, 20, 20, fill="bad_ink" if e.status == "critico" else "diamond",
                    forma=MSO_SHAPE.DIAMOND)
            usados.add("prazo")
            fim_x = max(fim_x or 0, dx + 10)
            ocupado.append((dx - 10, ym - 10, dx + 10, ym + 10))
        if e.observacao and fim_x:
            ow = largura_txt(e.observacao, 16, True) + 6
            ox = fim_x + 8
            if ox + ow > cx + cw:
                ox = (X(ini_b) if ini_b else fim_x) - ow - 8
            t.texto(ox, ym - 12, ow, 24, e.observacao, px=16, bold=True,
                    cor="bad" if e.status == "critico" else "muted", anchor="m", quebra=False)
            ocupado.append((ox, ym - 12, ox + ow, ym + 12))

    hx = X(hoje)
    t.linha(hx, top, hx, bottom, "brand", lw=2.5, dash=True)
    ocupado.append((hx - 4, top, hx + 4, bottom))

    # legenda
    ly = cy + ch - 20
    itens = []
    if usados & {"concluido", "realizado", "sem_sucesso"}:
        itens.append(("bar", "ok", "Realizado"))
    if "critico" in usados:
        itens.append(("bar", "bad", "Caminho crítico"))
    if "planejado" in usados:
        itens.append(("bar", "plan", "Planejado"))
    if "janela" in usados:
        itens.append(("dash", None, "Janela (início a definir)"))
    if "prazo" in usados:
        itens.append(("dia", None, "Prazo"))
    itens.append(("hoje", None, f"Hoje ({fmt(hoje, False)})"))
    lx = cx
    for kind, cor, rot in itens:
        if kind == "bar":
            t.caixa(lx, ly - 5, 28, 10, fill=cor, raio=5)
            lx += 36
        elif kind == "dash":
            t.caixa(lx, ly - 5, 28, 10, line="dash", raio=5, dash=True)
            lx += 36
        elif kind == "dia":
            t.caixa(lx, ly - 7, 14, 14, fill="diamond", forma=MSO_SHAPE.DIAMOND)
            lx += 22
        else:
            t.linha(lx, ly, lx + 28, ly, "brand", lw=2.5, dash=True)
            lx += 36
        tw = largura_txt(rot, 16) + 6
        t.texto(lx, ly - 12, tw, 24, rot, px=16, cor="muted", anchor="m", quebra=False)
        lx += tw + 20

    nota = _txt(p.get("nota_cronograma"))
    if not nota:
        return True
    if faixa_nota:
        _faixa(t, nota, cx, bottom + 8, tx1 - cx, avisos)
        return True
    return _nota(t, nota, ocupado, tx0, top, tx1, bottom)


def _partes_nota(nota, px):
    if ":" in nota:
        a, b = nota.split(":", 1)
        return [(a + ":", {"bold": True, "px": px}), (b, {"px": px})]
    return [(nota, {"px": px})]


def _faixa(t: Tela, nota, x, y, w, avisos):
    """Nota numa faixa de uma linha abaixo das barras (quando não há espaço livre no gráfico)."""
    px = 20
    while px > 15 and largura_txt(nota, px, folga=0.02) > w - 30:
        px -= 1
    if largura_txt(nota, px) > w - 30:
        avisos.append("Nota do cronograma longa demais para uma linha; encurte")
    t.caixa(x, y, w, 32, fill="note_bg", raio=6)
    t.caixa(x, y, 5, 32, fill="warn")
    t.texto(x + 16, y, w - 24, 32, _partes_nota(nota, px), px=px, cor="note_ink", anchor="m")


def _nota(t: Tela, nota, ocupado, x0, y0, x1, y1):
    def livre(r):
        return all(r[2] + 6 < o[0] or r[0] > o[2] + 6 or r[3] + 4 < o[1] or r[1] > o[3] + 4
                   for o in ocupado)

    for bw in (400, 440, 520, 620):
        bh = n_linhas(nota, bw - 28, 24, folga=0.05) * 30 + 18
        if bh > y1 - y0:
            continue
        melhor = None
        y = y1 - bh - 2
        while y >= y0 and melhor is None:
            x = x0 + 4
            while x + bw <= x1:
                if livre((x, y, x + bw, y + bh)):
                    melhor = (x, y)
                    break
                x += 8
            y -= 4
        if melhor:
            x, y = melhor
            t.caixa(x, y, bw, bh, fill="note_bg", raio=8)
            t.caixa(x, y, 5, bh, fill="warn")
            t.texto(x + 16, y + 8, bw - 28, bh - 16, _partes_nota(nota, 24), px=24, cor="note_ink",
                    linha=1.1)
            return True
    return False


def destravadores(t: Tela, r: Relatorio, x0, y0, w, h):
    t.cartao(x0, y0, w, h, _txt(r.projeto.get("titulo_destravadores")) or "O que destrava o projeto")
    y = y0 + 66
    grupo_atual = None
    cw_chip = max([180] + [largura_txt(s or "Em aberto", 24, True) + 36 for _, _, s in r.destravadores])
    tw = w - 48 - cw_chip - 16
    for g, item, st in r.destravadores:
        if g and g != grupo_atual:
            if y + 34 > y0 + h - 12:
                break
            t.texto(x0 + 24, y, w - 48, 32, g, px=24, cor="ok", bold=True)
            y += 38
            grupo_atual = g
        nl = n_linhas(item, tw, 24, True, folga=0.04)
        rh = max(nl * 30, 35)
        if y + rh > y0 + h - 12:
            r.avisos.append("Destravadores: itens demais; os últimos ficaram de fora")
            return
        t.texto(x0 + 24, y, tw, rh, item, px=24, bold=True, anchor="m")
        t.chip(x0 + w - 24 - cw_chip, y + (rh - 35) / 2, cw_chip, st or "Em aberto", tipo_status(st or ""))
        y += rh + 6


def progresso(t: Tela, r: Relatorio, x0, y0, w, h):
    t.cartao(x0, y0, w, h, "Progresso recente")
    y = y0 + 68
    for d, txt in r.progresso:
        nl = n_linhas(txt, w - 184, 24, folga=0.04)
        rh = nl * 30
        if y + rh > y0 + h - 12:
            r.avisos.append("Progresso: itens demais; os últimos ficaram de fora")
            return
        eh_data = any(ch.isdigit() for ch in d)
        t.texto(x0 + 24, y, 136, 30, d, px=24, bold=True, cor="ok" if eh_data else "muted")
        t.texto(x0 + 160, y, w - 184, rh, txt, px=24)
        y += rh + 14


def passos(t: Tela, r: Relatorio, x0, y0, w, h):
    t.cartao(x0, y0, w, h, "Próximos passos")
    y = y0 + 68
    tw = w - 24 - 74
    for i, (a, d) in enumerate(r.passos, start=1):
        n1, n2 = n_linhas(a, tw, 24, True, folga=0.04), n_linhas(d, tw, 24, folga=0.04)
        rh = (n1 + n2) * 30
        if y + rh > y0 + h - 12:
            r.avisos.append("Próximos passos: itens demais; os últimos ficaram de fora")
            return
        c = t.caixa(x0 + 24, y, 36, 36, fill="brand", forma=MSO_SHAPE.OVAL)
        t.texto(0, 0, 0, 0, str(i), px=22, cor="FFFFFF", bold=True, align="c", anchor="m", sh=c)
        t.texto(x0 + 74, y, tw, n1 * 30, a, px=24, bold=True)
        if d:
            t.texto(x0 + 74, y + n1 * 30, tw, n2 * 30, d, px=24, cor="muted")
        y += rh + 12


def pontos(t: Tela, r: Relatorio, x0, y0, w, h):
    t.cartao(x0, y0, w, h, "Pontos de atenção", "RIDAC")
    y = y0 + 68
    cw_chip = max([100] + [largura_txt(tp, 24, True) + 32 for tp, _ in r.ridac])
    tx = x0 + 24 + cw_chip + 12
    tw = x0 + w - 24 - tx
    ordem = {"issue": 0, "risco": 1}
    for tipo, desc in sorted(r.ridac, key=lambda x: ordem.get(x[0].lower(), 2)):
        nl = n_linhas(desc, tw, 24, folga=0.04)
        rh = max(nl * 30, 35)
        if y + rh > y0 + h - 12:
            r.avisos.append("RIDAC: itens demais; os últimos ficaram de fora")
            return
        k = {"issue": "bad", "risco": "warn"}.get(tipo.lower(), "neutral")
        t.chip(x0 + 24, y, cw_chip, tipo, k)
        t.texto(tx, y + 2, tw, rh, desc, px=24)
        y += rh + 10


# ----------------------------------------------------------------------------
# Montagem
# ----------------------------------------------------------------------------
def desenhar(prs: Presentation, r: Relatorio):
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # em branco
    bg = slide.background.fill
    bg.solid()
    bg.fore_color.rgb = rgb("bg")
    t = Tela(slide)
    hoje = r.projeto.get("data_referencia") or date.today()

    cabecalho(t, r.projeto, r.avisos)
    mensagem(t, r.projeto, r.avisos)
    # ensaio num slide descartável: a nota cabe solta no gráfico ou vai para a faixa?
    rascunho = Presentation()
    coube = cronograma(Tela(rascunho.slides.add_slide(rascunho.slide_layouts[6])), r, hoje,
                       52, 274, 1060, 392, avisos=[])
    cronograma(t, r, hoje, 52, 274, 1060, 392, faixa_nota=not coube)
    destravadores(t, r, 1124, 274, 744, 392)
    yb, hb = 678, H - 36 - 678
    cw = (1816 - 2 * 17) / 3
    progresso(t, r, 52, yb, cw, hb)
    passos(t, r, 52 + cw + 17, yb, cw, hb)
    pontos(t, r, 52 + 2 * (cw + 17), yb, cw, hb)

    notas = [f"Gerado em {date.today():%d/%m/%Y} a partir de {r.origem or 'planilha'}."]
    slide.notes_slide.notes_text_frame.text = "\n".join(notas)
    return slide


def gerar(relatorios: list[Relatorio], saida) -> list[str]:
    prs = Presentation()
    prs.slide_width, prs.slide_height = E(W), E(H)
    avisos = []
    for r in relatorios:
        desenhar(prs, r)
        avisos += [(f"[{r.origem}] " if len(relatorios) > 1 else "") + a for a in r.avisos]
    prs.save(saida)
    return avisos


# ----------------------------------------------------------------------------
# Prévia (opcional: precisa do LibreOffice instalado)
# ----------------------------------------------------------------------------
def achar_soffice():
    for c in ("soffice", "libreoffice"):
        if shutil.which(c):
            return shutil.which(c)
    for c in (r"C:\Program Files\LibreOffice\program\soffice.exe",
              r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
              "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        if os.path.exists(c):
            return c
    return None


def previa_png(pptx_path) -> list[bytes] | None:
    """Converte o .pptx em PNGs (um por slide). Retorna None se não houver LibreOffice."""
    soffice = achar_soffice()
    if not soffice:
        return None
    try:
        import pymupdf as fitz  # PyMuPDF
    except ImportError:
        return None
    with tempfile.TemporaryDirectory() as tmp:
        subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", tmp, str(pptx_path)],
                       check=True, capture_output=True, timeout=120)
        pdf = Path(tmp) / (Path(pptx_path).stem + ".pdf")
        doc = fitz.open(pdf)
        return [pg.get_pixmap(dpi=110).tobytes("png") for pg in doc]


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="Gera o status report one-page a partir de XLSX.")
    ap.add_argument("planilhas", nargs="+")
    ap.add_argument("-o", "--saida")
    a = ap.parse_args()
    rels = [ler_planilha(p, Path(p).name) for p in a.planilhas]
    saida = a.saida or str(Path(a.planilhas[0]).with_suffix(".pptx"))
    for aviso in gerar(rels, saida):
        print("AVISO:", aviso)
    print("Gerado:", saida)
