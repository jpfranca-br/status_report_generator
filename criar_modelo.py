"""Cria a planilha modelo do status report (one-page), preenchida com o
Projeto Aurora (dados fictícios) como exemplo.

Uso:  python criar_modelo.py            -> gera modelo_status_report.xlsx
"""
from datetime import date

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

HEADER_FILL = PatternFill("solid", fgColor="0E3B34")
HEADER_FONT = Font(bold=True, color="FFFFFF")

STATUS_CRONO = ["concluido", "realizado", "sem_sucesso", "critico", "planejado", "janela"]
STATUS_ITEM = ["Concluído", "Em andamento", "Em aberto", "Em definição", "Em teste", "Pendente", "Bloqueado"]
TIPOS_RIDAC = ["Risco", "Issue", "Dependência", "Ação", "Decisão"]
FAROL = ["verde", "amarelo", "vermelho"]


def _sheet(wb, title, headers, rows, widths):
    ws = wb.create_sheet(title)
    ws.append(headers)
    for c in ws[1]:
        c.fill, c.font = HEADER_FILL, HEADER_FONT
    for r in rows:
        ws.append(r)
    for i, w in enumerate(widths):
        ws.column_dimensions[chr(65 + i)].width = w
    ws.freeze_panes = "A2"
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
    return ws


def _lista(ws, col, valores, ate=60):
    dv = DataValidation(type="list", formula1='"' + ",".join(valores) + '"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"{col}2:{col}{ate}")


def criar(caminho="modelo_status_report.xlsx"):
    wb = Workbook()
    wb.remove(wb.active)

    ws = _sheet(wb, "Projeto", ["campo", "valor", "observação"], [
        ["nome", "Projeto Aurora", "Título do slide"],
        ["descricao", "Nova carteira digital para frotas: pagamento por QR Code e limite de crédito por motorista", "Uma linha"],
        ["inicio", date(2026, 2, 2), "Início do projeto"],
        ["fim", date(2027, 6, 30), "Fim do projeto (linha cheia no cronograma)"],
        ["classificacao", "Exemplo · fictício", "Aparece ao lado das datas"],
        ["data_referencia", None, "Data da linha 'Hoje'. Vazio = data de geração"],
        ["farol", "amarelo", "verde / amarelo / vermelho"],
        ["farol_escopo", "verde", ""],
        ["farol_tempo", "amarelo", ""],
        ["farol_custo", "verde", ""],
        ["farol_recursos", "amarelo", ""],
        ["farol_qualidade", "verde", ""],
        ["mensagem_destaque", "MVP em homologação com dois clientes-piloto; go-live depende do aceite do jurídico.", "Parte em destaque (amarelo)"],
        ["mensagem_complemento", "Crédito por motorista segue para a fase 2 em 2027.", "Continuação em branco"],
        ["subtitulo_cronograma", "entregas e prazos", ""],
        ["nota_cronograma", "Out/26: piloto estendido em 3 semanas a pedido do cliente Alfa Logística.", "Opcional: caixa amarela no cronograma"],
        ["titulo_destravadores", "O que destrava o go-live", "Título do bloco à direita"],
    ], [26, 90, 46])
    for r in range(2, ws.max_row + 1):
        if isinstance(ws.cell(r, 2).value, date):
            ws.cell(r, 2).number_format = "DD/MM/YYYY"

    ws = _sheet(wb, "Cronograma", ["titulo", "inicio", "fim", "prazo", "status", "observacao"], [
        ["Descoberta e desenho", date(2026, 2, 2), date(2026, 4, 15), None, "concluido", None],
        ["Integração com gateway", date(2026, 4, 1), date(2026, 7, 31), None, "realizado", None],
        ["Prova de conceito NFC", date(2026, 5, 1), date(2026, 8, 15), None, "sem_sucesso", "trocado por QR"],
        ["MVP pagamento por QR Code", date(2026, 8, 1), date(2026, 12, 18), date(2026, 12, 18), "critico", "piloto"],
        ["Portal do gestor de frota", date(2026, 10, 1), date(2027, 2, 26), date(2027, 2, 26), "planejado", None],
        ["Antifraude transacional", None, date(2027, 3, 31), date(2027, 3, 31), "janela", None],
        ["Crédito por motorista", None, date(2027, 6, 30), date(2027, 6, 30), "janela", None],
    ], [34, 13, 13, 13, 14, 24])
    for row in ws.iter_rows(min_row=2, min_col=2, max_col=4):
        for c in row:
            c.number_format = "DD/MM/YYYY"
    _lista(ws, "E", STATUS_CRONO)

    ws = _sheet(wb, "Destravadores", ["grupo", "item", "status"], [
        ["Jurídico e compliance", "Aceite dos termos de uso", "Em andamento"],
        ["Jurídico e compliance", "Parecer sobre crédito a PF", "Em aberto"],
        ["Clientes-piloto", "Cadastro de 40 veículos (Alfa)", "Concluído"],
        ["Clientes-piloto", "Treinamento dos gestores (Beta)", "Bloqueado"],
        ["Clientes-piloto", "Contrato do piloto Beta", "Em definição"],
    ], [48, 40, 16])
    _lista(ws, "C", STATUS_ITEM)

    _sheet(wb, "Progresso", ["data", "texto"], [
        ["03/11", "Homologação do QR Code concluída"],
        ["10/11", "Primeiras 120 transações no piloto"],
        ["17/11", "Ajuste de limite diário liberado"],
        ["Contínuo", "Comitê semanal com clientes-piloto"],
    ], [12, 50])

    _sheet(wb, "Proximos passos", ["acao", "detalhe"], [
        ["Obter aceite do jurídico", "Jurídico · até 05/12"],
        ["Treinar gestores do cliente Beta", "Operações · até 12/12"],
        ["Go/no-go do MVP", "Comitê executivo · 18/12"],
    ], [40, 44])

    ws = _sheet(wb, "RIDAC", ["tipo", "descricao"], [
        ["Risco", "Aceite jurídico atrasar o go-live de dezembro"],
        ["Issue", "Leitura de QR instável em Android antigo"],
        ["Dependência", "Equipe de dados para regras antifraude"],
        ["Decisão", "NFC descartado; seguir só com QR Code"],
    ], [12, 60])
    _lista(ws, "A", TIPOS_RIDAC)

    ws = wb.create_sheet("Instruções", 0)
    linhas = [
        "COMO PREENCHER",
        "",
        "Projeto: uma linha por campo. Não mude a coluna 'campo'.",
        "Cronograma: uma linha por entrega, na ordem em que devem aparecer.",
        "  inicio: pode ficar vazio quando status = janela (a barra começa em 'Hoje').",
        "  prazo: opcional. Desenha um losango na data.",
        "  status:",
        "    concluido   – barra verde com ✓ no fim",
        "    realizado   – barra verde",
        "    sem_sucesso – barra verde com ✕ vermelho no fim e 'sem sucesso' no rótulo",
        "    critico     – barra vermelha (caminho crítico)",
        "    planejado   – barra cinza (início definido)",
        "    janela      – barra tracejada até o prazo (início a definir)",
        "  observacao: texto curto ao lado da barra (ex.: pausada).",
        "Destravadores: linhas com o mesmo 'grupo' ficam juntas. Status 'Bloqueado' = vermelho,",
        "  'Concluído' = verde, os demais = amarelo.",
        "Progresso, Próximos passos e RIDAC: uma linha por item. Textos curtos (até ~45 caracteres).",
        "",
        "Limites que cabem no slide: 8 entregas, 7 linhas em Destravadores (contando grupos),",
        "5 itens em Progresso, 4 em Próximos passos e 5 no RIDAC. O gerador avisa se passar.",
    ]
    for l in linhas:
        ws.append([l])
    ws["A1"].font = Font(bold=True, size=14, color="0E3B34")
    ws.column_dimensions["A"].width = 100

    wb.save(caminho)
    return caminho


if __name__ == "__main__":
    print("Criado:", criar())
