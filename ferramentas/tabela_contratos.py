"""
Tabela de contratos: não concorrência e vencimento.

Como rodar (no PowerShell, dentro da pasta Mike):
    python ferramentas\\tabela_contratos.py
    python ferramentas\\tabela_contratos.py "Outra pasta" "saida.xlsx"

Lê todos os PDF e Word da pasta e gera uma planilha com:
  - aba "Não concorrência": trecho, até quando vale, fonte, situação;
  - aba "Vencimento": trecho, data de vencimento, fonte, situação;
  - aba "Resumo": números e lista do que precisa de revisão.

Regras: o trecho é copiado como está; nada é estimado. Na dúvida, "Revisar".
"""
import re
import sys
import unicodedata
from datetime import date
from pathlib import Path

import docx
import openpyxl
import pdfplumber
from openpyxl.styles import Alignment, Font, PatternFill

PASTA_PADRAO = "Contratos de fornecedores"
SAIDA_PADRAO = "contratos.xlsx"

CONFERIDO, REVISAR, SEM_LEITURA = "Conferido", "Revisar", "Sem leitura"

MESES = {"janeiro": 1, "fevereiro": 2, "marco": 3, "abril": 4, "maio": 5, "junho": 6,
         "julho": 7, "agosto": 8, "setembro": 9, "outubro": 10, "novembro": 11, "dezembro": 12}

NUMEROS = {"um": 1, "uma": 1, "primeiro": 1, "dois": 2, "duas": 2, "tres": 3, "quatro": 4,
           "cinco": 5, "seis": 6, "sete": 7, "oito": 8, "nove": 9, "dez": 10, "onze": 11,
           "doze": 12, "treze": 13, "quatorze": 14, "catorze": 14, "quinze": 15,
           "dezesseis": 16, "dezessete": 17, "dezoito": 18, "dezenove": 19, "vinte": 20,
           "trinta": 30, "quarenta": 40, "cinquenta": 50, "sessenta": 60, "setenta": 70,
           "oitenta": 80, "noventa": 90, "cem": 100, "mil": 1000}

# Termos procurados pelo sentido (comparação sem acentos e sem maiúsculas)
TITULOS_NAO_CONCORRENCIA = ["nao concorrencia", "nao competicao", "exclusividade",
                            "vedacao de atuacao", "restricao comercial"]
TEXTOS_NAO_CONCORRENCIA = [r"concorrente", r"nao competir", r"nao concorrer",
                           r"nao competicao", r"nao concorrencia",
                           r"vedad[ao].{0,40}atuacao", r"vedacao de atuacao",
                           r"abster-?se(?:-a)?\s+de\s+(?:fornecer|prestar|atuar|contratar)",
                           r"abstendo-se", r"exclusiv", r"mesmo ramo",
                           r"impedid[ao] de (contratar|fornecer|atuar)"]
# Restrições parecidas que NÃO são tratadas como não concorrência, mas são avisadas
TITULOS_PARECIDOS = ["aliciamento", "nao solicitacao"]
# Palavras que indicam que outro item muda o prazo de uma cláusula
RE_ALTERACAO = r"prorrog|alterad|modificad|substituid|estendid|passa a vigorar"
TEXTOS_INDEFINIDO = [r"prazo indeterminado", r"tempo indeterminado",
                     r"enquanto (?:durar|perdurar|vigorar)[^.,;]*", r"relacao comercial"]


def sem_acento(texto):
    t = unicodedata.normalize("NFD", texto)
    return "".join(c for c in t if unicodedata.category(c) != "Mn").lower()


# ---------------------------------------------------------------- leitura
def ler_pdf(arquivo):
    """Devolve (itens, paginas_sem_texto, total_paginas)."""
    linhas, sem_texto = [], []
    with pdfplumber.open(arquivo) as pdf:
        total = len(pdf.pages)
        for n, pagina in enumerate(pdf.pages, 1):
            texto = (pagina.extract_text() or "").strip()
            if not texto:
                sem_texto.append(n)
                continue
            for linha in texto.splitlines():
                linhas.append((linha.strip(), n))
    return montar_itens(linhas, "pdf"), sem_texto, total


def ler_word(arquivo):
    d = docx.Document(arquivo)
    linhas = [(p.text.strip(), None) for p in d.paragraphs if p.text.strip()]
    for tabela in d.tables:
        for linha in tabela.rows:
            for celula in linha.cells:
                if celula.text.strip():
                    linhas.append((celula.text.strip(), None))
    return montar_itens(linhas, "word"), [], None


RE_TITULO = re.compile(r"^CL[ÁA]USULA\s+(\d+)[ªº.]?\s*[.\-–:]?\s*(.*)$", re.I)
RE_ITEM = re.compile(r"^(\d+\.\d+(?:\.\d+)*)\.?\s")


def montar_itens(linhas, tipo):
    """Junta as linhas em itens de cláusula (ex.: 7.1), guardando título e página."""
    itens, atual = [], None
    clausula, titulo = None, ""
    for texto, pagina in linhas:
        m = RE_TITULO.match(texto)
        if m:
            clausula, titulo = m.group(1), m.group(2).strip()
            atual = None
            itens.append({"tipo": "titulo", "clausula": clausula, "titulo": titulo,
                          "pagina": pagina})
            continue
        m = RE_ITEM.match(texto)
        if m:
            atual = {"tipo": "item", "numero": m.group(1), "clausula": clausula,
                     "titulo": titulo, "texto": texto, "paginas": [pagina]}
            itens.append(atual)
        elif atual is not None:
            atual["texto"] += ("\n" if tipo == "pdf" else " ") + texto
            if pagina not in atual["paginas"]:
                atual["paginas"].append(pagina)
        else:
            itens.append({"tipo": "solto", "texto": texto, "paginas": [pagina]})
    for i in itens:
        if i["tipo"] == "item" and tipo == "pdf":
            # junta as quebras de linha do PDF sem alterar as palavras
            i["texto"] = re.sub(r"-\n", "-", i["texto"]).replace("\n", " ")
    return itens


def fonte(item, tipo):
    if tipo == "pdf":
        pags = item["paginas"]
        p = f"página {pags[0]}" if len(pags) == 1 else f"páginas {pags[0]} a {pags[-1]}"
        return f"{p}, cláusula {item['numero']}"
    return f"cláusula {item['numero']}"


# ---------------------------------------------------------------- datas
def numero_por_extenso(palavras):
    total, atual = 0, 0
    for p in palavras:
        if p == "e":
            continue
        if p not in NUMEROS:
            return None
        v = NUMEROS[p]
        if v == 1000:
            total += (atual or 1) * 1000
            atual = 0
        else:
            atual += v
    return total + atual


def achar_datas(texto):
    """Lista de (data, trecho) encontrados no texto."""
    t = sem_acento(texto)
    achadas = []
    for m in re.finditer(r"\b(\d{1,2})/(\d{1,2})/(\d{4})\b", t):
        achadas.append((m.start(), m.group(0), int(m.group(3)), int(m.group(2)), int(m.group(1))))
    for m in re.finditer(r"\b(\d{1,2})[ºo]?\s+de\s+([a-z]+)\s+de\s+(\d{4})\b", t):
        if m.group(2) in MESES:
            achadas.append((m.start(), m.group(0), int(m.group(3)), MESES[m.group(2)], int(m.group(1))))
    nums = "|".join(sorted(NUMEROS, key=len, reverse=True))
    seq = rf"((?:{nums})(?:(?:\s+e\s+|\s+)(?:{nums}))*)"
    for m in re.finditer(rf"\b{seq}\s+de\s+([a-z]+)\s+de\s+{seq}\b", t):
        dia = numero_por_extenso(m.group(1).split())
        ano = numero_por_extenso(m.group(3).split())
        if m.group(2) in MESES and dia and ano:
            achadas.append((m.start(), m.group(0).strip(), ano, MESES[m.group(2)], dia))
    resultado = []
    for _, trecho, a, me, d in sorted(achadas):
        try:
            resultado.append((date(a, me, d), trecho))
        except ValueError:
            pass
    return resultado


def achar_prazo_relativo(texto):
    """Ex.: 'pelo prazo de 24 (vinte e quatro) meses após o término' -> (24, 'meses')."""
    t = sem_acento(texto)
    m = re.search(r"(\d+)\s*(?:\([^)]*\))?\s*(meses|mes|anos|ano)\s+(?:apos|contados? d[oa])\s+"
                  r"(?:o\s+)?(termino|fim|encerramento|extincao|rescisao)", t)
    if m:
        return int(m.group(1)), "meses" if m.group(2).startswith("mes") else "anos"
    return None


def somar(data, qtd, unidade):
    meses = qtd if unidade == "meses" else qtd * 12
    a, me = divmod(data.month - 1 + meses, 12)
    ano, mes = data.year + a, me + 1
    import calendar
    dia = min(data.day, calendar.monthrange(ano, mes)[1])
    return date(ano, mes, dia)


def fmt(d):
    return d.strftime("%d/%m/%Y")


# ---------------------------------------------------------------- análise
def eh_nao_concorrencia(item):
    titulo = sem_acento(item.get("titulo") or "")
    texto = sem_acento(item["texto"])
    if "confidencial" in titulo or "sigilo" in titulo:
        return False
    if any(t in titulo for t in TITULOS_NAO_CONCORRENCIA):
        return True
    return any(re.search(p, texto) for p in TEXTOS_NAO_CONCORRENCIA)


def eh_vigencia(item):
    titulo = sem_acento(item.get("titulo") or "")
    texto = sem_acento(item["texto"])
    return ("vigencia" in titulo or "prazo" in titulo or "vencimento" in titulo
            or re.search(r"\bvigorara\b|\bvigorar\b|prazo de vigencia|data de vencimento", texto))


def anexo_citado(texto, pasta):
    m = re.search(r"\b(anexo|ap[eê]ndice|aditivo)\s+([ivxlc\d]+|[a-z])\b", texto, re.I)
    if not m:
        return None
    nome = f"{m.group(1)} {m.group(2)}"
    alvo = sem_acento(nome)
    for f in pasta.iterdir():
        if alvo in sem_acento(f.stem):
            return None  # o anexo está na pasta
    return nome


def alteracoes(itens, alvo):
    """Itens que citam a cláusula-alvo (ex.: 'cláusula 7.1') e mudam o seu prazo."""
    nums = {alvo["numero"], alvo["clausula"]} - {None}
    achados = []
    for i in itens:
        if i["tipo"] != "item" or i is alvo:
            continue
        t = sem_acento(i["texto"])
        cita = any(re.search(rf"clausulas?\s+{re.escape(n)}(?![\d.]*\d)", t) for n in nums)
        if cita and re.search(RE_ALTERACAO, t):
            achados.append(i)
    return achados


def aplicar_alteracoes(linha, itens, alvo, tipo):
    """Se outro item altera o prazo, mostra as duas datas e marca para revisar."""
    for alt in alteracoes(itens, alvo):
        novas = achar_datas(alt["texto"])
        linha["trecho"] += "\n\n" + alt["texto"]
        linha["fonte"] += "; " + fonte(alt, tipo)
        original = linha["data"]
        if novas:
            linha["data"] = (f"{'; '.join(fmt(d) for d, _ in novas)} (alterada pela cláusula "
                             f"{alt['numero']}; texto original: {original})")
        else:
            linha["data"] = f"{original} (pode ter sido alterada pela cláusula {alt['numero']})"
        linha["situacao"] = REVISAR
        linha.pop("data_obj", None)
        linha["obs"].append(f"A cláusula {alt['numero']} ({alt['titulo']}) altera o prazo "
                            f"da cláusula {alvo['numero']}; conferir qual data vale.")
    return linha


def parecidos(itens, tipo):
    """Cláusulas de restrição parecida (ex.: não aliciamento) para avisar o advogado."""
    return [i for i in itens if i["tipo"] == "item"
            and any(t in sem_acento(i.get("titulo") or "") for t in TITULOS_PARECIDOS)]


def indefinido(texto):
    t = sem_acento(texto)
    for p in TEXTOS_INDEFINIDO:
        m = re.search(p, t)
        if m:
            # devolve o trecho como está no contrato (com acentos)
            if len(t) == len(texto):
                return texto[m.start():m.end()]
            return m.group(0)
    return None


def analisar_vencimento(itens, tipo, pasta):
    cand = [i for i in itens if i["tipo"] == "item" and eh_vigencia(i)]
    if not cand:
        return None
    i = cand[0]
    linha = {"trecho": i["texto"], "fonte": fonte(i, tipo), "data": None, "obs": []}
    if len(cand) > 1:
        linha["obs"].append("Há outros itens sobre vigência: "
                            + ", ".join(c["numero"] for c in cand[1:]) + ".")
    datas = achar_datas(i["texto"])
    ind = indefinido(i["texto"])
    anexo = anexo_citado(i["texto"], pasta)
    if anexo:
        linha.update(data="depende de documento ausente", situacao=REVISAR)
        linha["obs"].append(f"O texto remete a \"{anexo}\", que não está na pasta.")
    elif ind and not datas:
        linha.update(data="sem data certa", situacao=REVISAR)
        linha["obs"].append(f"O contrato diz \"{ind}\".")
    elif len({d for d, _ in datas}) == 1:
        linha.update(data=fmt(datas[0][0]), data_obj=datas[0][0], situacao=CONFERIDO)
    elif datas:
        linha.update(data="; ".join(fmt(d) for d, _ in datas), situacao=REVISAR)
        linha["obs"].append("Mais de uma data no mesmo item; conferir qual é o vencimento.")
    else:
        linha.update(data="sem data certa", situacao=REVISAR)
        linha["obs"].append("O item de vigência não traz data.")
    return aplicar_alteracoes(linha, itens, i, tipo)


def analisar_nao_concorrencia(itens, tipo, pasta, venc):
    cand = [i for i in itens if i["tipo"] == "item" and eh_nao_concorrencia(i)]
    if not cand:
        return None
    i = cand[0]
    linha = {"trecho": i["texto"], "fonte": fonte(i, tipo), "data": None, "obs": []}
    if len(cand) > 1:
        linha["trecho"] = "\n\n".join(c["texto"] for c in cand)
        linha["fonte"] = "; ".join(fonte(c, tipo) for c in cand)
        linha["obs"].append("Mais de um item com restrição; conferir se tratam do mesmo assunto.")
    texto = " ".join(c["texto"] for c in cand)
    datas = achar_datas(texto)
    rel = achar_prazo_relativo(texto)
    ind = indefinido(texto)
    anexo = anexo_citado(texto, pasta)
    data_venc = venc.get("data_obj") if venc else None

    if anexo:
        linha.update(data="depende de documento ausente", situacao=REVISAR)
        linha["obs"].append(f"O prazo/alcance está no \"{anexo}\", que não está na pasta.")
    elif rel:
        qtd, un = rel
        if data_venc:
            calc = somar(data_venc, qtd, un)
            conta = f"vencimento {fmt(data_venc)} + {qtd} {un} = {fmt(calc)}"
            explicitas = {d for d, _ in datas}
            if not explicitas or explicitas == {calc}:
                linha.update(data=fmt(calc), situacao=CONFERIDO)
                linha["obs"].append("Prazo contado do fim do contrato: " + conta
                                    + (" (bate com a data escrita na cláusula)." if explicitas else "."))
            else:
                linha.update(data="; ".join(fmt(d) for d in sorted(explicitas)), situacao=REVISAR)
                linha["obs"].append("Divergência: a cláusula traz data diferente da conta "
                                    + conta + ".")
        else:
            linha.update(data="sem data certa", situacao=REVISAR)
            linha["obs"].append(f"Vale por {qtd} {un} após o fim do contrato, "
                                "mas o vencimento do contrato não tem data certa.")
    elif ind and not datas:
        linha.update(data="sem data certa", situacao=REVISAR)
        linha["obs"].append(f"O contrato diz \"{ind}\".")
    elif len({d for d, _ in datas}) == 1:
        linha.update(data=fmt(datas[0][0]), situacao=CONFERIDO)
        if not re.search(r"\d", datas[0][1]):
            linha["obs"].append(f"Data escrita por extenso no contrato: \"{datas[0][1]}\".")
    elif datas:
        linha.update(data="; ".join(fmt(d) for d, _ in datas), situacao=REVISAR)
        linha["obs"].append("Mais de uma data na cláusula; conferir qual vale.")
    else:
        linha.update(data="sem data certa", situacao=REVISAR)
        linha["obs"].append("A cláusula não traz prazo.")
    for c in cand:
        aplicar_alteracoes(linha, itens, c, tipo)
    return linha


def avisos_documento(itens):
    """Saltos na numeração e títulos repetidos: podem indicar páginas faltando."""
    avisos = []
    nums = [int(i["clausula"]) for i in itens if i["tipo"] == "titulo"]
    for a, b in zip(nums, nums[1:]):
        if b > a + 1:
            avisos.append(f"A numeração das cláusulas pula da {a} para a {b}; "
                          f"as cláusulas {a + 1} a {b - 1} não estão no arquivo.")
    titulos = [sem_acento(i["titulo"]) for i in itens if i["tipo"] == "titulo"]
    repetidos = sorted({t for t in titulos if t and titulos.count(t) > 1})
    if repetidos:
        avisos.append("Títulos de cláusula repetidos: " + ", ".join(r.upper() for r in repetidos) + ".")
    return avisos


def processar(arquivo, pasta):
    tipo = "pdf" if arquivo.suffix.lower() == ".pdf" else "word"
    try:
        itens, sem_texto, total = ler_pdf(arquivo) if tipo == "pdf" else ler_word(arquivo)
    except Exception as erro:
        motivo = f"Não foi possível abrir o arquivo ({erro})."
        return sem_leitura(arquivo, motivo)
    if tipo == "pdf" and len(sem_texto) == total:
        return sem_leitura(arquivo, "PDF digitalizado (só imagem, sem texto). "
                                    "Precisa de leitura manual ou de reconhecimento de texto.")
    avisos = avisos_documento(itens)
    if sem_texto:
        avisos.append("Páginas sem texto (imagem): " + ", ".join(map(str, sem_texto)) + ".")

    venc = analisar_vencimento(itens, tipo, pasta)
    nc = analisar_nao_concorrencia(itens, tipo, pasta, venc)
    procurado_nc = ("Procurei por: não concorrência, não competição, vedação de atuação, "
                    "exclusividade, restrição comercial, abster-se, concorrente, mesmo ramo, "
                    "impedido de contratar/fornecer.")
    procurado_v = "Procurei por: vigência, vigorará, prazo, vencimento."
    if venc is None:
        venc = {"trecho": "não encontrado", "fonte": "-", "data": "não encontrado",
                "situacao": REVISAR, "obs": [procurado_v]}
    if nc is None:
        nc = {"trecho": "não encontrado", "fonte": "-", "data": "não encontrado",
              "situacao": REVISAR, "obs": [procurado_nc + " Li todo o texto do arquivo."]}
        for p in parecidos(itens, tipo):
            nc["obs"].append(f"Há uma restrição parecida, que não tratei como não concorrência: "
                             f"{fonte(p, tipo)} ({p['titulo']}): \"{p['texto']}\". Conferir.")
        if avisos:
            nc["obs"].append("Não dá para afirmar que a cláusula não existe: " + " ".join(avisos))
    return {"arquivo": arquivo.name, "nc": nc, "venc": venc, "avisos": avisos,
            "situacao": REVISAR if REVISAR in (nc["situacao"], venc["situacao"]) else CONFERIDO}


def sem_leitura(arquivo, motivo):
    vazio = {"trecho": "sem leitura", "fonte": "-", "data": "-", "situacao": SEM_LEITURA,
             "obs": [motivo]}
    return {"arquivo": arquivo.name, "nc": dict(vazio), "venc": dict(vazio), "avisos": [],
            "situacao": SEM_LEITURA, "motivo": motivo}


# ---------------------------------------------------------------- planilha
CORES = {CONFERIDO: "C6EFCE", REVISAR: "FFEB9C", SEM_LEITURA: "FFC7CE"}


def aba(wb, nome, resultados, chave, titulo_data):
    ws = wb.create_sheet(nome)
    ws.append(["Arquivo", "Trecho encontrado", titulo_data, "Fonte", "Situação", "Observação"])
    for r in resultados:
        x = r[chave]
        ws.append([r["arquivo"], x["trecho"], x["data"], x["fonte"], x["situacao"],
                   " ".join(x["obs"])])
        ws.cell(ws.max_row, 5).fill = PatternFill("solid", fgColor=CORES[x["situacao"]])
    formatar(ws, [28, 70, 22, 24, 13, 55])


def formatar(ws, larguras):
    for c in ws[1]:
        c.font = Font(bold=True)
    for idx, w in enumerate(larguras, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(idx)].width = w
    for linha in ws.iter_rows():
        for c in linha:
            c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "A2"


def gerar_planilha(resultados, saida, pasta):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    aba(wb, "Não concorrência", resultados, "nc", "Data (até quando vale)")
    aba(wb, "Vencimento", resultados, "venc", "Data de vencimento")

    ws = wb.create_sheet("Resumo")
    cont = {s: sum(r["situacao"] == s for r in resultados) for s in (CONFERIDO, REVISAR, SEM_LEITURA)}
    ws.append(["Pasta lida", str(pasta)])
    ws.append(["Documentos na pasta", len(resultados)])
    ws.append(["Lidos", len(resultados) - cont[SEM_LEITURA]])
    ws.append(["Conferidos (as duas informações)", cont[CONFERIDO]])
    ws.append(["Para revisar", cont[REVISAR]])
    ws.append(["Sem leitura", cont[SEM_LEITURA]])
    ws.append([])
    ws.append(["O que não foi possível conferir", "Motivo"])
    ws.cell(ws.max_row, 1).font = Font(bold=True)
    ws.cell(ws.max_row, 2).font = Font(bold=True)
    for r in resultados:
        for chave, nome in (("nc", "Não concorrência"), ("venc", "Vencimento")):
            x = r[chave]
            if x["situacao"] != CONFERIDO:
                ws.append([f"{r['arquivo']} - {nome}", f"{x['data']}. " + " ".join(x["obs"])])
                if r["situacao"] == SEM_LEITURA:
                    break
    ws.append([])
    ws.append(["Pontos de atenção nos documentos", ""])
    ws.cell(ws.max_row, 1).font = Font(bold=True)
    for r in resultados:
        for a in r["avisos"]:
            ws.append([r["arquivo"], a])
    formatar(ws, [45, 100])
    ws.cell(1, 1).font = Font(bold=True)
    wb.move_sheet("Resumo", offset=-2)
    wb.save(saida)
    return cont


def main():
    pasta = Path(sys.argv[1] if len(sys.argv) > 1 else PASTA_PADRAO)
    saida = Path(sys.argv[2] if len(sys.argv) > 2 else SAIDA_PADRAO)
    if not pasta.is_dir():
        sys.exit(f"Não encontrei a pasta \"{pasta}\".")
    arquivos = sorted(f for f in pasta.iterdir()
                      if f.suffix.lower() in (".pdf", ".docx") and not f.name.startswith("~$"))
    ignorados = sorted(f.name for f in pasta.iterdir()
                       if f.is_file() and f not in arquivos)
    resultados = [processar(f, pasta) for f in arquivos]
    cont = gerar_planilha(resultados, saida, pasta)
    sys.stdout.reconfigure(encoding="utf-8")
    print(f"Documentos: {len(resultados)} | conferidos: {cont[CONFERIDO]} | "
          f"para revisar: {cont[REVISAR]} | sem leitura: {cont[SEM_LEITURA]}")
    if ignorados:
        print("Arquivos ignorados (não são PDF nem Word .docx): " + ", ".join(ignorados))
    print(f"Planilha salva em: {saida.resolve()}")


if __name__ == "__main__":
    main()
