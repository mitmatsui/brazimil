"""
omie_sync.py — Normalizador de dados Omie + Bradesco
=====================================================
Roda uma vez por semana antes do gerar.py.
Lê os CSVs exportados e gera os JSONs que o dashboard consome.

ARQUIVOS DE ENTRADA (colocar na pasta brazimil-dashboard):
  vendas_omie.csv     — exportado do Omie: Relatório de Faturamento
  boletos_bradesco.csv — exportado do Bradesco Net Empresa: Carteira de Cobrança

ARQUIVOS DE SAÍDA (gerados automaticamente):
  vendas.json         — histórico semanal de vendas por produto e cliente
  fluxo_caixa.json    — boletos projetados para as próximas 4 semanas

COMO EXPORTAR DO OMIE:
  1. Acesse: Relatórios > Vendas > Faturamento por Período
  2. Selecione as últimas 13 semanas
  3. Exporte como CSV (separador ponto-e-vírgula)
  Colunas esperadas: Data Emissão; Nº NF; Cliente; Produto; Qtd; Valor Unit; Valor Total

COMO EXPORTAR DO BRADESCO:
  1. Acesse Bradesco Net Empresa
  2. Cobrança > Consulta de Títulos > Exportar
  3. Formato CSV
  Colunas esperadas: Nosso Nº; Seu Nº; Cliente; CPF/CNPJ; Emissão; Vencimento; Valor; Status
"""

import csv, json, os, re
from datetime import datetime, date, timedelta

# ── Configuração ────────────────────────────────────────────────────────────
PASTA        = os.path.dirname(os.path.abspath(__file__))
ARQ_VENDAS   = os.path.join(PASTA, "vendas_omie.csv")
ARQ_BOLETOS  = os.path.join(PASTA, "boletos_bradesco.csv")
SAIDA_VENDAS = os.path.join(PASTA, "vendas.json")
SAIDA_FLUXO  = os.path.join(PASTA, "fluxo_caixa.json")

HOJE = date.today()

# ── Utilidades ──────────────────────────────────────────────────────────────
def parse_data(s):
    """Aceita dd/mm/yyyy, yyyy-mm-dd ou dd-mm-yyyy."""
    s = str(s).strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None

def parse_valor(s):
    """Converte strings como '1.234,56' ou '1234.56' para float."""
    s = str(s).strip().replace(" ", "")
    s = re.sub(r"[^\d,\.]", "", s)
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return round(float(s), 2)
    except ValueError:
        return 0.0

def semana_iso(d):
    """Retorna 'AAAA-Snn' ex: '2025-S23'."""
    ano, sem, _ = d.isocalendar()
    return f"{ano}-S{sem:02d}"

def label_semana(d):
    """Retorna 'Sem 23/Jun' para a semana que contém a data d."""
    seg = d - timedelta(days=d.weekday())
    dom = seg + timedelta(days=6)
    _, sem, _ = d.isocalendar()
    return f"Sem {sem:02d} ({seg.strftime('%d/%b')}–{dom.strftime('%d/%b')})"

def detectar_separador(arquivo):
    """Detecta se o CSV usa ; ou ,"""
    with open(arquivo, encoding="utf-8-sig", errors="replace") as f:
        linha = f.readline()
    return ";" if linha.count(";") >= linha.count(",") else ","

def normalizar_coluna(nome):
    """Remove acentos e normaliza nomes de colunas."""
    mapa = {"á":"a","ã":"a","â":"a","à":"a","é":"e","ê":"e","í":"i",
            "ó":"o","ô":"o","õ":"o","ú":"u","ç":"c","Á":"A","Ã":"A",
            "É":"E","Í":"I","Ó":"O","Ú":"U","Ç":"C"}
    r = nome.strip().lower()
    for k, v in mapa.items():
        r = r.replace(k, v)
    return re.sub(r"[^a-z0-9_]", "_", r).strip("_")

# ════════════════════════════════════════════════════════════════════════════
# BLOCO 1 — PROCESSAMENTO DO CSV DE VENDAS OMIE
# ════════════════════════════════════════════════════════════════════════════
def processar_vendas():
    if not os.path.exists(ARQ_VENDAS):
        print(f"  [AVISO] {ARQ_VENDAS} não encontrado — gerando dados de exemplo.")
        return gerar_exemplo_vendas()

    sep  = detectar_separador(ARQ_VENDAS)
    rows = []
    with open(ARQ_VENDAS, encoding="utf-8-sig", errors="replace") as f:
        reader = csv.DictReader(f, delimiter=sep)
        cols_orig = reader.fieldnames or []
        col_map   = {normalizar_coluna(c): c for c in cols_orig}

        for row in reader:
            # Mapeamento flexível de colunas — adapta ao relatório do Omie
            def get(*chaves):
                for ch in chaves:
                    ch_n = normalizar_coluna(ch)
                    if ch_n in col_map:
                        return str(row.get(col_map[ch_n], "")).strip()
                    for orig in cols_orig:
                        if ch.lower() in orig.lower():
                            return str(row.get(orig, "")).strip()
                return ""

            data_str   = get("Data Emissão","Data NF","Data","Emissão","data_emissao")
            cliente    = get("Cliente","Razão Social","cliente","razao_social")
            produto    = get("Produto","Descrição","produto","descricao","item")
            valor_str  = get("Valor Total","Total","valor_total","valor","vlr_total")
            qtd_str    = get("Qtd","Quantidade","qtd","quantidade")
            nf         = get("NF","Nº NF","nf","nota","numero_nf")

            data = parse_data(data_str)
            if not data:
                continue
            valor = parse_valor(valor_str)
            qtd   = parse_valor(qtd_str) if qtd_str else 1

            rows.append({
                "data":    data.isoformat(),
                "semana":  semana_iso(data),
                "label":   label_semana(data),
                "cliente": cliente or "N/I",
                "produto": produto or "N/I",
                "nf":      nf or "",
                "qtd":     qtd,
                "valor":   valor,
            })

    if not rows:
        print("  [AVISO] CSV de vendas sem dados válidos — gerando exemplo.")
        return gerar_exemplo_vendas()

    return montar_estrutura_vendas(rows)


def montar_estrutura_vendas(rows):
    """Agrupa os registros nas dimensões necessárias pelo dashboard."""
    # Últimas 13 semanas disponíveis
    semanas_presentes = sorted(set(r["semana"] for r in rows))
    semanas_13        = semanas_presentes[-13:]

    rows_13 = [r for r in rows if r["semana"] in semanas_13]

    # ── Por semana (total) ──────────────────────────────────────────────────
    por_semana = {}
    for r in rows_13:
        s = r["semana"]
        if s not in por_semana:
            por_semana[s] = {"semana": s, "label": r["label"], "total": 0.0, "nfs": 0}
        por_semana[s]["total"] = round(por_semana[s]["total"] + r["valor"], 2)
        por_semana[s]["nfs"]  += 1

    # ── Por produto (top 10 por valor total) ────────────────────────────────
    por_produto = {}
    for r in rows_13:
        p = r["produto"]
        if p not in por_produto:
            por_produto[p] = {"produto": p, "total": 0.0, "qtd": 0.0, "semanas": {}}
        por_produto[p]["total"] = round(por_produto[p]["total"] + r["valor"], 2)
        por_produto[p]["qtd"]  += r["qtd"]
        s = r["semana"]
        por_produto[p]["semanas"][s] = round(por_produto[p]["semanas"].get(s, 0) + r["valor"], 2)

    top_produtos = sorted(por_produto.values(), key=lambda x: x["total"], reverse=True)[:10]

    # ── Por cliente (top 10) ─────────────────────────────────────────────────
    por_cliente = {}
    for r in rows_13:
        c = r["cliente"]
        if c not in por_cliente:
            por_cliente[c] = {"cliente": c, "total": 0.0, "nfs": 0}
        por_cliente[c]["total"] = round(por_cliente[c]["total"] + r["valor"], 2)
        por_cliente[c]["nfs"]  += 1

    top_clientes = sorted(por_cliente.values(), key=lambda x: x["total"], reverse=True)[:10]

    # ── Comparativo semana atual vs semana anterior ──────────────────────────
    sem_atual   = semana_iso(HOJE)
    sem_ant     = semana_iso(HOJE - timedelta(weeks=1))
    total_atual = por_semana.get(sem_atual, {}).get("total", 0)
    total_ant   = por_semana.get(sem_ant,   {}).get("total", 0)
    variacao    = round(((total_atual - total_ant) / total_ant * 100) if total_ant else 0, 1)

    resultado = {
        "gerado_em": datetime.now().isoformat(),
        "total_periodo": round(sum(r["valor"] for r in rows_13), 2),
        "total_semana_atual": total_atual,
        "variacao_semanal_pct": variacao,
        "semanas": sorted(por_semana.values(), key=lambda x: x["semana"]),
        "labels_semanas": [por_semana[s]["label"] for s in semanas_13 if s in por_semana],
        "top_produtos": top_produtos,
        "top_clientes": top_clientes,
        "exemplo": False,
    }

    total_regs = len(rows_13)
    print(f"  Vendas processadas: {total_regs} registros em {len(semanas_13)} semanas")
    print(f"  Faturamento período: R$ {resultado['total_periodo']:,.2f}")
    return resultado


def gerar_exemplo_vendas():
    """Gera dados de exemplo para visualização enquanto o CSV real não chega."""
    import random
    random.seed(42)
    produtos_ex = [
        "Econômica 15L BN","Semi Brilho 15L BN","Corrida 20KG",
        "Esmalte Sint. 3L","Externa 15L BN","Vedamil 15KG",
        "Piso 15L BN","Acrílica 20KG","Verniz 3L","Zarcão 3L"
    ]
    clientes_ex = [
        "Tintas Manaus Ltda","Constrular AM","Depósito Central",
        "Ferro & Cor","Material AM","Reforma Total","ConstruBem"
    ]
    rows = []
    for sem_offset in range(12, -1, -1):
        seg = HOJE - timedelta(days=HOJE.weekday()) - timedelta(weeks=sem_offset)
        for _ in range(random.randint(12, 28)):
            d = seg + timedelta(days=random.randint(0, 4))
            p = random.choice(produtos_ex)
            v = round(random.uniform(800, 8500), 2)
            rows.append({
                "data":    d.isoformat(),
                "semana":  semana_iso(d),
                "label":   label_semana(d),
                "cliente": random.choice(clientes_ex),
                "produto": p,
                "nf":      str(random.randint(1000, 9999)),
                "qtd":     random.randint(1, 20),
                "valor":   v,
            })
    print("  [EXEMPLO] Dados de vendas simulados gerados (13 semanas)")
    resultado = montar_estrutura_vendas(rows)
    resultado["exemplo"] = True
    return resultado


# ════════════════════════════════════════════════════════════════════════════
# BLOCO 2 — PROCESSAMENTO DO CSV DE BOLETOS BRADESCO
# ════════════════════════════════════════════════════════════════════════════
def processar_boletos():
    if not os.path.exists(ARQ_BOLETOS):
        print(f"  [AVISO] {ARQ_BOLETOS} não encontrado — gerando dados de exemplo.")
        return gerar_exemplo_boletos()

    sep  = detectar_separador(ARQ_BOLETOS)
    rows = []
    with open(ARQ_BOLETOS, encoding="utf-8-sig", errors="replace") as f:
        reader = csv.DictReader(f, delimiter=sep)
        cols_orig = reader.fieldnames or []

        for row in reader:
            def get(*chaves):
                for ch in chaves:
                    for orig in cols_orig:
                        if ch.lower() in orig.lower():
                            return str(row.get(orig, "")).strip()
                return ""

            nosso_num  = get("Nosso","Nosso Nº","nosso_numero")
            cliente    = get("Cliente","Sacado","Pagador","cliente","sacado")
            vencimento = get("Vencimento","Venc","data_vencimento","vencto")
            emissao    = get("Emissão","Emissao","data_emissao","emissao")
            valor_str  = get("Valor","vlr","value","valor_documento","valor_titulo")
            status_raw = get("Status","Situação","Situacao","status","situacao")

            data_venc = parse_data(vencimento)
            if not data_venc:
                continue
            valor = parse_valor(valor_str)
            if valor <= 0:
                continue

            # Normaliza status
            s = status_raw.upper()
            if any(x in s for x in ["PAGO","LIQUIDADO","BAIXADO","RECEBIDO"]):
                status = "pago"
            elif any(x in s for x in ["VENCIDO","ATRASADO","INADIMPL"]):
                status = "vencido"
            elif any(x in s for x in ["CANC","BAIXA","DEVOLVIDO"]):
                status = "cancelado"
            else:
                status = "aberto"

            dias_venc = (data_venc - HOJE).days

            rows.append({
                "nosso_num": nosso_num,
                "cliente":   cliente or "N/I",
                "emissao":   parse_data(emissao).isoformat() if parse_data(emissao) else "",
                "vencimento":data_venc.isoformat(),
                "semana":    semana_iso(data_venc),
                "label":     label_semana(data_venc),
                "valor":     valor,
                "status":    status,
                "dias_venc": dias_venc,
            })

    if not rows:
        print("  [AVISO] CSV de boletos sem dados válidos — gerando exemplo.")
        return gerar_exemplo_boletos()

    return montar_estrutura_fluxo(rows)


def montar_estrutura_fluxo(rows):
    """Monta o fluxo de caixa das próximas 4 semanas + histórico de inadimplência."""
    # Próximas 4 semanas
    semanas_futuras = []
    seg_corrente = HOJE - timedelta(days=HOJE.weekday())
    for i in range(4):
        seg = seg_corrente + timedelta(weeks=i)
        dom = seg + timedelta(days=6)
        s   = semana_iso(seg)
        semanas_futuras.append({
            "semana": s,
            "label":  label_semana(seg),
            "inicio": seg.isoformat(),
            "fim":    dom.isoformat(),
            "previsto":     0.0,
            "vencido_nper": 0.0,
            "boletos_n":    0,
        })
    sem_map = {s["semana"]: s for s in semanas_futuras}

    # Preenche previsão para abertos e vencidos recentes
    boletos_futuros = []
    for r in rows:
        if r["status"] in ("pago", "cancelado"):
            continue
        if r["semana"] in sem_map:
            sem_map[r["semana"]]["previsto"]     = round(sem_map[r["semana"]]["previsto"] + r["valor"], 2)
            sem_map[r["semana"]]["boletos_n"]   += 1
            if r["status"] == "vencido":
                sem_map[r["semana"]]["vencido_nper"] = round(sem_map[r["semana"]]["vencido_nper"] + r["valor"], 2)
            boletos_futuros.append(r)

    # Inadimplência histórica (vencidos antes de hoje)
    vencidos = [r for r in rows if r["status"] == "vencido" and r["dias_venc"] < 0]
    total_vencido  = round(sum(r["valor"] for r in vencidos), 2)
    total_aberto   = round(sum(r["valor"] for r in rows if r["status"] in ("aberto","vencido")), 2)
    total_recebido = round(sum(r["valor"] for r in rows if r["status"] == "pago"), 2)

    # Aging — faixas de vencimento dos inadimplentes
    aging = {"1_30": 0.0, "31_60": 0.0, "61_90": 0.0, "mais_90": 0.0}
    for r in vencidos:
        dias = abs(r["dias_venc"])
        if dias <= 30:   aging["1_30"]    = round(aging["1_30"]    + r["valor"], 2)
        elif dias <= 60: aging["31_60"]   = round(aging["31_60"]   + r["valor"], 2)
        elif dias <= 90: aging["61_90"]   = round(aging["61_90"]   + r["valor"], 2)
        else:            aging["mais_90"] = round(aging["mais_90"] + r["valor"], 2)

    # Top inadimplentes por valor
    por_cliente_venc = {}
    for r in vencidos:
        c = r["cliente"]
        por_cliente_venc[c] = round(por_cliente_venc.get(c, 0) + r["valor"], 2)
    top_inadimplentes = sorted(
        [{"cliente": k, "valor": v} for k, v in por_cliente_venc.items()],
        key=lambda x: x["valor"], reverse=True
    )[:8]

    taxa_inadimplencia = round((total_vencido / total_aberto * 100) if total_aberto else 0, 1)

    resultado = {
        "gerado_em": datetime.now().isoformat(),
        "resumo": {
            "total_carteira":      round(sum(r["valor"] for r in rows), 2),
            "total_aberto":        total_aberto,
            "total_recebido":      total_recebido,
            "total_vencido":       total_vencido,
            "taxa_inadimplencia":  taxa_inadimplencia,
            "qtd_boletos_abertos": len([r for r in rows if r["status"] in ("aberto","vencido")]),
            "qtd_boletos_vencidos":len(vencidos),
        },
        "proximas_4_semanas": semanas_futuras,
        "aging": aging,
        "top_inadimplentes": top_inadimplentes,
        "boletos_detalhes": sorted(boletos_futuros, key=lambda x: x["vencimento"])[:50],
        "exemplo": False,
    }

    print(f"  Boletos processados: {len(rows)} títulos")
    print(f"  Carteira total: R$ {resultado['resumo']['total_carteira']:,.2f}")
    print(f"  Inadimplência: {taxa_inadimplencia}% (R$ {total_vencido:,.2f})")
    return resultado


def gerar_exemplo_boletos():
    """Gera dados de exemplo para visualização."""
    import random
    random.seed(99)
    clientes_ex = [
        "Tintas Manaus Ltda","Constrular AM","Depósito Central",
        "Ferro & Cor","Material AM","Reforma Total","ConstruBem",
        "DepósitoAM","MisTintas","Tudo & Cor"
    ]
    rows = []
    for i in range(60):
        # Distribui vencimentos: passado (-90 a -1), futuro (0 a +90)
        offset = random.randint(-90, 90)
        d_venc = HOJE + timedelta(days=offset)
        valor  = round(random.uniform(500, 15000), 2)
        if offset < -30:
            status = random.choice(["vencido","vencido","pago"])
        elif offset < 0:
            status = random.choice(["vencido","aberto","pago"])
        else:
            status = "aberto"
        rows.append({
            "nosso_num": f"000{i+1:05d}",
            "cliente":   random.choice(clientes_ex),
            "emissao":   (d_venc - timedelta(days=random.randint(25,90))).isoformat(),
            "vencimento":d_venc.isoformat(),
            "semana":    semana_iso(d_venc),
            "label":     label_semana(d_venc),
            "valor":     valor,
            "status":    status,
            "dias_venc": offset,
        })
    print("  [EXEMPLO] Dados de boletos simulados gerados")
    resultado = montar_estrutura_fluxo(rows)
    resultado["exemplo"] = True
    return resultado


# ════════════════════════════════════════════════════════════════════════════
# MAIN
# ════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    print("=" * 55)
    print("omie_sync.py — Normalizador Omie + Bradesco")
    print(f"Data de referência: {HOJE.strftime('%d/%m/%Y')}")
    print("=" * 55)

    print("\n[1/2] Processando vendas Omie...")
    vendas = processar_vendas()
    with open(SAIDA_VENDAS, "w", encoding="utf-8") as f:
        json.dump(vendas, f, ensure_ascii=False, indent=2)
    print(f"  → {SAIDA_VENDAS}")

    print("\n[2/2] Processando boletos Bradesco...")
    fluxo = processar_boletos()
    with open(SAIDA_FLUXO, "w", encoding="utf-8") as f:
        json.dump(fluxo, f, ensure_ascii=False, indent=2)
    print(f"  → {SAIDA_FLUXO}")

    print("\n" + "=" * 55)
    if vendas.get("exemplo") or fluxo.get("exemplo"):
        print("⚠️  MODO EXEMPLO: coloque os CSVs reais na pasta e rode novamente.")
        print("    vendas_omie.csv      → Omie: Relatórios > Faturamento")
        print("    boletos_bradesco.csv → Bradesco Net Empresa: Carteira de Cobrança")
    else:
        print("✅  Dados reais processados com sucesso!")
    print("   Execute agora: python3 gerar.py")
    print("=" * 55)
