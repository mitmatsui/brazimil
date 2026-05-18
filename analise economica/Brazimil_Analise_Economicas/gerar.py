import json, os, openpyxl
from datetime import datetime

print("Lendo Excel de custos e pesos...")
wb1 = openpyxl.load_workbook("SIMULADOR TINTAS MIL BRAZIMIL 1.xlsx", data_only=True)

dados_custo = {}
for row in wb1["Custo"].iter_rows(min_row=2, values_only=True):
    if row[0] and row[1] and isinstance(row[1], (int,float)) and row[1] > 0:
        dados_custo[row[0]] = round(row[1], 4)

dados_peso = {}
for row in wb1["Peso"].iter_rows(min_row=2, values_only=True):
    if row[0] and row[1] and isinstance(row[1], (int,float)) and row[1] > 0:
        dados_peso[row[0]] = round(row[1], 4)

print(f"  Custos: {len(dados_custo)} | Pesos: {len(dados_peso)}")

print("Lendo Excel do time (preços praticados)...")
wb2 = openpyxl.load_workbook("Simulador_Brazimil_Analise_do_time_.xlsx", data_only=True)
ws_tab = wb2["Tabela de Produtos"]

precos_time = {}
custos_time = {}
for row in ws_tab.iter_rows(min_row=3, values_only=True):
    cod  = row[0]
    nome = row[1]
    custo_insumo = row[12]   # col 13 — custo insumos
    custo_embal  = row[13]   # col 14 — custo embalagem
    frete        = row[15]   # col 16 — frete
    preco        = row[20]   # col 21 — preço final praticado
    if cod and nome and preco and isinstance(preco, (int,float)) and preco > 0:
        precos_time[str(cod).strip()] = round(float(preco), 2)
        custo_total = 0
        for v in [custo_insumo, custo_embal]:
            if v and isinstance(v, (int,float)):
                custo_total += v
        custos_time[str(cod).strip()] = round(custo_total, 4)

print(f"  Preços do time: {len(precos_time)} produtos")

# Tabela mestre: nome Excel → (categoria, subcategoria, cod_time)
tabela = {
    "BRAZIMIL ECONÔMICA 3,0L REFIL BRANCO NEVE":      ("Tinta Base Água","Econômica",     None),
    "BRAZIMIL ECONÔMICA 3,0L REFIL COR":               ("Tinta Base Água","Econômica",     None),
    "BRAZIMIL ECONÔMICA 3,0L GALÃO BRANCO NEVE":       ("Tinta Base Água","Econômica",     None),
    "BRAZIMIL ECONÔMICA 3,0L GALÃO COR":               ("Tinta Base Água","Econômica",     None),
    "BRAZIMIL ECONÔMICA 15L BALDE BRANCO NEVE":        ("Tinta Base Água","Econômica",     "PRD00039"),
    "BRAZIMIL ECONÔMICA 15L BALDE COR":                ("Tinta Base Água","Econômica",     "PRD00040"),
    "BRAZIMIL EXTERNA 3,0L BALDE BRANCO NEVE":         ("Tinta Base Água","Externa",       None),
    "BRAZIMIL EXTERNA 3,0L BALDE COR":                 ("Tinta Base Água","Externa",       None),
    "BRAZIMIL EXTERNA 15L BALDE BRANCO NEVE":          ("Tinta Base Água","Externa",       None),
    "BRAZIMIL EXTERNA 15L BALDE COR":                  ("Tinta Base Água","Externa",       None),
    "BRAZIMIL SEMI BRILHO 3,0L BALDE BRANCO NEVE":     ("Tinta Base Água","Semi Brilho",   None),
    "BRAZIMIL SEMI BRILHO 3,0L BALDE COR":             ("Tinta Base Água","Semi Brilho",   None),
    "BRAZIMIL SEMI BRILHO 15L BALDE BRANCO NEVE":      ("Tinta Base Água","Semi Brilho",   None),
    "BRAZIMIL SEMI BRILHO 15L BALDE COR":              ("Tinta Base Água","Semi Brilho",   None),
    "BRAZIMIL TINTA PISO 3,0L BALDE BRANCO NEVE":      ("Tinta Base Água","Piso",          None),
    "BRAZIMIL TINTA PISO 3,0L BALDE COR":              ("Tinta Base Água","Piso",          None),
    "BRAZIMIL TINTA PISO 15L BALDE BRANCO NEVE":       ("Tinta Base Água","Piso",          None),
    "BRAZIMIL TINTA PISO 15L BALDE COR":               ("Tinta Base Água","Piso",          None),
    "BRAZIMIL SELADOR ACRIL 3,0L BALDE PIGMENTADO":    ("Tinta Base Água","Selador Acril", "PRD00003"),
    "BRAZIMIL SELADOR ACRIL 15L BALDE PIGMENTADO":     ("Tinta Base Água","Selador Acril", "PRD00004"),
    "BRAZIMIL TINTA P GESSO 3,0L":                     ("Tinta Base Água","Gesso",         None),
    "BRAZIMIL TINTA P GESSO 15L":                      ("Tinta Base Água","Gesso",         None),
    "LIQUIBRILHO 3,0L BALDE INCOLOR":                  ("Tinta Base Água","Liquibrilho",   None),
    "LIQUIBRILHO 15L BALDE INCOLOR":                   ("Tinta Base Água","Liquibrilho",   None),
    "BRAZIMIL ESMALTE SINTÉTICO 0,75L LATA ALUMÍNIO":  ("Tinta Base Solvente","Esmalte Sint.", None),
    "BRAZIMIL ESMALTE SINTÉTICO 0,75L LATA BRANCO":    ("Tinta Base Solvente","Esmalte Sint.", None),
    "BRAZIMIL ESMALTE SINTÉTICO 0,75L LATA PRETO":     ("Tinta Base Solvente","Esmalte Sint.", None),
    "BRAZIMIL ESMALTE SINTÉTICO 0,75L LATA COR":       ("Tinta Base Solvente","Esmalte Sint.", None),
    "BRAZIMIL ESMALTE SINTÉTICO 3,0L LATA ALUMÍNIO":   ("Tinta Base Solvente","Esmalte Sint.", None),
    "BRAZIMIL ESMALTE SINTÉTICO 3,0L LATA BRANCO":     ("Tinta Base Solvente","Esmalte Sint.", None),
    "BRAZIMIL ESMALTE SINTÉTICO 3,0L LATA PRETO":      ("Tinta Base Solvente","Esmalte Sint.", None),
    "BRAZIMIL ESMALTE SINTÉTICO 3,0L LATA COR":        ("Tinta Base Solvente","Esmalte Sint.", None),
    "BRAZIMIL ESMALTE SINTÉTICO 3,0L LATA FOSCO":      ("Tinta Base Solvente","Esmalte Sint.", None),
    "ESMALTE BASE ÁGUA 0,75L":                         ("Tinta Base Solvente","Esmalte Base Água", None),
    "ESMALTE BASE ÁGUA 3,0L":                          ("Tinta Base Solvente","Esmalte Base Água", None),
    "BRAZIMIL VERNIZ 0,75L LATA INCOLOR":              ("Tinta Base Solvente","Verniz",     None),
    "BRAZIMIL VERNIZ 0,75L LATA PIGMENTADO":           ("Tinta Base Solvente","Verniz",     None),
    "BRAZIMIL VERNIZ 3,0L LATA INCOLOR":               ("Tinta Base Solvente","Verniz",     None),
    "BRAZIMIL VERNIZ 3,0L LATA PIGMENTADO":            ("Tinta Base Solvente","Verniz",     None),
    "BRAZIMIL ZARCÃO 0,75L LATA BRANCO":               ("Tinta Base Solvente","Zarcão",     None),
    "BRAZIMIL ZARCÃO 0,75L LATA CINZA":                ("Tinta Base Solvente","Zarcão",     None),
    "BRAZIMIL ZARCÃO 0,75L LATA COR":                  ("Tinta Base Solvente","Zarcão",     None),
    "BRAZIMIL ZARCÃO 3,0L LATA CINZA":                 ("Tinta Base Solvente","Zarcão",     None),
    "BRAZIMIL ZARCÃO  18L LATA COR":                   ("Tinta Base Solvente","Zarcão",     None),
    "BRAZIMIL SELADOR P/ MAD. 0,75L LATA INCOLOR":     ("Tinta Base Solvente","Selador Mad.", None),
    "BRAZIMIL SELADOR P/ MAD. 3,0L LATA INCOLOR":      ("Tinta Base Solvente","Selador Mad.", None),
    "BRAZIMIL CORRIDA POTE 1KG":                       ("Massas e Texturas","Corrida",      "PRD00001"),
    "BRAZIMIL CORRIDA GALÃO 5KG":                      ("Massas e Texturas","Corrida",      None),
    "BRAZIMIL CORRIDA SACO VALV 10 KG":                ("Massas e Texturas","Corrida",      None),
    "BRAZIMIL CORRIDA BALDE 20KG":                     ("Massas e Texturas","Corrida",      "PRD00002"),
    "BRAZIMIL ACRILICA POTE 1KG":                      ("Massas e Texturas","Acrílica",     None),
    "BRAZIMIL ACRILICA GALÃO 5KG":                     ("Massas e Texturas","Acrílica",     "PRD00007"),
    "BRAZIMIL ACRILICA SACO VALV 10,0KG":              ("Massas e Texturas","Acrílica",     None),
    "BRAZIMIL ACRILICA BALDE 20KG":                    ("Massas e Texturas","Acrílica",     "PRD00008"),
    "VEDAMIL / MANTA LÍQUIDA BALDE 3,6KG":            ("Complementos","Vedamil",            None),
    "VEDAMIL / MANTA LÍQUIDA BALDE 12KG":             ("Complementos","Vedamil",            None),
    "VEDAMIL / MANTA LÍQUIDA BALDE 15KG":             ("Complementos","Vedamil",            None),
    "VEDAMIL / MANTA LÍQUIDA BALDE 18KG":             ("Complementos","Vedamil",            None),
    "BRAZIMIL BORRACHA LIQUIDA 3,6 KG":               ("Complementos","Borracha Líquida",   None),
    "BRAZIMIL BORRACHA LIQUIDA 14,0 KG":              ("Complementos","Borracha Líquida",   None),
    "CIMENTO QUEIMADO GALÃO":                          ("Complementos","Cimento Queimado",   None),
    "BRAZIMIL COLA BRANCA 0,5KG":                      ("Complementos","Cola e Solvente",    None),
    "BRAZIMIL COLA BRANCA 1,0KG":                      ("Complementos","Cola e Solvente",    None),
    "AGUARRAS 0,900":                                   ("Complementos","Cola e Solvente",   None),
    "AGUARRAS 0,450":                                   ("Complementos","Cola e Solvente",   None),
}

# Preços de venda — prioridade: planilha do time > fallback manual
precos_fallback = {
    "BRAZIMIL ECONÔMICA 3,0L REFIL BRANCO NEVE":       9.50,
    "BRAZIMIL ECONÔMICA 3,0L REFIL COR":                9.50,
    "BRAZIMIL ECONÔMICA 3,0L GALÃO BRANCO NEVE":       17.99,
    "BRAZIMIL ECONÔMICA 3,0L GALÃO COR":               17.99,
    "BRAZIMIL ECONÔMICA 15L BALDE BRANCO NEVE":        69.90,
    "BRAZIMIL ECONÔMICA 15L BALDE COR":                69.90,
    "BRAZIMIL EXTERNA 3,0L BALDE BRANCO NEVE":         26.00,
    "BRAZIMIL EXTERNA 3,0L BALDE COR":                 24.20,
    "BRAZIMIL EXTERNA 15L BALDE BRANCO NEVE":          94.00,
    "BRAZIMIL EXTERNA 15L BALDE COR":                  90.11,
    "BRAZIMIL SEMI BRILHO 3,0L BALDE BRANCO NEVE":     43.00,
    "BRAZIMIL SEMI BRILHO 3,0L BALDE COR":             35.00,
    "BRAZIMIL SEMI BRILHO 15L BALDE BRANCO NEVE":     170.00,
    "BRAZIMIL SEMI BRILHO 15L BALDE COR":             170.00,
    "BRAZIMIL TINTA PISO 3,0L BALDE BRANCO NEVE":      30.00,
    "BRAZIMIL TINTA PISO 3,0L BALDE COR":              28.50,
    "BRAZIMIL TINTA PISO 15L BALDE BRANCO NEVE":      130.00,
    "BRAZIMIL TINTA PISO 15L BALDE COR":              108.00,
    "BRAZIMIL TINTA P GESSO 3,0L":                     20.00,
    "BRAZIMIL TINTA P GESSO 15L":                      73.00,
    "LIQUIBRILHO 3,0L BALDE INCOLOR":                  32.00,
    "LIQUIBRILHO 15L BALDE INCOLOR":                  135.00,
    "BRAZIMIL ESMALTE SINTÉTICO 0,75L LATA ALUMÍNIO":  27.20,
    "BRAZIMIL ESMALTE SINTÉTICO 0,75L LATA BRANCO":    16.11,
    "BRAZIMIL ESMALTE SINTÉTICO 0,75L LATA PRETO":     16.87,
    "BRAZIMIL ESMALTE SINTÉTICO 0,75L LATA COR":       16.87,
    "BRAZIMIL ESMALTE SINTÉTICO 3,0L LATA ALUMÍNIO":   99.00,
    "BRAZIMIL ESMALTE SINTÉTICO 3,0L LATA BRANCO":     52.26,
    "BRAZIMIL ESMALTE SINTÉTICO 3,0L LATA PRETO":      39.90,
    "BRAZIMIL ESMALTE SINTÉTICO 3,0L LATA COR":        46.90,
    "BRAZIMIL ESMALTE SINTÉTICO 3,0L LATA FOSCO":      54.50,
    "ESMALTE BASE ÁGUA 0,75L":                         20.00,
    "ESMALTE BASE ÁGUA 3,0L":                          80.00,
    "BRAZIMIL VERNIZ 0,75L LATA INCOLOR":              20.50,
    "BRAZIMIL VERNIZ 0,75L LATA PIGMENTADO":           20.50,
    "BRAZIMIL VERNIZ 3,0L LATA INCOLOR":               63.00,
    "BRAZIMIL VERNIZ 3,0L LATA PIGMENTADO":            63.00,
    "BRAZIMIL ZARCÃO 0,75L LATA BRANCO":               15.90,
    "BRAZIMIL ZARCÃO 0,75L LATA CINZA":                14.90,
    "BRAZIMIL ZARCÃO 0,75L LATA COR":                  14.00,
    "BRAZIMIL ZARCÃO 3,0L LATA CINZA":                 45.55,
    "BRAZIMIL ZARCÃO  18L LATA COR":                  200.00,
    "BRAZIMIL SELADOR P/ MAD. 0,75L LATA INCOLOR":     19.00,
    "BRAZIMIL SELADOR P/ MAD. 3,0L LATA INCOLOR":      68.00,
    "BRAZIMIL CORRIDA POTE 1KG":                        15.99,
    "BRAZIMIL CORRIDA GALÃO 5KG":                       14.90,
    "BRAZIMIL CORRIDA SACO VALV 10 KG":                11.00,
    "BRAZIMIL CORRIDA BALDE 20KG":                      38.90,
    "BRAZIMIL ACRILICA POTE 1KG":                        6.00,
    "BRAZIMIL ACRILICA GALÃO 5KG":                      19.99,
    "BRAZIMIL ACRILICA SACO VALV 10,0KG":              27.00,
    "BRAZIMIL ACRILICA BALDE 20KG":                    69.90,
    "VEDAMIL / MANTA LÍQUIDA BALDE 3,6KG":            32.50,
    "VEDAMIL / MANTA LÍQUIDA BALDE 12KG":            133.00,
    "VEDAMIL / MANTA LÍQUIDA BALDE 15KG":            100.00,
    "VEDAMIL / MANTA LÍQUIDA BALDE 18KG":            150.00,
    "BRAZIMIL BORRACHA LIQUIDA 3,6 KG":               86.67,
    "BRAZIMIL BORRACHA LIQUIDA 14,0 KG":             285.00,
    "CIMENTO QUEIMADO GALÃO":                          40.00,
    "BRAZIMIL COLA BRANCA 0,5KG":                       7.00,
    "BRAZIMIL COLA BRANCA 1,0KG":                      11.50,
    "AGUARRAS 0,900":                                  15.00,
    "AGUARRAS 0,450":                                   7.80,
}

# Monta lista de produtos
produtos = []
for nome_excel, (cat, sub, cod_time) in tabela.items():
    custo = dados_custo.get(nome_excel)
    peso  = dados_peso.get(nome_excel)
    if not custo or not peso:
        continue
    # Preço: planilha do time > fallback
    preco = None
    if cod_time and cod_time in precos_time:
        preco = precos_time[cod_time]
    if not preco:
        preco = precos_fallback.get(nome_excel)
    if not preco:
        continue
    nome_exib = nome_excel.replace("BRAZIMIL ","").strip().title()
    produtos.append({
        "nome":  nome_exib,
        "cat":   cat,
        "sub":   sub,
        "preco": round(preco, 2),
        "custo": round(custo, 2),
        "peso":  round(peso, 3),
        "fonte": "time" if (cod_time and cod_time in precos_time) else "manual",
    })

print(f"Produtos no dashboard: {len(produtos)}")

produtos_json = json.dumps(produtos, ensure_ascii=False)
data_hoje     = datetime.now().strftime("%d/%m/%Y %H:%M")
total         = len(produtos)


# ── Lê JSONs de vendas e fluxo de caixa (gerados pelo omie_sync.py) ──────
def ler_json(arquivo, fallback):
    caminho = os.path.join(os.path.dirname(os.path.abspath(__file__)), arquivo)
    if os.path.exists(caminho):
        with open(caminho, encoding="utf-8") as f:
            return json.load(f)
    print(f"  [AVISO] {arquivo} não encontrado — usando dados de exemplo.")
    return fallback

vendas_fallback = {
    "gerado_em": "", "total_periodo": 0, "total_semana_atual": 0,
    "variacao_semanal_pct": 0, "semanas": [], "labels_semanas": [],
    "top_produtos": [], "top_clientes": [], "exemplo": True
}
fluxo_fallback = {
    "gerado_em": "", "resumo": {
        "total_carteira": 0, "total_aberto": 0, "total_recebido": 0,
        "total_vencido": 0, "taxa_inadimplencia": 0,
        "qtd_boletos_abertos": 0, "qtd_boletos_vencidos": 0
    },
    "proximas_4_semanas": [], "aging": {}, "top_inadimplentes": [],
    "boletos_detalhes": [], "exemplo": True
}

vendas_data = ler_json("vendas.json", vendas_fallback)
fluxo_data  = ler_json("fluxo_caixa.json", fluxo_fallback)

vendas_json_str = json.dumps(vendas_data, ensure_ascii=False)
fluxo_json_str  = json.dumps(fluxo_data,  ensure_ascii=False)

with open("template.html", "r", encoding="utf-8") as f:
    template = f.read()

html = template.replace("__PRODUTOS_JSON__", produtos_json)
html = html.replace("__DATA_HOJE__", data_hoje)
html = html.replace("__TOTAL__", str(total))
html = html.replace("__VENDAS_JSON__", vendas_json_str)
html = html.replace("__FLUXO_JSON__", fluxo_json_str)

with open("dashboard_brazimil.html", "w", encoding="utf-8") as f:
    f.write(html)

print("=" * 55)
print("Dashboard gerado: dashboard_brazimil.html"
print(f"  Vendas: {'exemplo' if vendas_data.get('exemplo') else 'dados reais'} | Fluxo: {'exemplo' if fluxo_data.get('exemplo') else 'dados reais'}"))
print(f"  Preços da planilha do time: {sum(1 for p in produtos if p['fonte']=='time')} produtos")
print(f"  Preços fallback manual:     {sum(1 for p in produtos if p['fonte']=='manual')} produtos")
print("=" * 55)
