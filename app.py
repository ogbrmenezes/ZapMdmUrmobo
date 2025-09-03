from flask import Flask, render_template, request, jsonify, send_file
import os, re, unicodedata, pandas as pd
from datetime import datetime

APP = Flask(__name__)
ROOT = os.path.dirname(__file__)
UPLOAD = os.path.join(ROOT, "uploads")
os.makedirs(UPLOAD, exist_ok=True)

# Se quiser fixar caminhos por variável de ambiente:
LOJAS_PATH = os.environ.get("LOJAS_PATH")
ROLLOUT_PATH = os.environ.get("ROLLOUT_PATH")

LOG_CSV = os.path.join(UPLOAD, "envios_log.csv")

# -------------------------
# Helpers básicos
# -------------------------

def norm(s):
    """Normaliza string: remove acentos, baixa caixa, colapsa espaços."""
    if s is None:
        return ""
    s = ''.join(ch for ch in unicodedata.normalize('NFKD', str(s)) if not unicodedata.combining(ch))
    return re.sub(r"\s+", " ", s.lower().strip())

def numkey(x):
    """Chave numérica da loja (converte '157.0' -> '157')."""
    if x is None:
        return ""
    xs = str(x).strip()
    try:
        if xs.replace(".0", "").isdigit():
            return str(int(float(xs)))
    except:
        pass
    return xs

def read_excel_auto(path, sheet_name=None, header=0):
    """Lê .xlsx/.xlsb automaticamente (usa pyxlsb p/ .xlsb)."""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".xlsb":
        try:
            return pd.read_excel(path, sheet_name=sheet_name, header=header, engine="pyxlsb")
        except Exception as e:
            raise RuntimeError("Para ler .xlsb, instale 'pyxlsb' (pip install pyxlsb). Erro: %s" % e)
    return pd.read_excel(path, sheet_name=sheet_name, header=header)

def path_lojas():
    if LOJAS_PATH and os.path.exists(LOJAS_PATH): return LOJAS_PATH
    return os.path.join(UPLOAD, "lojas.xlsx")

def path_rollout():
    if ROLLOUT_PATH and os.path.exists(ROLLOUT_PATH): return ROLLOUT_PATH
    return os.path.join(UPLOAD, "rollout.xlsx")

# -------------------------
# Leitura das planilhas
# -------------------------

def load_lojas(path):
    # A planilha tem linhas acima do cabeçalho. Detecta a linha do header.
    df = read_excel_auto(path, sheet_name="DADOS DE LOJAS BRASIL", header=None)
    hdr = None
    for i in range(0, 30):
        if any("Divis" in str(x) for x in df.iloc[i].astype(str).tolist()):
            hdr = i
            break
    if hdr is None:
        raise RuntimeError("Cabeçalho não encontrado na aba DADOS DE LOJAS BRASIL")
    header = df.iloc[hdr].astype(str).tolist()
    df = df.iloc[hdr + 1 :].copy()
    df.columns = [str(c).strip().replace("\n", " ") for c in header]

    def mapc(c):
        k = norm(c)
        if k == "loja": return "loja_numero"
        if k == "nome da loja": return "loja_nome"
        if k == "regional": return "regional"
        if "sub regional" in k: return "sub_regional"
        if "telefone da loja" in k: return "telefone_loja"
        if "celular" in k: return "celular_loja"
        if "endereco" in k: return "endereco"
        if k == "cidade": return "cidade"
        if k == "estado": return "estado"
        return None

    mp = {c: mapc(c) for c in df.columns if mapc(c)}
    df = df.rename(columns=mp)
    keep = [c for c in ["loja_numero", "loja_nome", "regional", "sub_regional",
                        "telefone_loja", "celular_loja", "cidade", "estado", "endereco"]
             if c in df.columns]
    df = df[keep].copy()
    df = df[~(df.get("loja_numero").isna() & df.get("loja_nome").isna())].copy()
    df["key_loja_num"] = df.get("loja_numero", "").map(numkey)
    df["key_loja_nome"] = df.get("loja_nome", "").map(norm)
    return df

def load_rollout(path):
    base = read_excel_auto(path, sheet_name="LOJA", header=2)

    # Identificação (Nº, Loja, Regional, Sub)
    ids = base.iloc[:, 2:6].copy()
    ids.columns = ["loja_numero", "loja_nome", "regional", "sub_regional"]

    # Normaliza nomes (sem acento; minúsculo; sem quebras)
    import unicodedata as _ud
    def _noacc(t):
        return ''.join(ch for ch in _ud.normalize('NFKD', str(t)) if not _ud.combining(ch))
    names = [_noacc(c).lower().replace("\n", " ").strip() for c in base.columns]

    # Localiza colunas pelo nome (podem existir 1 ou 2 ocorrências)
    im  = [i for i, n in enumerate(names) if n == "modelo" or "modelo coletor" in n]
    ifi = [i for i, n in enumerate(names) if "qtde dispositivo" in n and "fisico" in n]
    ia  = [i for i, n in enumerate(names) if "qtde dispositivo" in n and "ativo" in n]
    ip  = [i for i, n in enumerate(names) if "pendentes" in n]
    ipr = [i for i, n in enumerate(names) if "percentual" in n]

    def pick(lst, pos):
        if not lst: return None
        return lst[pos] if pos < len(lst) else lst[-1]

    # Conjunto de colunas do bloco 0 (LOJA) e bloco 1 (RECEB)
    b0 = {"fis": pick(ifi, 0), "ati": pick(ia, 0), "pen": pick(ip, 0), "per": pick(ipr, 0)}
    b1 = {"fis": pick(ifi, 1), "ati": pick(ia, 1), "pen": pick(ip, 1), "per": pick(ipr, 1)}

    def min_defined(dct):
        vals = [v for v in dct.values() if v is not None]
        return min(vals) if vals else None

    # Heurística robusta: o "Modelo" fica à esquerda do primeiro número do bloco.
    def resolve_model_idx(fallback_from_names, block_cols, search_back=3):
        # 1) se já achamos por nome (fallback_from_names), use se for distinto/válido
        if fallback_from_names is not None:
            return fallback_from_names
        # 2) tenta por posição: pega a menor coluna numérica do bloco e volta 1..3 colunas
        right = min_defined(block_cols)  # primeira coluna do bloco (Físico/Ativo/Pendentes/Percentual)
        if right is None:
            return None
        start = max(0, right - search_back)
        # Varre da mais à esquerda para a direita até encostar em 'right'
        best = None
        for j in range(start, right):
            # Preferência: cabeçalho com 'modelo'
            if "modelo" in names[j]:
                return j
            # Alternativa: coluna com maioria de strings (nome de modelo)
            col = base.iloc[:, j]
            # conta quantos são strings não vazios
            non_na = col.dropna()
            if len(non_na) == 0: 
                continue
            sample = non_na.head(10).astype(str).str.strip()
            # se grande parte são strings não-numéricas, é um bom candidato
            str_ratio = (sample.apply(lambda s: not s.replace('.', '', 1).isdigit() and s != "")).mean()
            if str_ratio >= 0.6:
                best = j
        return best

    # Resolve índices de Modelo para cada bloco
    m0_by_name = pick(im, 0)
    m1_by_name = pick(im, 1)
    m0 = resolve_model_idx(m0_by_name, b0, search_back=4)
    m1 = resolve_model_idx(m1_by_name, b1, search_back=4)

    # Se ainda assim o modelo 1 cair no mesmo do 0, puxa a coluna anterior ao 1º número do bloco 1
    if m1 is not None and m0 is not None and m1 == m0:
        right1 = min_defined(b1)
        if right1 is not None and right1 - 1 >= 0:
            m1 = right1 - 1

    # Helper para extrair coluna segura
    def col(i):
        import pandas as _pd
        return base.iloc[:, i] if i is not None and i < base.shape[1] else _pd.Series([None] * len(base))

    out = pd.DataFrame({
        "loja_numero": ids["loja_numero"],
        "loja_nome":    ids["loja_nome"],
        "regional":     ids["regional"],
        "sub_regional": ids["sub_regional"],

        # COLETORES LOJA
        "coletores_loja_modelo":           col(m0),
        "coletores_loja_qtde_fisico":      pd.to_numeric(col(b0["fis"]), errors="coerce"),
        "coletores_loja_qtde_ativo":       pd.to_numeric(col(b0["ati"]), errors="coerce"),
        "coletores_loja_qtde_pendentes":   pd.to_numeric(col(b0["pen"]), errors="coerce"),
        "coletores_loja_percentual":       pd.to_numeric(col(b0["per"]), errors="coerce"),

        # COLETORES - RECEBIMENTO
        "receb_modelo":                    col(m1),
        "receb_qtde_fisico":               pd.to_numeric(col(b1["fis"]), errors="coerce"),
        "receb_qtde_ativo":                pd.to_numeric(col(b1["ati"]), errors="coerce"),
        "receb_qtde_pendentes":            pd.to_numeric(col(b1["pen"]), errors="coerce"),
        "receb_percentual":                pd.to_numeric(col(b1["per"]), errors="coerce"),
    })

    # Chaves & limpeza
    out["key_loja_num"]  = out["loja_numero"].map(numkey)
    out["key_loja_nome"] = out["loja_nome"].map(norm)
    out = out[out["key_loja_num"] != ""].copy()

    # Remove linhas totalmente vazias nos dois blocos
    sec1 = out[["coletores_loja_modelo","coletores_loja_qtde_fisico","coletores_loja_qtde_ativo","coletores_loja_qtde_pendentes","coletores_loja_percentual"]]
    sec2 = out[["receb_modelo","receb_qtde_fisico","receb_qtde_ativo","receb_qtde_pendentes","receb_percentual"]]
    mask_all_empty = (sec1.isna().all(axis=1) & sec2.isna().all(axis=1))
    out = out[~mask_all_empty].copy()

    return out

# -------------------------
# Formatação de mensagem
# -------------------------

def fmt_percent(v):
    """0.75 -> 75% ; 1.0 -> 100% ; '75,86%' -> 76% ; 75 -> 75%"""
    try:
        if v is None:
            return None
        s = str(v).strip().replace('%', '').replace(',', '.')
        if s == "":
            return None
        val = float(s)
        if val <= 1:
            return f"{val * 100:.0f}%"
        return f"{val:.0f}%"
    except:
        return v
def is_complete(percent_value):
    """True quando percentual representa 100% (aceita 1.0, 100, '100%', '100,00%', etc.)."""
    try:
        # Normaliza a string removendo espaços e tratando vírgulas como ponto
        s = str(percent_value).strip().replace('%', '').replace(',', '.')
        
        if not s:
            return False
        
        # Converte para float
        v = float(s)
        
        # Caso o valor seja maior que 1 (por exemplo, '100'), converte para percentual
        if v > 1:
            v = v / 100.0
        
        # Verifica se o valor é 100% ou mais
        return v >= 1.0
    except ValueError:
        # Se o valor não for convertível para float, retorna False
        return False



def _pct_to_float(p):
    """Converte 100, '100%', '1.0', 0.75 -> 100.0, 100.0, 100.0, 75.0"""
    try:
        s = str(p).strip().replace('%','').replace(',', '.')
        if s == '':
            return None
        v = float(s)
        # se vier 0-1, trata como fração
        if v <= 1.0:
            v = v * 100.0
        return v
    except:
        return None

def _modelo_str(v):
    s = str(v).strip()
    return s if s else "(sem modelo)"



def build_section(title, items):
    lines = []
    for label, val in items:
        if val is None:
            continue
        try:
            import pandas as _pd
            if _pd.isna(val):
                continue
        except:
            pass

        # se for Percentual, formata
        lbl_norm = str(label).strip().lower()
        if "percentual" in lbl_norm:
            val = fmt_percent(val)

        s = str(val).strip()
        if s == "" or s.lower() == "nan":
            continue

        lines.append(f"- {label}: {s}")

    if not lines:
        return ""
    return f"{title}\n" + "\n".join(lines) + "\n\n"

def build_msg(row):
    # Bloqueia contato quando já está 100%
    if is_complete(row.get("coletores_loja_percentual")):
        return "Loja já está 100% concluída, não precisa de contato."

    header = (
        "Bom dia, aqui é a Zhaz Soluções e estamos validando como está o andamento dos coletores no MDM.\n"
        f"Loja {row.get('loja_numero','?')} - {row.get('loja_nome','')}\n"
        f"Regional: {row.get('regional','')}\n\n"
    )

    sec_loja = build_section("COLETORES LOJA:", [
        ("Modelo", row.get("coletores_loja_modelo")),
        ("Qtde Dispositivo (Físico)", row.get("coletores_loja_qtde_fisico")),
        ("Qtde Dispositivo Ativo (Urmobo)", row.get("coletores_loja_qtde_ativo")),
        ("Rollout (Pendentes)", row.get("coletores_loja_qtde_pendentes")),
        ("Percentual", fmt_percent(row.get("coletores_loja_percentual"))),
    ])

    sec_receb = build_section("COLETORES - RECEBIMENTO:", [
        ("Modelo", row.get("receb_modelo")),
        ("Qtde Dispositivo (Físico)", row.get("receb_qtde_fisico")),
        ("Qtde Dispositivo Ativo (Urmobo)", row.get("receb_qtde_ativo")),
        ("Rollout (Pendentes)", row.get("receb_qtde_pendentes")),
        ("Percentual", fmt_percent(row.get("receb_percentual"))),
    ])

    return header + sec_loja + sec_receb + "Pode nos retornar como estão os coletores pendentes?"


def build_msg_multi(df_rows):
    """Monta mensagem considerando todos os modelos da loja (Loja e Recebimento)."""
    if df_rows.empty:
        return "Sem dados para a loja."

    # Cabeçalho (pega da primeira linha)
    r0 = df_rows.iloc[0]
    header = (
        "Bom dia, aqui é a Zhaz Soluções e estamos validando como está o andamento dos coletores no MDM URMOBO.\n"
        f"Loja {r0.get('loja_numero','?')} - {r0.get('loja_nome','')}\n"
        f"Regional: {r0.get('regional','')}\n\n"
    )

    pend_lines = []

    # Bloco COLETORES LOJA
    for _, r in df_rows.iterrows():
        modelo = r.get("coletores_loja_modelo")
        p = _pct_to_float(r.get("coletores_loja_percentual"))
        # considera como pendente se existir modelo e p<100
        if pd.notna(modelo) and (p is None or p < 100.0):
            fis = r.get("coletores_loja_qtde_fisico")
            ati = r.get("coletores_loja_qtde_ativo")
            pen = r.get("coletores_loja_qtde_pendentes")
            pend_lines.append(
                f"- COLETORES LOJA • {_modelo_str(modelo)}: {fmt_percent(p)}"
                f" (Físico: {fis or 0}, Ativo: {ati or 0}, Pendentes: {pen if not pd.isna(pen) else '?'} )"
            )

    # Bloco RECEBIMENTO
    for _, r in df_rows.iterrows():
        modelo = r.get("receb_modelo")
        p = _pct_to_float(r.get("receb_percentual"))
        if pd.notna(modelo) and (p is None or p < 100.0):
            fis = r.get("receb_qtde_fisico")
            ati = r.get("receb_qtde_ativo")
            pen = r.get("receb_qtde_pendentes")
            pend_lines.append(
                f"- RECEBIMENTO • {_modelo_str(modelo)}: {fmt_percent(p)}"
                f" (Físico: {fis or 0}, Ativo: {ati or 0}, Pendentes: {pen if not pd.isna(pen) else '?'} )"
            )

    if not pend_lines:
        return f"*** Loja {r0.get('loja_numero','?')} está 100% concluída — não precisa de contato. ***"

    corpo = "⚠️ Pendências encontradas no rollout:\n" + "\n".join(pend_lines)
    return header + corpo + "\n\nPode nos retornar como estão os coletores pendentes?"


# -------------------------
# Rotas
# -------------------------

@APP.route("/")
def index():
    pr = path_rollout(); pl = path_lojas()
    status = {"rollout": os.path.exists(pr), "lojas": os.path.exists(pl), "rollout_path": pr, "lojas_path": pl}
    return render_template("index.html", status=status)

@APP.route("/upload", methods=["POST"])
def upload():
    for tag, name in [("rollout", "rollout.xlsx"), ("lojas", "lojas.xlsx")]:
        f = request.files.get(tag)
        if f and f.filename:
            f.save(os.path.join(UPLOAD, name))
    return jsonify({"ok": True})

@APP.route("/buscar")
def buscar():
    # --- entrada ---
    numero = request.args.get("numero", "").strip()
    if not numero:
        return jsonify({"ok": False, "error": "Informe o número da loja."}), 400

    # --- paths e existência dos arquivos ---
    pr = path_rollout()
    pl = path_lojas()
    if not os.path.exists(pr):
        return jsonify({"ok": False, "error": "Envie a planilha de CONTROLE (rollout) primeiro."}), 400
    if not os.path.exists(pl):
        return jsonify({"ok": False, "error": "Envie a planilha de DADOS DE LOJAS ou defina LOJAS_PATH."}), 400

    # --- carrega planilhas ---
    roll = load_rollout(pr)
    lojas = load_lojas(pl)

    # --- filtra todas as linhas (todos os modelos) da loja ---
    key = numkey(numero)
    df_loja = roll[roll["key_loja_num"] == key].copy()
    if df_loja.empty:
        return jsonify({"ok": False, "error": "Loja não encontrada no rollout (ou sem dados preenchidos)."}), 404

    # --- util local: normaliza NaN -> None pra JSON seguro ---
    def _json_safe_df(df):
        return df.applymap(lambda v: None if (pd.isna(v) if hasattr(pd, "isna") else v is None) else v)

    df_json = _json_safe_df(df_loja)

    # --- verifica se TODOS os modelos (Loja e Recebimento) estão 100% ---
    todos_ok = True
    for _, r in df_loja.iterrows():
        has_loja  = pd.notna(r.get("coletores_loja_modelo"))
        has_receb = pd.notna(r.get("receb_modelo"))

        p1 = _pct_to_float(r.get("coletores_loja_percentual")) if has_loja else 100.0
        p2 = _pct_to_float(r.get("receb_percentual"))          if has_receb else 100.0

        if (p1 is None or p1 < 100.0) or (p2 is None or p2 < 100.0):
            todos_ok = False
            break

    # --- telefone/celular destino ---
    lj = lojas[lojas["key_loja_num"] == key].head(1)
    dest = ""
    if not lj.empty:
        tel = str(lj.iloc[0].get("celular_loja") or lj.iloc[0].get("telefone_loja") or "")
        tel = re.sub(r"\D+", "", tel)
        if tel.startswith("0"):
            tel = tel[1:]
        if tel and not tel.startswith("55"):
            tel = "55" + tel
        dest = tel

    # --- respostas ---
    if todos_ok:
        r0 = df_loja.iloc[0]
        return jsonify({
            "ok": True,
            "dados": df_json.to_dict(orient="records"),
            "mensagem": f"*** Loja {r0.get('loja_numero')} está 100% concluída — não precisa de contato. ***",
            "destinatario": dest,
            "concluida": True
        })

    # Há pendências -> monta mensagem listando apenas os modelos <100%
    msg = build_msg_multi(df_loja)
    return jsonify({
        "ok": True,
        "dados": df_json.to_dict(orient="records"),
        "mensagem": msg,
        "destinatario": dest,
        "concluida": False
    })



def append_log(data):
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    rec = pd.DataFrame([{
        "timestamp": now,
        "loja_numero": data.get("numero",""),
        "loja_nome": data.get("loja_nome",""),
        "regional": data.get("regional",""),
        "destinatario": data.get("destinatario",""),
        "meio": "whatsapp_web",
        "status": "registrado",
        "mensagem": data.get("mensagem","")
    }])
    if os.path.exists(LOG_CSV):
        rec.to_csv(LOG_CSV, mode="a", header=False, index=False, encoding="utf-8")
    else:
        rec.to_csv(LOG_CSV, index=False, encoding="utf-8")

@APP.route("/log", methods=["POST"])
def log_envio():
    data = request.get_json(force=True, silent=True) or {}
    append_log(data)
    return jsonify({"ok": True})

@APP.route("/relatorio.xlsx")
def relatorio():
    if not os.path.exists(LOG_CSV):
        df = pd.DataFrame(columns=["timestamp","loja_numero","loja_nome","regional","destinatario","meio","status","mensagem"])
    else:
        df = pd.read_csv(LOG_CSV, encoding="utf-8")
    out = os.path.join(UPLOAD, "relatorio_envios.xlsx")
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="Envios")
    return send_file(out, as_attachment=True, download_name="relatorio_envios.xlsx")

if __name__ == "__main__":
    APP.run(host="0.0.0.0", port=5000, debug=True)
