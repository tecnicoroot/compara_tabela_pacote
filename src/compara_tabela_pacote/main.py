"""
Compara os pacotes cadastrados no Tasy com a tabela de valores em Excel.

Ligação entre as tabelas:  Tasy.CD_PROCEDIMENTO  =  Excel.ITEM
"""
import re
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl.styles import PatternFill
from robo_utils.banco_dados.bancodadostasy import BancoDadosTasy  # type: ignore[import-untyped]

PASTA_PROJETO = Path(__file__).parent.parent.parent
PASTA_SQL = PASTA_PROJETO / "sql"
PASTA_DADOS = PASTA_PROJETO / "dados"
PASTA_SAIDA = PASTA_PROJETO / "saida"

ARQUIVO_TABELA = PASTA_DADOS / "HUHB - VALORES - LOCAL - INTER NAC - INTER ESTD.xlsx"
LINHA_CABECALHO = 4          # linha do Excel onde estão PRESTADOR, NM PRESTADOR, ITEM...
COLUNAS_EXCEL = "B:I"


# --------------------------------------------------------------------------- #
# 1. Tasy
# --------------------------------------------------------------------------- #
def ler_sql(nome):
    """Lê o .sql tirando ';' e '/' do final (aceita UTF-8 ou CP1252)."""
    dados = (PASTA_SQL / nome).read_bytes()
    try:
        texto = dados.decode("utf-8-sig")
    except UnicodeDecodeError:
        texto = dados.decode("cp1252")
    texto = texto.strip()
    if texto.endswith("/"):
        texto = texto[:-1].strip()
    return texto.rstrip(";").strip()


def colunas_do_select(sql):
    """
    Descobre o nome das colunas a partir do SELECT do arquivo .sql.
    Usa o apelido (alias) quando existe; senão, o nome da coluna (c.ds_convenio -> DS_CONVENIO).
    """
    sql = re.sub(r"--[^\n]*|/\*.*?\*/", " ", sql, flags=re.S)      # tira comentários
    sql = re.sub(r"'[^']*'", "''", sql)                               # ignora textos entre aspas
    inicio = re.search(r"\bselect\b", sql, re.I).end()

    # separa as expressões pelas vírgulas "de fora" dos parênteses, até o FROM principal
    expressoes, atual, nivel, i = [], "", 0, inicio
    while i < len(sql):
        ch = sql[i]
        if ch == "(":
            nivel += 1
        elif ch == ")":
            nivel -= 1
        if nivel == 0 and re.match(r"\bfrom\b", sql[i:], re.I) and not re.match(r"\w", sql[i - 1]):
            break
        if ch == "," and nivel == 0:
            expressoes.append(atual)
            atual = ""
        else:
            atual += ch
        i += 1
    expressoes.append(atual)

    colunas = []
    for e in expressoes:
        nome = re.findall(r"[\w$#]+|\"[^\"]+\"", e.strip())[-1].strip('"')
        colunas.append(nome.upper())
    return colunas


def para_dataframe(resultado, sql):
    """
    Converte o retorno do execSql em DataFrame.
    Aceita DataFrame, lista de dicionários ou lista de tuplas (nomes vêm do SELECT).
    """
    if isinstance(resultado, pd.DataFrame):
        return resultado
    if not resultado:
        return pd.DataFrame(columns=colunas_do_select(sql))
    if isinstance(resultado[0], dict):
        return pd.DataFrame(resultado)
    colunas = colunas_do_select(sql)
    if len(colunas) != len(resultado[0]):
        raise ValueError(
            f"O SELECT tem {len(colunas)} colunas ({', '.join(colunas)}), "
            f"mas o banco devolveu {len(resultado[0])} valores por linha."
        )
    return pd.DataFrame([list(l) for l in resultado], columns=colunas)


def buscar_pacotes_tasy():
    db = BancoDadosTasy()
    sql = ler_sql("busca_pacotes.sql")
    df = para_dataframe(db.execSql(sql), sql)
    df.columns = [str(c).upper() for c in df.columns]
    return df


# --------------------------------------------------------------------------- #
# 2. Excel
# --------------------------------------------------------------------------- #
def ler_tabela_excel(caminho=ARQUIVO_TABELA):
    """Lê todas as abas (HUHB, CDI, GASTRO...) e junta numa tabela só, com a coluna ABA."""
    abas = pd.read_excel(caminho, sheet_name=None, header=LINHA_CABECALHO - 1,
                         usecols=COLUNAS_EXCEL)
    partes = []
    for nome, df in abas.items():
        df = df.dropna(subset=["ITEM"])                 # ignora linhas em branco no fim
        df.insert(0, "ABA", nome)
        partes.append(df)
    return pd.concat(partes, ignore_index=True)


# --------------------------------------------------------------------------- #
# 3. Junção
# --------------------------------------------------------------------------- #
def codigo(serie):
    """Padroniza o código: '80800011', 80800011 e 80800011.0 viram o mesmo número."""
    return pd.to_numeric(serie, errors="coerce").astype("Int64")


# Qual coluna da tabela Excel vale para cada convênio/categoria do Tasy.
# Chave: (convênio, categoria). Categoria None = qualquer categoria daquele convênio.
REGRAS_COLUNA = {
    (3, None): "L",              # Unimed Juiz de Fora  -> Local
    (5, "ESTADUAL"): "E",        # Intercâmbio Estadual
    (5, "NACIONAL"): "N",        # Intercâmbio Nacional
}
COL_VALOR_TASY = "VALOR_PACOTE_TASY"
TOLERANCIA = 0.005               # diferença menor que meio centavo conta como igual

# Comando sugerido para corrigir o valor no Tasy (só nas linhas com diferença)
TABELA_UPDATE = "PACOTE_TIPO_ACOMODACAO"
COLUNA_UPDATE = "VL_PACOTE"      # coluna do valor do pacote (b.VL_PACOTE na consulta)


def coluna_da_regra(convenio, categoria):
    cat = str(categoria).strip().upper() if pd.notna(categoria) else None
    return REGRAS_COLUNA.get((convenio, cat)) or REGRAS_COLUNA.get((convenio, None))


def comando_update(linha):
    """UPDATE que leva o valor do Tasy para o valor da tabela; vazio quando não há o que corrigir."""
    if linha["SITUACAO"] != "Diferente" or pd.isna(linha["NR_SEQUENCIA"]):
        return None
    valor = f"{float(linha['VALOR_TABELA']):.2f}"          # ponto decimal, como o Oracle espera
    nr_seq = int(linha["NR_SEQUENCIA"])
    return (f"update {TABELA_UPDATE} set {COLUNA_UPDATE} = {valor} "
            f"where nr_sequencia = {nr_seq};")


def juntar(tasy, tabela):
    tasy = tasy.copy()
    tabela = tabela.copy()
    tasy["CODIGO"] = codigo(tasy["CD_PROCEDIMENTO"])
    tabela["CODIGO"] = codigo(tabela["ITEM"])

    junto = tasy.merge(tabela, on="CODIGO", how="outer", indicator="ORIGEM")
    junto["ORIGEM"] = junto["ORIGEM"].map(
        {"both": "Nos dois", "left_only": "Só no Tasy", "right_only": "Só na tabela Excel"}
    )

    # valor da tabela conforme convênio/categoria, diferença e situação
    junto["COLUNA_TABELA"] = [
        coluna_da_regra(cv, ct) if o == "Nos dois" else None
        for cv, ct, o in zip(junto["CD_CONVENIO"], junto["CATEGORIA"], junto["ORIGEM"])
    ]
    junto["VALOR_TABELA"] = [
        junto.at[i, col] if isinstance(col, str) else None
        for i, col in junto["COLUNA_TABELA"].items()
    ]
    vl_tasy = pd.to_numeric(junto[COL_VALOR_TASY], errors="coerce")
    vl_tab = pd.to_numeric(junto["VALOR_TABELA"], errors="coerce")
    junto["DIFERENCA"] = (vl_tasy - vl_tab).round(2)

    def situacao(i):
        if junto.at[i, "ORIGEM"] != "Nos dois":
            return junto.at[i, "ORIGEM"]
        if not isinstance(junto.at[i, "COLUNA_TABELA"], str):
            return "Sem regra p/ convênio/categoria"
        return "Igual" if abs(junto.at[i, "DIFERENCA"]) < TOLERANCIA else "Diferente"

    junto["SITUACAO"] = [situacao(i) for i in junto.index]
    junto["COMANDO SUGERIDO"] = [comando_update(junto.loc[i]) for i in junto.index]
    return junto.sort_values(["CODIGO", "CD_CONVENIO", "CATEGORIA"]).reset_index(drop=True)


# --------------------------------------------------------------------------- #
# 4. Saída
# --------------------------------------------------------------------------- #
COLUNAS_COMPARATIVO = [
    "CODIGO", "ABA", "CD_CONVENIO", "CONVENIO", "CATEGORIA", "CLASSIF_ACOMODACAO",
    "NR_SEQUENCIA", "DT_VIGENCIA", "VIGÊNCIA SGU 2.0", COL_VALOR_TASY,
    "COLUNA_TABELA", "VALOR_TABELA", "DIFERENCA", "SITUACAO", "COMANDO SUGERIDO",
]
AMARELO = PatternFill("solid", fgColor="FFF2CC")


def salvar(junto, pasta=PASTA_SAIDA):
    pasta.mkdir(parents=True, exist_ok=True)
    arquivo = pasta / f"comparativo_pacotes_{datetime.now():%Y%m%d_%H%M}.xlsx"
    nos_dois = junto[junto["ORIGEM"] == "Nos dois"]
    cols = [c for c in COLUNAS_COMPARATIVO if c in junto.columns]
    abas = {
        "Comparativo": nos_dois[cols],
        "Diferenças": nos_dois.loc[nos_dois["SITUACAO"] != "Igual", cols],
        "Só no Tasy": junto[junto["ORIGEM"] == "Só no Tasy"],
        "Só na tabela Excel": junto[junto["ORIGEM"] == "Só na tabela Excel"],
    }
    with pd.ExcelWriter(arquivo, engine="openpyxl", datetime_format="DD/MM/YYYY") as xl:
        for nome, df in abas.items():
            if not df.empty:
                df = df.dropna(axis=1, how="all")       # tira colunas vazias de cada aba
            df.to_excel(xl, sheet_name=nome, index=False)
            ws = xl.sheets[nome]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for col in ws.columns:
                largura = max(len(str(c.value)) if c.value is not None else 0 for c in col)
                limite = 95 if col[0].value == "COMANDO SUGERIDO" else 45
                ws.column_dimensions[col[0].column_letter].width = min(largura + 2, limite)
            # destaca as linhas com diferença
            if "SITUACAO" in df.columns:
                pos = list(df.columns).index("SITUACAO") + 1
                for row in ws.iter_rows(min_row=2):
                    if row[pos - 1].value not in ("Igual", None):
                        for c in row:
                            c.fill = AMARELO
    return arquivo


def main():
    tasy = buscar_pacotes_tasy()
    print(f"Tasy: {len(tasy)} linhas")

    caminho = Path(sys.argv[1]) if len(sys.argv) > 1 else ARQUIVO_TABELA
    tabela = ler_tabela_excel(caminho)
    print(f"Tabela Excel: {len(tabela)} itens")

    junto = juntar(tasy, tabela)
    print(junto["SITUACAO"].value_counts().to_string())

    arquivo = salvar(junto)
    print(f"Planilha gerada: {arquivo}")


if __name__ == "__main__":
    main()