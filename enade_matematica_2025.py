"""
ENADE das Licenciaturas 2025 - Matemática (Licenciatura), CO_GRUPO = 702.

Calcula, a partir dos microdados do Inep:
  (A) taxa de acerto por questão objetiva;
  (B) Conceito Enade do curso, 1 a 5 (NT 8/2026);
  (C) Nível de Desempenho individual do aluno (NT 44/2025).

Entradas:
    microdados2025_arq1.txt   -> CO_CURSO, CO_GRUPO, CO_MODALIDADE, CO_CATEGAD
    microdados2025_arq3.txt   -> vetores de resposta, TP_PRES, notas
    mapeamento_questoes.csv   -> posição de cada questão em cada caderno

Saída: planilha única com a aba "Resumo" e as abas de tabelas.
"""

import numpy as np
import pandas as pd
from itertools import combinations
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ==================================================================
# CONFIGURAÇÃO
# ==================================================================
PATH_ARQ1 = "microdados2025_arq1.txt"
PATH_ARQ3 = "microdados2025_arq3.txt"
PATH_MAPEAMENTO = "mapeamento_questoes.csv"
PATH_SAIDA_XLSX = "enade_matematica_2025_tabelas.xlsx"

CO_GRUPO_MATEMATICA = 702
TP_PRES_VALIDO = "555"          # presente com resultado válido
CADERNOS = ["1", "2", "3", "4"]  # cadernos da aplicação regular (o 5 é reaplicação)
TAMANHO_VETOR = 80

CODIGOS_VALIDOS_ITEM = ("0", "1")   # 0 = errou, 1 = acertou
CODIGO_DESCONSIDERADO_TRI = "8"

# Questões a excluir manualmente, além das desconsideradas pela TRI
# (que são detectadas automaticamente na base).
QUESTOES_EXCLUIDAS_ADICIONAIS = []

# NT 44/2025, item 9.3: nível de desempenho individual a partir da NT_GER.
LIMIAR_BASICO = 50
LIMIAR_ADEQUADO = 70
ORDEM_NIVEIS = ["Abaixo do Básico", "Básico", "Adequado"]

# NT 8/2026, item 7.4: conversão do PDB_ck em Conceito Enade.
FAIXAS_CONCEITO_CURSO = [(0.400, 1), (0.600, 2), (0.750, 3), (0.900, 4)]
ORDEM_CONCEITOS = [1, 2, 3, 4, 5]
MIN_CONCLUINTES_VALIDOS_POR_CURSO = 2   # NT 8/2026, itens 5.1 e 6.6

# Questões analisadas uma a uma nas Abas 9, 10 e 11. Devem estar entre Q1 e
# Q80 e não podem ser questões excluídas da análise por item.
QUESTOES_SELECIONADAS = ["Q48", "Q54", "Q67"]

# Abas posicionadas no final da planilha, nesta ordem. Os nomes devem ser
# idênticos aos usados em registra_tabela. As demais abas seguem a ordem
# em que são geradas, e a aba "Resumo" é sempre a primeira.
ABAS_NO_FINAL = [
    "Aba 9 - Questões em foco",
    "Aba 10 - Conceito x acerto",
    "Aba 11 - Alternativas marcadas",
]

# NT 44/2025, item 4.1.1: 30 itens de Formação Geral + 50 de Componente Específico.
BLOCO_GERAL = "Formação Geral Docente (Q1-Q30)"
BLOCO_ESPECIFICO = "Componente Específico (Q31-Q80)"

ALTERNATIVAS_OBJETIVAS = ["A", "B", "C", "D"]
CODIGOS_ESPECIAIS_ESCOLHA = {".": "Em branco (.)", "*": "Múltiplas escolhas (*)"}
ORDEM_ALTERNATIVAS = ALTERNATIVAS_OBJETIVAS + list(CODIGOS_ESPECIAIS_ESCOLHA.values())

# Faixas de % de acerto individual da Tabela 7: fechadas à esquerda,
# abertas à direita; a última inclui 100%.
LARGURA_FAIXA_TABELA7 = 20
ROTULOS_FAIXA_TABELA7 = [
    f"[{inicio}%, {inicio + LARGURA_FAIXA_TABELA7}%)"
    for inicio in range(0, 100 - LARGURA_FAIXA_TABELA7, LARGURA_FAIXA_TABELA7)
] + [f"[{100 - LARGURA_FAIXA_TABELA7}%, 100%]"]
LIMITES_FAIXA_TABELA7 = (
    [-0.001] + list(range(LARGURA_FAIXA_TABELA7, 100, LARGURA_FAIXA_TABELA7)) + [100.001]
)

# Conjuntos de questões da Tabela 7 (Aba 7). Cada conjunto gera três
# tabelas, uma por Nível de Desempenho. O valor de cada entrada é:
#   - BLOCO_ESPECIFICO ou BLOCO_GERAL: todas as questões do bloco;
#   - uma lista de questões, que pode ser editada livremente.
# Questões fora da análise por item (por exemplo, as desconsideradas
# pela TRI) são descartadas, com aviso no console.
CONJUNTOS_TABELA7 = {
    "todas as questões do Componente Específico (Q31 a Q80)": BLOCO_ESPECIFICO,
    "todas as questões do Componente de Formação Geral (Q1 a Q30)": BLOCO_GERAL,
    "questões definidas pelo usuário": [
        "Q32", "Q34", "Q38", "Q42", "Q47", "Q48", "Q50",
        "Q51", "Q54", "Q62", "Q65", "Q66", "Q67", "Q68", "Q69", "Q70",
        "Q71", "Q72", "Q75", "Q76", "Q77", "Q78", "Q79", "Q80",
    ],
}

NOMES_TABELAS = {
    "1": "taxa de acerto por questão objetiva",
    "2": "Conceito Enade por curso (PDB_ck e conceito 1-5)",
    "3": "cursos 'Sem Conceito (SC)'",
    "4": "concentração de cursos e estudantes por Conceito Enade do curso",
    "5": "alunos por Nível de Desempenho individual",
    "6": "acerto de todas as questões por Conceito Enade do curso (anexo)",
    "7": (
        "distribuição do % de acerto individual de cada aluno, por Nível "
        "de Desempenho individual"
    ),
    "8": "acerto de todas as questões por Nível de Desempenho individual (anexo)",
    "9": f"acerto de {', '.join(QUESTOES_SELECIONADAS)} por Nível de Desempenho individual",
    "10": f"acerto de {', '.join(QUESTOES_SELECIONADAS)} por Conceito Enade do curso",
    "11": "contingência: Conceito Enade do curso x padrão de acerto",
    "13": (
        f"alternativa marcada pelos candidatos em {', '.join(QUESTOES_SELECIONADAS)} "
        "(quantidade e % do total de respostas validadas)"
    ),
}

# ==================================================================
# FILA DE EXPORTAÇÃO
# ==================================================================
# Cada aba guarda uma lista de blocos; blocos registrados sob o mesmo
# nome de aba são empilhados verticalmente nela, na ordem de registro.
BLOCOS_POR_ABA = {}
ORDEM_ABAS = []


def registra_tabela(nome_aba, df, com_indice=False, titulo=None, colunas_percentual=None):
    if colunas_percentual is None:
        colunas_percentual = [str(c) for c in df.columns if "%" in str(c)]
    if nome_aba not in BLOCOS_POR_ABA:
        BLOCOS_POR_ABA[nome_aba] = []
        ORDEM_ABAS.append(nome_aba)
    BLOCOS_POR_ABA[nome_aba].append(
        {
            "df": df,
            "com_indice": com_indice,
            "titulo": titulo,
            "colunas_percentual": colunas_percentual,
        }
    )


def percentual(parte, total):
    """Percentual arredondado a 2 casas, com proteção contra total zero."""
    if isinstance(total, (int, float, np.integer, np.floating)) and not total:
        return parte * 0.0
    return (parte / total * 100).round(2)


def acumulada(parte, total):
    """Percentual acumulado: acumula antes de arredondar, para que a
    última linha feche exatamente em 100,00."""
    if isinstance(total, (int, float, np.integer, np.floating)) and not total:
        return parte * 0.0
    return (parte.cumsum() / total * 100).round(2)


# ==================================================================
# 1) ARQ1 - cursos de Matemática e suas características
# ==================================================================
arq1 = pd.read_csv(PATH_ARQ1, sep=";", encoding="latin1")
matematica = arq1[arq1["CO_GRUPO"] == CO_GRUPO_MATEMATICA]
cursos_matematica = set(matematica["CO_CURSO"].unique())
print(f"Cursos de Matemática (Licenciatura) encontrados: {len(cursos_matematica)}")

# CO_MODALIDADE: 0 = EaD, 1 = Presencial.
# CO_CATEGAD: 1-3 = pública (federal/estadual/municipal); 4, 5, 7 e 8 = privada.
CATEGAD_PUBLICA = {1, 2, 3}

info_cursos = (
    matematica[["CO_CURSO", "CO_MODALIDADE", "CO_CATEGAD"]]
    .drop_duplicates(subset="CO_CURSO")
    .set_index("CO_CURSO")
)
info_cursos["modalidade_ensino"] = np.where(
    info_cursos["CO_MODALIDADE"] == 0, "EAD", "Presencial"
)
info_cursos["tipo_instituicao"] = np.where(
    info_cursos["CO_CATEGAD"].isin(CATEGAD_PUBLICA), "Pública", "Privada"
)
print(
    "Por modalidade: "
    + ", ".join(f"{k}={v}" for k, v in info_cursos["modalidade_ensino"].value_counts().items())
)
print(
    "Por tipo de instituição: "
    + ", ".join(f"{k}={v}" for k, v in info_cursos["tipo_instituicao"].value_counts().items())
)

# ==================================================================
# 2) ARQ3 - registros dos estudantes desses cursos
# ==================================================================
colunas_necessarias = [
    "CO_CURSO", "CO_CADERNO", "DS_VT_GAB_OBJ", "DS_VT_ESC_OBJ", "DS_VT_ACE_OBJ",
    "TP_PRES", "NT_OBJ", "NT_DIS", "NT_GER",
]

pedacos = []
for pedaco in pd.read_csv(
    PATH_ARQ3, sep=";", encoding="latin1",
    usecols=colunas_necessarias, dtype=str, chunksize=200_000,
):
    pedaco["CO_CURSO"] = pedaco["CO_CURSO"].astype(int)
    pedaco = pedaco[pedaco["CO_CURSO"].isin(cursos_matematica)]
    if len(pedaco):
        pedacos.append(pedaco)

base_matematica = pd.concat(pedacos, ignore_index=True)

# Único filtro de validade em nível de estudante (NT 8/2026, item 6.2).
# Os demais códigos de TP_PRES - inclusive o ".", dos estudantes que
# optaram por prestar a prova em outra área - ficam de fora.
base_valida = base_matematica[base_matematica["TP_PRES"] == TP_PRES_VALIDO].copy()
print(f"Registros com resultado válido (TP_PRES=555): {len(base_valida)}")

# ==================================================================
# 3) Base para as análises por item
# ==================================================================
# Exclui o caderno 5 (reaplicação: prova distinta, sem mapeamento de
# posições) e os registros cujos vetores vieram em branco.
base_itens = base_valida[
    base_valida["CO_CADERNO"].isin(CADERNOS)
    & (base_valida["DS_VT_ACE_OBJ"].str.len() == TAMANHO_VETOR)
].copy()
print(
    f"Registros usados nas análises por item: {len(base_itens)} "
    f"({len(base_valida) - len(base_itens)} fora - caderno 5 ou vetor ausente)"
)

mapeamento = pd.read_csv(PATH_MAPEAMENTO, sep=";").set_index("questao")
TODAS_QUESTOES = list(mapeamento.index)


def coluna_canonica(questao: str) -> int:
    """Índice (0-based) da questão nos vetores em ordem canônica."""
    return int(questao[1:]) - 1


def para_matriz(serie: pd.Series) -> np.ndarray:
    """Converte uma coluna de vetores de 80 caracteres em matriz
    (linhas = estudantes, colunas = posições do vetor)."""
    return np.array(serie.map(list).tolist())


# DS_VT_ACE_OBJ já vem em ordem canônica (a posição de uma questão não
# muda com o caderno). DS_VT_ESC_OBJ e DS_VT_GAB_OBJ vêm na ordem do
# caderno do estudante e são reordenados aqui para a ordem canônica,
# de modo que todas as matrizes fiquem indexadas do mesmo jeito.
ACE = para_matriz(base_itens["DS_VT_ACE_OBJ"])
ESC_CADERNO = para_matriz(base_itens["DS_VT_ESC_OBJ"])
GAB_CADERNO = para_matriz(base_itens["DS_VT_GAB_OBJ"])
caderno_por_linha = base_itens["CO_CADERNO"].to_numpy()

ESC = np.empty_like(ESC_CADERNO)
GAB = np.empty_like(GAB_CADERNO)
for cad in CADERNOS:
    linhas = caderno_por_linha == cad
    permutacao = [int(mapeamento.loc[q, f"caderno{cad}"]) - 1 for q in TODAS_QUESTOES]
    ESC[linhas] = ESC_CADERNO[linhas][:, permutacao]
    GAB[linhas] = GAB_CADERNO[linhas][:, permutacao]

VALIDO = np.isin(ACE, CODIGOS_VALIDOS_ITEM)
ACERTO = ACE == "1"

# Verificação de integridade: o acerto registrado pelo Inep tem de
# coincidir com "marcou a letra do gabarito". Se não coincidir, ou o
# mapeamento de cadernos ou a hipótese de ordenação está errada.
concordancia = ((ESC == GAB)[VALIDO] == ACERTO[VALIDO]).mean()
print(f"Concordância entre DS_VT_ACE_OBJ e (DS_VT_ESC_OBJ == DS_VT_GAB_OBJ): {concordancia:.5%}")
assert concordancia > 0.999, (
    "Os vetores de escolha/gabarito não reproduzem o vetor de acertos. "
    "Verifique mapeamento_questoes.csv e a ordenação dos vetores."
)

# Questões desconsideradas pela TRI (código 8 para todos os
# estudantes): excluídas de todas as análises por item.
questoes_desconsideradas = [
    q for q in TODAS_QUESTOES
    if (ACE[:, coluna_canonica(q)] == CODIGO_DESCONSIDERADO_TRI).all()
]
QUESTOES_EXCLUIDAS = sorted(
    set(questoes_desconsideradas) | set(QUESTOES_EXCLUIDAS_ADICIONAIS),
    key=lambda q: int(q[1:]),
)
QUESTOES = [q for q in TODAS_QUESTOES if q not in QUESTOES_EXCLUIDAS]
print(
    f"Questões excluídas das análises por item ({len(QUESTOES_EXCLUIDAS)}): "
    f"{', '.join(QUESTOES_EXCLUIDAS) or 'nenhuma'}"
)


def classifica_bloco_questao(questao: str) -> str:
    return BLOCO_GERAL if int(questao[1:]) <= 30 else BLOCO_ESPECIFICO


QUESTOES_ESPECIFICO = [q for q in QUESTOES if classifica_bloco_questao(q) == BLOCO_ESPECIFICO]
COLUNAS_ESPECIFICO = [coluna_canonica(q) for q in QUESTOES_ESPECIFICO]

LINHAS_POR_CADERNO = {cad: (caderno_por_linha == cad) for cad in CADERNOS}


def conta_acertos(linhas, questao):
    """(nº de respostas válidas, nº de acertos) da questão entre as
    linhas selecionadas."""
    coluna = coluna_canonica(questao)
    validos = VALIDO[linhas, coluna]
    return int(validos.sum()), int((ACERTO[linhas, coluna] & validos).sum())


def linha_taxa_acerto(linhas, questao, prefixo):
    """Acertos, erros e taxa de acerto de uma questão, nomeados com o
    prefixo do grupo (usado nas Tabelas 6, 8, 9 e 10)."""
    n, acertos = conta_acertos(linhas, questao)
    taxa = round(acertos / n * 100, 2) if n else float("nan")
    return {
        f"{prefixo}_acertos": acertos,
        f"{prefixo}_erros": n - acertos,
        f"{prefixo}_taxa_acerto_%": taxa,
    }


# ==================================================================
# 4) Aba "Resumo"
# ==================================================================
resumo_tabelas = pd.DataFrame(
    [
        {"Tabela": "1", "Aba": "Aba 1 - Acerto por questão",
         "O que mede": "Taxa de acerto de cada questão objetiva, por caderno de prova e no total, "
                       "com o número de respostas válidas e de acertos em cada caso"},
        {"Tabela": "2", "Aba": "Aba 2 - Conceito Enade do curso",
         "O que mede": "Conceito Enade (1 a 5) de cada curso, calculado a partir do percentual de "
                       "concluintes com desempenho igual ou acima do Básico (PDB_ck)"},
        {"Tabela": "3", "Aba": "Aba 3 - Cursos sem conceito",
         "O que mede": "Cursos classificados como 'Sem Conceito (SC)' por terem menos de 2 "
                       "concluintes com resultado válido"},
        {"Tabela": "4", "Aba": "Aba 4 - Cursos x Conceito Enade",
         "O que mede": "Quantos cursos e quantos estudantes estão em cada faixa de Conceito Enade, "
                       "no total e nos recortes por modalidade e categoria administrativa"},
        {"Tabela": "5", "Aba": "Aba 5 - Nível individual",
         "O que mede": "Quantos estudantes foram classificados em cada Nível de Desempenho "
                       "individual, no total e nos mesmos recortes"},
        {"Tabela": "6", "Aba": "Aba 6 - Questões x Enade curso",
         "O que mede": "Taxa de acerto de todas as questões objetivas, quebrada pelo Conceito "
                       "Enade do curso do aluno"},
        {"Tabela": "7", "Aba": "Aba 7 - Acerto individual",
         "O que mede": "Distribuição dos alunos em faixas fixas de % de acerto individual, "
                       "separada por Nível de Desempenho, para o Componente Específico, para "
                       "o Componente de Formação Geral e para um conjunto de questões "
                       "definido pelo usuário"},
        {"Tabela": "8", "Aba": "Aba 8 - Questões x Individual",
         "O que mede": "Taxa de acerto de todas as questões objetivas, quebrada pelo Nível de "
                       "Desempenho individual do aluno"},
        {"Tabela": "9", "Aba": "Aba 9 - Questões em foco",
         "O que mede": f"Taxa de acerto de {', '.join(QUESTOES_SELECIONADAS)} pelo Nível de "
                       "Desempenho individual do aluno que respondeu"},
        {"Tabela": "10", "Aba": "Aba 9 - Questões em foco",
         "O que mede": f"Taxa de acerto de {', '.join(QUESTOES_SELECIONADAS)} pelo Conceito Enade "
                       "do curso em que o aluno está matriculado"},
        {"Tabela": "11", "Aba": "Aba 10 - Conceito x acerto",
         "O que mede": f"Conceito Enade do curso cruzado com o padrão de acerto do aluno em "
                       f"{', '.join(QUESTOES_SELECIONADAS)}, em contagem de alunos"},
        {"Tabela": "12", "Aba": "Aba 10 - Conceito x acerto",
         "O que mede": "Mesma contingência da Tabela 11, em percentual dentro de cada linha "
                       "(Conceito Enade do curso)"},
        {"Tabela": "13", "Aba": "Aba 11 - Alternativas marcadas",
         "O que mede": f"Distribuição das alternativas marcadas em {', '.join(QUESTOES_SELECIONADAS)}, "
                       "entre os candidatos com resposta validada no item"},
    ]
).set_index("Tabela")

# O Resumo segue a mesma ordem das abas na planilha.
resumo_tabelas = resumo_tabelas.iloc[sorted(
    range(len(resumo_tabelas)),
    key=lambda i: (
        ABAS_NO_FINAL.index(resumo_tabelas["Aba"].iloc[i]) + 1
        if resumo_tabelas["Aba"].iloc[i] in ABAS_NO_FINAL else 0
    ),
)]

registra_tabela(
    "Resumo", resumo_tabelas, com_indice=True,
    titulo="Resumo — o que cada tabela desta planilha mede",
)

# ==================================================================
# 5) TABELA 1 - taxa de acerto por questão
# ==================================================================
linhas_tabela1 = []
for questao in QUESTOES:
    registro = {"questao": questao, "posicao_vetor_ACE_OBJ": coluna_canonica(questao) + 1}
    n_total = acertos_total = 0
    for cad in CADERNOS:
        n, acertos = conta_acertos(LINHAS_POR_CADERNO[cad], questao)
        registro[f"posicao_caderno{cad}"] = int(mapeamento.loc[questao, f"caderno{cad}"])
        registro[f"N_valido_caderno{cad}"] = n
        registro[f"acertos_caderno{cad}"] = acertos
        registro[f"taxa_acerto_caderno{cad}_%"] = round(acertos / n * 100, 2) if n else float("nan")
        n_total += n
        acertos_total += acertos
    registro["N_valido_TOTAL"] = n_total
    registro["acertos_TOTAL"] = acertos_total
    registro["taxa_acerto_TOTAL_%"] = round(acertos_total / n_total * 100, 2) if n_total else float("nan")
    linhas_tabela1.append(registro)

ordem_colunas = ["questao", "posicao_vetor_ACE_OBJ"]
for cad in CADERNOS:
    ordem_colunas += [
        f"posicao_caderno{cad}", f"N_valido_caderno{cad}",
        f"acertos_caderno{cad}", f"taxa_acerto_caderno{cad}_%",
    ]
ordem_colunas += ["N_valido_TOTAL", "acertos_TOTAL", "taxa_acerto_TOTAL_%"]
tabela1 = pd.DataFrame(linhas_tabela1)[ordem_colunas]

print(f"\n=== TABELA 1 — {NOMES_TABELAS['1']} (primeiras linhas) ===")
print(tabela1.head(10).to_string(index=False))
registra_tabela("Aba 1 - Acerto por questão", tabela1, titulo=f"Tabela 1 — {NOMES_TABELAS['1']}")

# ==================================================================
# 6) Notas por aluno - base do conceito do curso e do nível individual
# ==================================================================
base_notas = base_valida.copy()
base_notas["NT_GER_num"] = pd.to_numeric(
    base_notas["NT_GER"].str.replace(",", ".", regex=False), errors="coerce"
)
n_sem_nota = int(base_notas["NT_GER_num"].isna().sum())
if n_sem_nota:
    print(
        f"\nAviso: {n_sem_nota} registros com TP_PRES=555 não têm NT_GER preenchida "
        f"e ficam fora do conceito do curso e do nível individual."
    )
base_notas = base_notas.dropna(subset=["NT_GER_num"]).copy()

base_notas["nivel_individual"] = pd.cut(
    base_notas["NT_GER_num"],
    bins=[-np.inf, LIMIAR_BASICO, LIMIAR_ADEQUADO, np.inf],
    labels=ORDEM_NIVEIS,
    right=False,
)
base_notas["acima_do_basico"] = base_notas["NT_GER_num"] >= LIMIAR_BASICO

# ==================================================================
# 7) TABELAS 2 e 3 - Conceito Enade do curso
# ==================================================================
def classifica_conceito_curso(pdb):
    for limite_superior, conceito in FAIXAS_CONCEITO_CURSO:
        if pdb < limite_superior:
            return conceito
    return 5


por_curso = base_notas.groupby("CO_CURSO").agg(
    NC_ck=("NT_GER_num", "size"),
    NCDB_ck=("acima_do_basico", "sum"),
).reset_index()
por_curso["PDB_ck"] = (por_curso["NCDB_ck"] / por_curso["NC_ck"]).round(3)

cursos_com_conceito = por_curso[por_curso["NC_ck"] >= MIN_CONCLUINTES_VALIDOS_POR_CURSO].copy()
cursos_com_conceito["conceito_enade_curso"] = cursos_com_conceito["PDB_ck"].apply(
    classifica_conceito_curso
)

# "Sem Conceito" é obtido por diferença sobre o universo completo de
# cursos do arq1: cursos sem nenhum registro válido não formam grupo
# no groupby e ficariam de fora das duas listas.
cursos_sem_conceito = cursos_matematica - set(cursos_com_conceito["CO_CURSO"])
tabela3 = pd.DataFrame({"CO_CURSO": sorted(cursos_sem_conceito)})
tabela3["concluintes_validos_TP_PRES_555"] = (
    tabela3["CO_CURSO"].map(base_valida.groupby("CO_CURSO").size()).fillna(0).astype(int)
)

print(f"\nCursos com Conceito Enade calculado: {len(cursos_com_conceito)}")
print(f"Cursos 'Sem Conceito' (SC): {len(cursos_sem_conceito)}")
assert len(cursos_com_conceito) + len(cursos_sem_conceito) == len(cursos_matematica), (
    "A soma de cursos com e sem conceito não fecha com o total de cursos do arq1."
)

registra_tabela(
    "Aba 2 - Conceito Enade do curso", cursos_com_conceito,
    titulo=f"Tabela 2 — {NOMES_TABELAS['2']}",
)
registra_tabela(
    "Aba 3 - Cursos sem conceito", tabela3,
    titulo=f"Tabela 3 — {NOMES_TABELAS['3']}",
)

# ==================================================================
# 8) TABELA 4 - distribuição dos cursos por Conceito Enade
# ==================================================================
# Base de percentuais: apenas os cursos COM conceito; cursos "SC" não
# entram na distribuição.
def monta_tabela4(cursos_subset: pd.DataFrame) -> pd.DataFrame:
    tabela = (
        cursos_subset.groupby("conceito_enade_curso")
        .agg(numero_de_cursos=("CO_CURSO", "size"), total_de_estudantes=("NC_ck", "sum"))
        .reindex(ORDEM_CONCEITOS, fill_value=0)
    )
    total_cursos = tabela["numero_de_cursos"].sum()
    total_estudantes = tabela["total_de_estudantes"].sum()
    tabela["porcentagem_cursos_%"] = percentual(tabela["numero_de_cursos"], total_cursos)
    tabela["porcentagem_cursos_acumulada_%"] = acumulada(tabela["numero_de_cursos"], total_cursos)
    tabela["porcentagem_estudantes_%"] = percentual(tabela["total_de_estudantes"], total_estudantes)
    tabela["porcentagem_estudantes_acumulada_%"] = acumulada(
        tabela["total_de_estudantes"], total_estudantes
    )
    tabela["numero_medio_estudantes_por_curso"] = (
        tabela["total_de_estudantes"]
        / tabela["numero_de_cursos"].where(tabela["numero_de_cursos"] != 0)
    ).round(2)
    tabela.index.name = "Conceito Enade do curso"
    return tabela


cursos_com_conceito_seg = cursos_com_conceito.merge(
    info_cursos[["modalidade_ensino", "tipo_instituicao"]],
    left_on="CO_CURSO", right_index=True, how="left",
)


def define_recortes(df: pd.DataFrame):
    """Recortes por modalidade e tipo de instituição, aplicáveis tanto à
    base de cursos (Tabela 4) quanto à base de alunos (Tabela 5)."""
    modalidade = df["modalidade_ensino"]
    tipo = df["tipo_instituicao"]
    return [
        ("EAD", "modalidade: EAD", modalidade == "EAD"),
        ("Presencial", "modalidade: Presencial", modalidade == "Presencial"),
        ("Pública", "tipo de instituição: Pública", tipo == "Pública"),
        ("Privada", "tipo de instituição: Privada", tipo == "Privada"),
        ("Pública sem EAD", "Pública, apenas Presencial",
         (tipo == "Pública") & (modalidade == "Presencial")),
        ("Privada sem EAD", "Privada, apenas Presencial",
         (tipo == "Privada") & (modalidade == "Presencial")),
        ("Pública sem Presencial", "Pública, apenas EAD",
         (tipo == "Pública") & (modalidade == "EAD")),
        ("Privada sem Presencial", "Privada, apenas EAD",
         (tipo == "Privada") & (modalidade == "EAD")),
    ]


NOME_ABA_TABELA4 = "Aba 4 - Cursos x Conceito Enade"
tabela4 = monta_tabela4(cursos_com_conceito)
print(f"\n=== TABELA 4 — {NOMES_TABELAS['4']} — Geral ===")
print(tabela4.to_string())
registra_tabela(
    NOME_ABA_TABELA4, tabela4, com_indice=True,
    titulo=f"Tabela 4 (Geral) — {NOMES_TABELAS['4']}",
)

for rotulo, descricao, filtro in define_recortes(cursos_com_conceito_seg):
    tabela4_recorte = monta_tabela4(cursos_com_conceito_seg[filtro])
    print(f"\n=== TABELA 4 — {NOMES_TABELAS['4']} — {rotulo} ===")
    print(tabela4_recorte.to_string())
    registra_tabela(
        NOME_ABA_TABELA4, tabela4_recorte, com_indice=True,
        titulo=f"Tabela 4 ({rotulo}) — {NOMES_TABELAS['4']} — recorte: {descricao}",
    )

# ==================================================================
# 9) TABELA 5 - distribuição dos alunos por Nível de Desempenho
# ==================================================================
def monta_tabela5(base_subset: pd.DataFrame) -> pd.DataFrame:
    tabela = (
        base_subset["nivel_individual"]
        .value_counts()
        .reindex(ORDEM_NIVEIS, fill_value=0)
        .to_frame("numero_de_alunos")
    )
    total = tabela["numero_de_alunos"].sum()
    tabela["pct_alunos_com_nivel_%"] = percentual(tabela["numero_de_alunos"], total)
    tabela["pct_acumulada_%"] = acumulada(tabela["numero_de_alunos"], total)
    tabela.index.name = "Nível de Desempenho"
    return tabela


NOME_ABA_TABELA5 = "Aba 5 - Nível individual"
tabela5 = monta_tabela5(base_notas)
print(f"\n=== TABELA 5 — {NOMES_TABELAS['5']} — Geral ===")
print(tabela5.to_string())
registra_tabela(
    NOME_ABA_TABELA5, tabela5, com_indice=True,
    titulo=f"Tabela 5 (Geral) — {NOMES_TABELAS['5']}",
)

base_notas_seg = base_notas.merge(
    info_cursos[["modalidade_ensino", "tipo_instituicao"]],
    left_on="CO_CURSO", right_index=True, how="left",
)

for rotulo, descricao, filtro in define_recortes(base_notas_seg):
    tabela5_recorte = monta_tabela5(base_notas_seg[filtro])
    print(f"\n=== TABELA 5 — {NOMES_TABELAS['5']} — {rotulo} ===")
    print(tabela5_recorte.to_string())
    registra_tabela(
        NOME_ABA_TABELA5, tabela5_recorte, com_indice=True,
        titulo=f"Tabela 5 ({rotulo}) — {NOMES_TABELAS['5']} — recorte: {descricao}",
    )

# ==================================================================
# 10) Cruzamento das respostas por item com nível e conceito
# ==================================================================
# base_itens e base_notas vêm ambas de base_valida, então compartilham o
# índice original do arq3 e podem ser casadas por ele. O nível é do
# próprio aluno; o conceito é o do curso em que ele está matriculado.
nivel_por_linha = base_notas["nivel_individual"].reindex(base_itens.index)
LINHAS_POR_NIVEL = {
    nivel: (nivel_por_linha == nivel).to_numpy() for nivel in ORDEM_NIVEIS
}

conceito_por_linha = base_itens["CO_CURSO"].map(
    cursos_com_conceito.set_index("CO_CURSO")["conceito_enade_curso"]
)
LINHAS_POR_CONCEITO = {
    conceito: (conceito_por_linha == conceito).to_numpy() for conceito in ORDEM_CONCEITOS
}

n_sem_conceito = int(conceito_por_linha.isna().sum())
if n_sem_conceito:
    print(
        f"\nAviso: {n_sem_conceito} estudantes pertencem a cursos 'Sem Conceito (SC)' "
        f"e ficam fora das tabelas que agrupam por Conceito Enade do curso."
    )

# ==================================================================
# 11) TABELAS 9 e 10 - acerto das questões selecionadas
# ==================================================================
NOME_ABA_TABELA9_10 = "Aba 9 - Questões em foco"

tabela9 = pd.DataFrame([
    {"Nível de Desempenho": nivel,
     **{k: v for q in QUESTOES_SELECIONADAS
        for k, v in linha_taxa_acerto(LINHAS_POR_NIVEL[nivel], q, q).items()}}
    for nivel in ORDEM_NIVEIS
]).set_index("Nível de Desempenho")

print(f"\n=== TABELA 9 — {NOMES_TABELAS['9']} ===")
print(tabela9.to_string())
registra_tabela(
    NOME_ABA_TABELA9_10, tabela9, com_indice=True,
    titulo=f"Tabela 9 — {NOMES_TABELAS['9']}",
)

tabela10 = pd.DataFrame([
    {"Conceito Enade do curso": conceito,
     **{k: v for q in QUESTOES_SELECIONADAS
        for k, v in linha_taxa_acerto(LINHAS_POR_CONCEITO[conceito], q, q).items()}}
    for conceito in ORDEM_CONCEITOS
]).set_index("Conceito Enade do curso")

print(f"\n=== TABELA 10 — {NOMES_TABELAS['10']} ===")
print(tabela10.to_string())
registra_tabela(
    NOME_ABA_TABELA9_10, tabela10, com_indice=True,
    titulo=f"Tabela 10 — {NOMES_TABELAS['10']}",
)

# ==================================================================
# 12) TABELAS 11 e 12 - contingência conceito x padrão de acerto
# ==================================================================
# Só entra o aluno com código válido em todas as questões selecionadas
# ao mesmo tempo e cujo curso tem Conceito Enade.
NOME_ABA_TABELA11_12 = "Aba 10 - Conceito x acerto"
colunas_selecionadas = [coluna_canonica(q) for q in QUESTOES_SELECIONADAS]
linhas_contingencia = VALIDO[:, colunas_selecionadas].all(axis=1) & conceito_por_linha.notna().to_numpy()

n_excluidos = int((~VALIDO[:, colunas_selecionadas].all(axis=1)).sum())
if n_excluidos:
    print(
        f"\nAviso: {n_excluidos} estudantes não têm código válido nas "
        f"{len(QUESTOES_SELECIONADAS)} questões simultaneamente e ficam fora da Tabela 11."
    )

padrao_acerto = [
    "{" + ", ".join(q for q, acertou in zip(QUESTOES_SELECIONADAS, acertos) if acertou) + "}"
    if any(acertos) else "∅"
    for acertos in ACERTO[np.ix_(linhas_contingencia, colunas_selecionadas)]
]

ORDEM_PADROES = ["∅"] + [
    "{" + ", ".join(combo) + "}"
    for tamanho in range(1, len(QUESTOES_SELECIONADAS) + 1)
    for combo in combinations(QUESTOES_SELECIONADAS, tamanho)
]

tabela11 = pd.crosstab(
    pd.Series(
        conceito_por_linha[linhas_contingencia].astype(int).to_numpy(),
        name="Conceito Enade do curso",
    ),
    pd.Series(padrao_acerto, name="padrao_acerto"),
).reindex(index=ORDEM_CONCEITOS, columns=ORDEM_PADROES, fill_value=0)
tabela12 = tabela11.div(tabela11.sum(axis=1), axis=0).mul(100).round(2)

print(f"\n=== TABELA 11 — {NOMES_TABELAS['11']} — contagem de alunos ===")
print(tabela11.to_string())
print("\n=== TABELA 12 — mesma contingência em % dentro de cada linha ===")
print(tabela12.to_string())

registra_tabela(
    NOME_ABA_TABELA11_12, tabela11, com_indice=True,
    titulo=f"Tabela 11 — {NOMES_TABELAS['11']} (contagem de alunos)",
)
registra_tabela(
    NOME_ABA_TABELA11_12, tabela12, com_indice=True,
    titulo="Tabela 12 — mesma contingência em percentual dentro de cada linha (Conceito Enade)",
    colunas_percentual=list(tabela12.columns),
)

# ==================================================================
# 13) TABELA 6 - acerto de todas as questões por Conceito Enade (anexo)
# ==================================================================
def monta_anexo_por_grupo(grupos: dict, prefixos: dict) -> pd.DataFrame:
    """Taxa de acerto de cada questão em cada grupo de alunos."""
    return pd.DataFrame([
        {"questao": questao, "bloco": classifica_bloco_questao(questao),
         **{k: v for chave, linhas in grupos.items()
            for k, v in linha_taxa_acerto(linhas, questao, prefixos[chave]).items()}}
        for questao in QUESTOES
    ])


tabela6 = monta_anexo_por_grupo(
    LINHAS_POR_CONCEITO, {c: f"Conceito{c}" for c in ORDEM_CONCEITOS}
)
print(f"\n=== TABELA 6 — {NOMES_TABELAS['6']} (primeiras linhas) ===")
print(tabela6.head(10).to_string(index=False))
registra_tabela(
    "Aba 6 - Questões x Enade curso", tabela6,
    titulo=f"Tabela 6 — {NOMES_TABELAS['6']}",
)

# ==================================================================
# 14) TABELA 13 - alternativa marcada nas questões selecionadas
# ==================================================================
# Denominador: candidatos com resposta validada (código 0/1) no item -
# o mesmo N das demais tabelas por questão.
NOME_ABA_TABELA13 = "Aba 11 - Alternativas marcadas"


def monta_tabela13_questao(questao: str) -> pd.DataFrame:
    coluna = coluna_canonica(questao)
    escolhas = pd.Series(ESC[VALIDO[:, coluna], coluna]).replace(CODIGOS_ESPECIAIS_ESCOLHA)
    contagem = escolhas.value_counts().reindex(ORDEM_ALTERNATIVAS, fill_value=0).astype(int)
    tabela = contagem.to_frame("Qtd. Candidatos")
    tabela.index.name = "Alternativa Marcada"
    tabela["% do Total"] = percentual(tabela["Qtd. Candidatos"], contagem.sum())
    return tabela


for questao in QUESTOES_SELECIONADAS:
    tabela13_questao = monta_tabela13_questao(questao)
    print(f"\n=== TABELA 13 — alternativa marcada — {questao} ===")
    print(tabela13_questao.to_string())
    registra_tabela(
        NOME_ABA_TABELA13, tabela13_questao, com_indice=True,
        titulo=f"Tabela 13 ({questao}) — alternativa marcada pelos candidatos",
    )

# ==================================================================
# 15) TABELA 7 - distribuição do % de acerto individual
# ==================================================================
# O % de acerto é calculado aluno a aluno (acertos próprios / itens
# válidos próprios) e só depois agrupado em faixas fixas.
NOME_ABA_TABELA7 = "Aba 7 - Acerto individual"


def resolve_conjunto(definicao):
    """Traduz uma entrada de CONJUNTOS_TABELA7 nas questões que de fato
    entram na análise. Um bloco (BLOCO_ESPECIFICO ou BLOCO_GERAL) vira
    todas as suas questões válidas; uma lista explícita é filtrada pelas
    questões válidas."""
    if isinstance(definicao, str):
        if definicao not in (BLOCO_ESPECIFICO, BLOCO_GERAL):
            raise ValueError(
                f"Bloco desconhecido em CONJUNTOS_TABELA7: {definicao!r}. "
                "Use BLOCO_ESPECIFICO, BLOCO_GERAL ou uma lista de questões."
            )
        return [q for q in QUESTOES if classifica_bloco_questao(q) == definicao]
    fora = [q for q in definicao if q not in QUESTOES]
    if fora:
        print(
            f"\nAviso: questões ignoradas por não estarem na análise por item: "
            f"{', '.join(fora)}"
        )
    return [q for q in definicao if q in QUESTOES]


def calcula_faixa_individual(questoes_conjunto, rotulo_conjunto):
    """Nível, % de acerto e faixa de cada aluno no conjunto de questões.
    Alunos sem nenhum item válido no conjunto ficam de fora."""
    colunas = [coluna_canonica(q) for q in questoes_conjunto]
    desempenho = pd.DataFrame(
        {
            "nivel_individual": nivel_por_linha,
            "n_validos": VALIDO[:, colunas].sum(axis=1),
            "n_acertos": (ACERTO[:, colunas] & VALIDO[:, colunas]).sum(axis=1),
        },
        index=base_itens.index,
    )
    n_sem_item_valido = int((desempenho["n_validos"] == 0).sum())
    if n_sem_item_valido:
        print(
            f"\nAviso: {n_sem_item_valido} alunos não têm nenhum item válido em "
            f"'{rotulo_conjunto}' e ficam fora desse bloco da Tabela 7."
        )
    desempenho = desempenho[desempenho["n_validos"] > 0].copy()
    desempenho["faixa_pct_acerto"] = pd.cut(
        desempenho["n_acertos"] / desempenho["n_validos"] * 100,
        bins=LIMITES_FAIXA_TABELA7,
        labels=ROTULOS_FAIXA_TABELA7,
        right=False,
    )
    return desempenho


for rotulo_conjunto, definicao in CONJUNTOS_TABELA7.items():
    questoes_conjunto = resolve_conjunto(definicao)
    desempenho_conjunto = calcula_faixa_individual(questoes_conjunto, rotulo_conjunto)

    for nivel in ORDEM_NIVEIS:
        subset = desempenho_conjunto[desempenho_conjunto["nivel_individual"] == nivel]
        contagem = (
            subset["faixa_pct_acerto"].value_counts().reindex(ROTULOS_FAIXA_TABELA7, fill_value=0)
        )
        tabela7_nivel = pd.DataFrame({
            "Faixa": ROTULOS_FAIXA_TABELA7,
            "Número de alunos": contagem.values,
            "% de alunos na faixa": percentual(contagem, len(subset)).values,
        })
        descricao_bloco = (
            f"{rotulo_conjunto} ({len(questoes_conjunto)} questões) — {nivel} (N={len(subset)})"
        )
        print(f"\n=== TABELA 7 — {descricao_bloco} ===")
        print(tabela7_nivel.to_string(index=False))
        registra_tabela(
            NOME_ABA_TABELA7, tabela7_nivel,
            titulo=f"Tabela 7 — {descricao_bloco} — {NOMES_TABELAS['7']}",
        )

# ==================================================================
# 16) TABELA 8 - acerto de todas as questões por Nível de Desempenho (anexo)
# ==================================================================
tabela8 = monta_anexo_por_grupo(
    LINHAS_POR_NIVEL, {n: n.replace(" ", "_") for n in ORDEM_NIVEIS}
)
print(f"\n=== TABELA 8 — {NOMES_TABELAS['8']} (primeiras linhas) ===")
print(tabela8.head(10).to_string(index=False))
registra_tabela(
    "Aba 8 - Questões x Individual", tabela8,
    titulo=f"Tabela 8 — {NOMES_TABELAS['8']}",
)

# ==================================================================
# 17) EXPORTAÇÃO PARA A PLANILHA
# ==================================================================
FONTE_PADRAO = "Calibri"
FORMATO_PERCENTUAL = '0.00"%"'   # exibe 26,36% sem multiplicar o valor por 100

FONTE_TITULO_BLOCO = Font(name=FONTE_PADRAO, bold=True, size=12, color="1F4E78")
FONTE_CABECALHO = Font(name=FONTE_PADRAO, bold=True, color="FFFFFF", size=11)
FONTE_CORPO = Font(name=FONTE_PADRAO, size=10)
PREENCHIMENTO_CABECALHO = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
PREENCHIMENTO_ZEBRA = PatternFill(start_color="EAF1F8", end_color="EAF1F8", fill_type="solid")
BORDA_FINA = Side(style="thin", color="BFBFBF")
BORDA_CELULA = Border(left=BORDA_FINA, right=BORDA_FINA, top=BORDA_FINA, bottom=BORDA_FINA)
ALINHAMENTO_CABECALHO = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALINHAMENTO_CORPO = Alignment(horizontal="center", vertical="center")
ALINHAMENTO_TITULO = Alignment(horizontal="left", vertical="center")

LINHAS_ENTRE_BLOCOS = 2


def escreve_bloco(planilha, linha_inicio, bloco):
    """Escreve título, cabeçalho e dados a partir de `linha_inicio`.
    Devolve (linha do cabeçalho, última linha de dados, nº de colunas,
    próxima linha livre)."""
    df = bloco["df"]
    com_indice = bloco["com_indice"]
    colunas_percentual = set(bloco["colunas_percentual"])

    nomes_colunas = ([df.index.name or ""] if com_indice else []) + [str(c) for c in df.columns]
    n_colunas = len(nomes_colunas)
    linha = linha_inicio

    if bloco["titulo"]:
        celula = planilha.cell(row=linha, column=1, value=bloco["titulo"])
        planilha.merge_cells(start_row=linha, start_column=1, end_row=linha, end_column=n_colunas)
        celula.font = FONTE_TITULO_BLOCO
        celula.alignment = ALINHAMENTO_TITULO
        planilha.row_dimensions[linha].height = 22
        linha += 1

    linha_cabecalho = linha
    for coluna, nome_coluna in enumerate(nomes_colunas, start=1):
        celula = planilha.cell(row=linha_cabecalho, column=coluna, value=nome_coluna)
        celula.font = FONTE_CABECALHO
        celula.fill = PREENCHIMENTO_CABECALHO
        celula.alignment = ALINHAMENTO_CABECALHO
        celula.border = BORDA_CELULA
    planilha.row_dimensions[linha_cabecalho].height = 32
    linha += 1

    linha_dados_inicio = linha
    dados = df.reset_index() if com_indice else df
    for _, valores in dados.iterrows():
        for coluna, valor in enumerate(valores, start=1):
            if isinstance(valor, float) and np.isnan(valor):
                valor = None
            elif isinstance(valor, (np.integer, np.floating, np.bool_)):
                valor = valor.item()
            celula = planilha.cell(row=linha, column=coluna, value=valor)
            celula.font = FONTE_CORPO
            celula.border = BORDA_CELULA
            celula.alignment = ALINHAMENTO_CORPO
            if nomes_colunas[coluna - 1] in colunas_percentual and isinstance(valor, (int, float)):
                celula.number_format = FORMATO_PERCENTUAL
            if (linha - linha_dados_inicio) % 2 == 1:
                celula.fill = PREENCHIMENTO_ZEBRA
        linha += 1

    return linha_cabecalho, linha - 1, n_colunas, linha + LINHAS_ENTRE_BLOCOS


# Move as abas de ABAS_NO_FINAL para o fim, conferindo se os nomes existem.
_faltando = [a for a in ABAS_NO_FINAL if a not in ORDEM_ABAS]
if _faltando:
    raise ValueError(f"ABAS_NO_FINAL contém abas que não existem: {_faltando}")
ORDEM_ABAS = [a for a in ORDEM_ABAS if a not in ABAS_NO_FINAL] + ABAS_NO_FINAL

livro = Workbook()
livro.remove(livro.active)

for nome_aba in ORDEM_ABAS:
    blocos = BLOCOS_POR_ABA[nome_aba]
    planilha = livro.create_sheet(title=nome_aba[:31])
    planilha.sheet_view.showGridLines = False
    linha_atual = 1
    primeiro_cabecalho = ultima_linha = None
    largura_maxima_colunas = 0
    for bloco in blocos:
        linha_cabecalho, linha_fim, n_colunas, linha_atual = escreve_bloco(
            planilha, linha_atual, bloco
        )
        if primeiro_cabecalho is None:
            primeiro_cabecalho = linha_cabecalho
        ultima_linha = linha_fim
        largura_maxima_colunas = max(largura_maxima_colunas, n_colunas)

    for coluna in range(1, largura_maxima_colunas + 1):
        largura = 0
        for linha in range(1, planilha.max_row + 1):
            valor = planilha.cell(row=linha, column=coluna).value
            if valor is not None:
                largura = max(largura, len(str(valor)))
        planilha.column_dimensions[get_column_letter(coluna)].width = min(max(largura + 3, 11), 45)

    # Congelar e filtrar só faz sentido quando a aba tem uma única
    # tabela; o intervalo começa no cabeçalho, não na linha do título.
    planilha.freeze_panes = f"A{primeiro_cabecalho + 1}"
    if len(blocos) == 1:
        ultima_coluna = get_column_letter(largura_maxima_colunas)
        planilha.auto_filter.ref = f"A{primeiro_cabecalho}:{ultima_coluna}{ultima_linha}"

livro.save(PATH_SAIDA_XLSX)
print(f"\nPlanilha gerada com {len(ORDEM_ABAS)} abas: {PATH_SAIDA_XLSX}")
print("Abas: " + ", ".join(ORDEM_ABAS))
