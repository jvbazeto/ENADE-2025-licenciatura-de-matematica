# Enade 2025 – Licenciatura em Matemática: desempenho por questão, curso e aluno

Script em Python que lê os microdados do Enade 2025 e gera uma planilha Excel com o desempenho dos estudantes de **Matemática (Licenciatura)** (`CO_GRUPO = 702`).

A partir dos microdados do Inep, o script calcula:

- a **taxa de acerto de cada questão** objetiva, por caderno de prova e no total;
- o **Conceito Enade de cada curso** (1 a 5, conforme a NT 8/2026);
- o **Nível de Desempenho individual** de cada aluno (conforme a NT 44/2025);
- cruzamentos entre esses resultados e o acerto em questões específicas.

## Requisitos

- Python 3
- Bibliotecas: `pandas`, `numpy` e `openpyxl`

Testado com Python 3.12, pandas 3.0, numpy 2.5 e openpyxl 3.1.

```
pip install pandas numpy openpyxl
```

## Arquivos de entrada

Coloque os três arquivos na mesma pasta do script.

| Arquivo | Origem | Conteúdo |
|---|---|---|
| `microdados2025_arq1.txt` | Microdados do Enade 2025 (site do Inep) | Características dos cursos (`CO_CURSO`, `CO_GRUPO`, `CO_MODALIDADE`, `CO_CATEGAD`) |
| `microdados2025_arq3.txt` | Microdados do Enade 2025 (site do Inep) | Registros dos estudantes: vetores de resposta, presença e notas |
| `mapeamento_questoes.csv` | Este repositório | Posição de cada questão (Q1 a Q80) em cada um dos cadernos 1 a 4 |

Os microdados **não estão neste repositório**: baixe-os no site do Inep (gov.br/inep), na seção de microdados do Enade. O `mapeamento_questoes.csv` usa `;` como separador e tem as colunas `questao;caderno1;caderno2;caderno3;caderno4`.

## Como executar

Na pasta que contém o script e os arquivos de entrada:

```
python enade_matematica_2025.py
```

No Windows, use sempre o comando acima. Digitar apenas o nome do arquivo (`enade_matematica_2025.py`) pode abrir o script no editor em vez de executá-lo. Se o comando `python` não for reconhecido, tente `py enade_matematica_2025.py`.

O script mostra o andamento e as tabelas no console e, ao final, grava `enade_matematica_2025_tabelas.xlsx` na mesma pasta. Feche a planilha no Excel antes de executar de novo, senão o arquivo não poderá ser sobrescrito.

## O que a planilha contém

A primeira aba, **Resumo**, lista cada tabela e o que ela mede. As demais abas são:

| Aba | Tabelas | Conteúdo |
|---|---|---|
| 1 – Acerto por questão | 1 | Taxa de acerto de cada questão, por caderno e total |
| 2 – Conceito Enade do curso | 2 | Conceito Enade (1 a 5) de cada curso |
| 3 – Cursos sem conceito | 3 | Cursos "Sem Conceito (SC)" |
| 4 – Cursos x Conceito Enade | 4 | Cursos e estudantes por Conceito Enade, no geral e em 8 recortes (modalidade e tipo de instituição) |
| 5 – Nível individual | 5 | Alunos por Nível de Desempenho, no geral e nos mesmos 8 recortes |
| 6 – Questões x Enade curso | 6 | Taxa de acerto de todas as questões, por Conceito Enade do curso |
| 7 – Acerto individual | 7 | Alunos por faixa de % de acerto, por Nível de Desempenho, para três conjuntos de questões |
| 8 – Questões x Individual | 8 | Taxa de acerto de todas as questões, por Nível de Desempenho |
| 9 – Questões em foco | 9 e 10 | Acerto das questões selecionadas, por Nível de Desempenho (9) e por Conceito Enade (10) |
| 10 – Conceito x acerto | 11 e 12 | Conceito Enade do curso cruzado com o padrão de acerto nas questões selecionadas (contagem e %) |
| 11 – Alternativas marcadas | 13 | Alternativa marcada em cada questão selecionada |

## Como ler as tabelas

Os exemplos de leitura abaixo vêm da execução com os microdados de Matemática 2025. Se você trocar os conjuntos de questões ou usar outros microdados, os números mudam.

### Conceitos comuns a todas as tabelas

**Quem entra na análise.** Só entram estudantes de Matemática (Licenciatura) presentes com resultado válido (`TP_PRES = 555`, NT 8/2026, item 6.2). Os demais códigos de presença ficam de fora. Em 2025 são 10.459 registros (o `arq1` lista 482 cursos de Matemática). Deles saem duas bases:

- **Base com nota:** os registros que têm a nota geral (`NT_GER`) preenchida (10.429 em 2025; 30 registros não têm). Serve para calcular o Nível de Desempenho de cada aluno e o Conceito Enade de cada curso.
- **Base por questão:** os mesmos registros, mas só dos cadernos 1 a 4 (o caderno 5 é uma reaplicação, com prova diferente) e com o vetor de acertos preenchido (10.409 em 2025; 50 ficam de fora por esses motivos). Serve para todas as tabelas que mostram acerto por questão. Ela não exige a nota geral, mas as tabelas que agrupam por nível ou por conceito só incluem alunos que têm essa classificação.

**Resposta válida, acerto e erro.** Para cada questão, o Inep registra um código por aluno: `1` (acertou), `0` (errou) ou `8` (questão desconsiderada). Uma *resposta válida* é a que tem código `0` ou `1`. Todo *erro* é uma resposta válida que não é acerto, o que inclui respostas em branco e com múltiplas marcações (a Tabela 13 as mostra separadamente). A taxa de acerto é sempre:

> taxa de acerto (%) = acertos ÷ (acertos + erros) × 100, arredondada a 2 casas decimais

**Questões desconsideradas pela TRI.** Questões em que todos os estudantes têm código `8` são identificadas automaticamente e excluídas das análises por questão. Com os dados de 2025 são a **Q3, a Q15 e a Q59**. Por isso a Tabela 1 tem 77 linhas (80 − 3), e os conjuntos "Formação Geral" e "Componente Específico" da Tabela 7 têm 28 e 49 questões. Para excluir outras questões manualmente, use `QUESTOES_EXCLUIDAS_ADICIONAIS`.

**Nível de Desempenho individual** (NT 44/2025, item 9.3), a partir da nota geral do aluno (`NT_GER`):

| `NT_GER` | Nível |
|---|---|
| menor que 50 | Abaixo do Básico |
| de 50 até menos de 70 | Básico |
| 70 ou mais | Adequado |

**Conceito Enade do curso** (NT 8/2026, item 7.4). Para cada curso, o script calcula:

- `NC_ck`: número de estudantes do curso na base com nota;
- `NCDB_ck`: quantos desses têm `NT_GER` de 50 ou mais (Básico ou acima);
- `PDB_ck = NCDB_ck ÷ NC_ck`, com 3 casas decimais.

O Conceito é definido pelo `PDB_ck`:

| `PDB_ck` | Conceito |
|---|---|
| menor que 0,400 | 1 |
| de 0,400 até menos de 0,600 | 2 |
| de 0,600 até menos de 0,750 | 3 |
| de 0,750 até menos de 0,900 | 4 |
| 0,900 ou mais | 5 |

Nas tabelas "por Conceito", cada aluno recebe o Conceito do **curso em que está matriculado**.

**Sem Conceito (SC).** O curso fica sem Conceito quando tem menos de 2 estudantes na base com nota (NT 8/2026, itens 5.1 e 6.6), o que inclui cursos sem nenhum registro válido. Em 2025: 441 cursos com Conceito e 41 SC. Os alunos de cursos SC **não aparecem** nas tabelas que agrupam por Conceito (Tabelas 4, 6, 10, 11 e 12), mas continuam nas demais.

---

### Aba 1 – Tabela 1: taxa de acerto por questão

**Uma linha por questão** (77 em 2025, sem as excluídas). Base: base por questão.

A mesma questão ocupa posições diferentes em cada caderno de prova, por isso a tabela mostra o resultado por caderno e também o total.

| Coluna | O que é |
|---|---|
| `questao` | Número da questão (Q1 a Q80) |
| `posicao_vetor_ACE_OBJ` | Posição da questão no vetor de acertos do Inep (igual ao número da questão) |
| `posicao_caderno1` ... `posicao_caderno4` | Posição que a questão ocupa em cada caderno, conforme `mapeamento_questoes.csv` |
| `N_valido_caderno1` ... `4` | Alunos desse caderno com resposta válida na questão |
| `acertos_caderno1` ... `4` | Desses, quantos acertaram |
| `taxa_acerto_caderno1_%` ... `4` | Acertos ÷ respostas válidas, em % |
| `N_valido_TOTAL`, `acertos_TOTAL` | Soma dos quatro cadernos |
| `taxa_acerto_TOTAL_%` | `acertos_TOTAL ÷ N_valido_TOTAL`, em % (não é a média das taxas dos cadernos) |

*Exemplo:* a Q1, no caderno 1, teve 2.636 respostas válidas e 788 acertos, ou seja, 29,89% de acerto.

### Aba 2 – Tabela 2: Conceito Enade de cada curso

**Uma linha por curso que tem Conceito** (441 em 2025).

| Coluna | O que é |
|---|---|
| `CO_CURSO` | Código do curso |
| `NC_ck` | Estudantes do curso na base com nota |
| `NCDB_ck` | Quantos têm `NT_GER` de 50 ou mais |
| `PDB_ck` | `NCDB_ck ÷ NC_ck` (3 casas) |
| `conceito_enade_curso` | Conceito Enade, de 1 a 5 |

*Exemplo:* o curso 26 tem 11 estudantes, 8 deles com nota de 50 ou mais. O `PDB_ck` é 0,727, que cai na faixa de 0,600 a menos de 0,750, então o Conceito é 3.

### Aba 3 – Tabela 3: cursos sem conceito

**Uma linha por curso "Sem Conceito (SC)"** (41 em 2025). Os cursos das Tabelas 2 e 3 somam o total de cursos de Matemática do `arq1` (441 + 41 = 482).

| Coluna | O que é |
|---|---|
| `CO_CURSO` | Código do curso |
| `concluintes_validos_TP_PRES_555` | Quantidade de registros do curso com `TP_PRES = 555`, mesmo que a nota geral esteja em branco. `0` significa que o curso não tem nenhum registro válido, e `1` que tem só um, abaixo do mínimo de 2 |

### Aba 4 – Tabela 4: cursos e estudantes por Conceito Enade

**Nove blocos:** o geral e 8 recortes. Em cada bloco há **sempre 5 linhas**, uma por Conceito (1 a 5), mesmo que algum Conceito não tenha cursos (então aparece com zeros). Base: apenas cursos **com** Conceito, de modo que os cursos SC não entram nos percentuais.

| Coluna | O que é |
|---|---|
| `numero_de_cursos` | Cursos com aquele Conceito |
| `total_de_estudantes` | Soma do `NC_ck` desses cursos (estudantes na base com nota) |
| `porcentagem_cursos_%` | Parte dos cursos do bloco que tem aquele Conceito |
| `porcentagem_cursos_acumulada_%` | Acumulado dessa porcentagem, do Conceito 1 até o 5 (fecha em 100) |
| `porcentagem_estudantes_%` | Parte dos estudantes do bloco que está em cursos com aquele Conceito |
| `porcentagem_estudantes_acumulada_%` | Acumulado dessa porcentagem |
| `numero_medio_estudantes_por_curso` | `total_de_estudantes ÷ numero_de_cursos` (vazio se não há cursos) |

**Os 8 recortes**, que também valem para a Tabela 5 e se baseiam nas características do curso:

| Recorte | Critério |
|---|---|
| EAD / Presencial | Modalidade do curso (`CO_MODALIDADE`: 0 = EAD, demais = Presencial) |
| Pública / Privada | Categoria administrativa (`CO_CATEGAD`: 1 a 3 = Pública, demais = Privada) |
| Pública sem EAD | Públicas apenas Presenciais |
| Privada sem EAD | Privadas apenas Presenciais |
| Pública sem Presencial | Públicas apenas EAD |
| Privada sem Presencial | Privadas apenas EAD |

*Exemplo (bloco geral):* 163 cursos (36,96%) têm Conceito 1, e eles reúnem 5.025 estudantes (48,18% do total), com média de 30,83 estudantes por curso.

### Aba 5 – Tabela 5: alunos por Nível de Desempenho

**Nove blocos** (geral e os mesmos 8 recortes da Tabela 4), cada um com **3 linhas**, uma por Nível. Base: base com nota, ou seja, cada aluno conta uma vez, **inclusive os de cursos SC** (diferente da Tabela 4). O recorte é definido pelo curso do aluno.

| Coluna | O que é |
|---|---|
| `numero_de_alunos` | Alunos naquele Nível |
| `pct_alunos_com_nivel_%` | Parte dos alunos do bloco que está naquele Nível |
| `pct_acumulada_%` | Acumulado, de Abaixo do Básico até Adequado (fecha em 100) |

*Exemplo (bloco geral):* 5.880 alunos (56,38%) estão no Abaixo do Básico, 3.939 (37,77%) no Básico e 610 (5,85%) no Adequado.

### Aba 6 – Tabela 6: acerto de todas as questões por Conceito Enade do curso

**Uma linha por questão** (77 em 2025). Base: base por questão, só com alunos de cursos que têm Conceito. Serve para comparar o desempenho na mesma questão entre cursos de Conceitos diferentes.

| Coluna | O que é |
|---|---|
| `questao` | Número da questão |
| `bloco` | `Formação Geral Docente (Q1-Q30)` ou `Componente Específico (Q31-Q80)` |
| `Conceito1_acertos` ... `Conceito5_acertos` | Acertos dos alunos de cursos com aquele Conceito |
| `Conceito1_erros` ... `Conceito5_erros` | Erros desses alunos |
| `Conceito1_taxa_acerto_%` ... `Conceito5_taxa_acerto_%` | Acertos ÷ (acertos + erros), em % |

*Exemplo:* na Q1, os alunos de cursos de Conceito 1 tiveram 1.526 acertos e 3.489 erros, ou seja, 30,43% de acerto.

### Aba 7 – Tabela 7: distribuição do % de acerto individual

Mostra como os alunos se distribuem em faixas de **percentual de acerto individual**, separando por Nível de Desempenho e por conjunto de questões.

**Como é calculada.** Para cada aluno, o script divide os acertos pelas respostas válidas **dele** nas questões do conjunto (acertou 12 de 20 respostas válidas = 60%). Esse percentual é então enquadrado em uma das faixas. O aluno que não tem nenhuma resposta válida no conjunto fica de fora.

**Blocos.** Cada conjunto de questões gera 3 blocos (um por Nível), então são 9 blocos com os 3 conjuntos padrão: Componente Específico, Formação Geral e o conjunto definido pelo usuário (ver "Como trocar o conjunto de questões"). O título de cada bloco traz o nome do conjunto, o número de questões que de fato entraram (ex.: `49 questões`), o Nível e o `N`, o total de alunos daquele Nível no bloco. O `N` pode ser menor que o da Tabela 5, pois aqui só entram alunos da base por questão.

**Cada bloco tem 5 linhas** (faixas fechadas à esquerda e abertas à direita; a última inclui 100%):

| Faixa |
|---|
| `[0%, 20%)` |
| `[20%, 40%)` |
| `[40%, 60%)` |
| `[60%, 80%)` |
| `[80%, 100%]` |

| Coluna | O que é |
|---|---|
| `Faixa` | Faixa de % de acerto individual |
| `Número de alunos` | Alunos do Nível cujo % de acerto cai na faixa |
| `% de alunos na faixa` | Parte do `N` do bloco que está na faixa (cada bloco soma 100%) |

*Exemplo (Componente Específico, Abaixo do Básico, N = 5.870):* 4.627 alunos (78,82%) acertaram de 20% a menos de 40% das questões.

### Aba 8 – Tabela 8: acerto de todas as questões por Nível de Desempenho

**Uma linha por questão** (77 em 2025). É a mesma ideia da Tabela 6, mas agrupando pelo Nível de Desempenho do aluno em vez do Conceito do curso. Base: base por questão, com os alunos que têm Nível (inclui alunos de cursos SC).

| Coluna | O que é |
|---|---|
| `questao`, `bloco` | Igual à Tabela 6 |
| `Abaixo_do_Básico_acertos`, `Básico_acertos`, `Adequado_acertos` | Acertos dos alunos de cada Nível |
| `Abaixo_do_Básico_erros`, `Básico_erros`, `Adequado_erros` | Erros desses alunos |
| `Abaixo_do_Básico_taxa_acerto_%`, `Básico_taxa_acerto_%`, `Adequado_taxa_acerto_%` | Taxa de acerto, em % |

*Exemplo:* na Q1, a taxa de acerto é 26,03% no Abaixo do Básico, 41,10% no Básico e 47,87% no Adequado.

### Aba 9 – Tabelas 9 e 10: acerto das questões selecionadas

Mostram o acerto **somente nas questões da lista `QUESTOES_SELECIONADAS`** (padrão: Q48, Q54 e Q67). Cada questão é calculada separadamente, e o aluno entra na conta de uma questão se tiver resposta válida nela.

- **Tabela 9:** uma linha por **Nível de Desempenho** (3 linhas).
- **Tabela 10:** uma linha por **Conceito Enade do curso** (5 linhas, sem cursos SC).

Para cada questão da lista há três colunas:

| Coluna | O que é |
|---|---|
| `Q48_acertos` | Acertos na Q48 entre os alunos da linha |
| `Q48_erros` | Erros na Q48 entre os alunos da linha |
| `Q48_taxa_acerto_%` | Acertos ÷ (acertos + erros), em % |

*(O mesmo vale para `Q54_...` e `Q67_...`.)*

*Exemplos:* na Tabela 9, os alunos do Abaixo do Básico tiveram 1.376 acertos e 4.494 erros na Q48, ou seja, 23,44%. Na Tabela 10, os alunos de cursos de Conceito 1 tiveram 24,73% de acerto na Q48.

### Aba 10 – Tabelas 11 e 12: Conceito do curso x padrão de acerto

Respondem à pergunta: *entre os alunos de cursos de cada Conceito, quantos acertaram exatamente quais das questões selecionadas?*

**Padrão de acerto.** É o conjunto das questões selecionadas que o aluno acertou. Com as três questões padrão, os padrões possíveis são 8:

| Padrão | Significado |
|---|---|
| `∅` | Não acertou nenhuma das três |
| `{Q48}`, `{Q54}`, `{Q67}` | Acertou apenas aquela questão |
| `{Q48, Q54}`, `{Q48, Q67}`, `{Q54, Q67}` | Acertou exatamente essas duas e errou a terceira |
| `{Q48, Q54, Q67}` | Acertou as três |

Cada aluno cai em **um único** padrão, então as colunas de uma linha somam o total de alunos daquele Conceito. A ordem das colunas é: `∅`, as questões individuais, os pares e, por fim, todas juntas.

**Quem entra.** Apenas alunos com resposta válida em **todas** as questões selecionadas ao mesmo tempo e que estejam em cursos com Conceito.

- **Tabela 11:** linhas = Conceito Enade do curso (1 a 5); colunas = padrões de acerto; células = **número de alunos**.
- **Tabela 12:** a mesma contingência, mas em **percentual dentro de cada linha** (cada linha soma 100%, a menos de arredondamento).

*Exemplo:* os cursos de Conceito 1 têm 5.015 alunos na Tabela 11. Destes, 1.762 (35,13%, na Tabela 12) não acertaram nenhuma das três questões, e 290 (5,78%) acertaram as três.

### Aba 11 – Tabela 13: alternativa marcada em cada questão selecionada

**Uma tabela por questão** da lista `QUESTOES_SELECIONADAS` (3 tabelas por padrão). Mostra quantos candidatos marcaram cada alternativa.

| Coluna | O que é |
|---|---|
| `Alternativa Marcada` | `A`, `B`, `C`, `D`, `Em branco (.)` ou `Múltiplas escolhas (*)` (sempre as 6 linhas, mesmo que alguma tenha zero) |
| `Qtd. Candidatos` | Candidatos que marcaram aquela opção |
| `% do Total` | Parte dos candidatos da tabela que marcou a opção |

**Base.** Todos os candidatos da base por questão com resposta válida naquela questão, de qualquer curso, com ou sem Conceito e Nível, e de qualquer caderno (a resposta de cada aluno é reposicionada para a questão correta, independentemente do caderno em que ela apareceu). Por isso o total de cada tabela é o mesmo `N_valido_TOTAL` da Tabela 1 para aquela questão.

*Exemplo (Q48):* 3.900 candidatos (37,47%) marcaram a alternativa D, e 21 (0,20%) deixaram em branco.

## Como trocar o conjunto de questões

As duas listas abaixo ficam no bloco **CONFIGURAÇÃO**, no início do script, e são independentes entre si. Depois de editar, basta executar o script de novo: os títulos das tabelas e o Resumo se atualizam sozinhos.

### Aba 7: `CONJUNTOS_TABELA7`

A Aba 7 mostra, para cada Nível de Desempenho (Abaixo do Básico, Básico e Adequado), quantos alunos ficaram em cada faixa de % de acerto. Cada entrada do dicionário `CONJUNTOS_TABELA7` gera um bloco de **três tabelas** (uma por nível). Por padrão há três conjuntos:

```python
CONJUNTOS_TABELA7 = {
    "todas as questões do Componente Específico (Q31 a Q80)": BLOCO_ESPECIFICO,
    "todas as questões do Componente de Formação Geral (Q1 a Q30)": BLOCO_GERAL,
    "questões definidas pelo usuário": [
        "Q32", "Q34", "Q38", "Q42", "Q47", "Q48", "Q50",
        "Q51", "Q54", "Q62", "Q65", "Q66", "Q67", "Q68", "Q69", "Q70",
        "Q71", "Q72", "Q75", "Q76", "Q77", "Q78", "Q79", "Q80",
    ],
}
```

- **A chave** (o texto entre aspas) é apenas o nome do conjunto e aparece nos títulos da aba. Pode ser alterada livremente.
- **O valor** é `BLOCO_ESPECIFICO` ou `BLOCO_GERAL`, que usam todas as questões do bloco, ou uma **lista de questões**.
- **Para trocar o conjunto personalizado**, edite a lista: cada questão é escrita entre aspas como `"Q"` mais o número (de `"Q1"` a `"Q80"`), separadas por vírgula. Pode misturar questões de Formação Geral e do Componente Específico.
- **Para criar outro conjunto**, acrescente uma nova entrada ao dicionário. Cada conjunto novo adiciona três tabelas à aba. Para remover um conjunto, apague a entrada.
- Questões inexistentes ou excluídas da análise (por exemplo, as desconsideradas pela TRI) são **ignoradas**, e o console mostra `Aviso: questões ignoradas por não estarem na análise por item`. O título de cada bloco informa quantas questões realmente entraram, como em `(24 questões)`.

As faixas de acerto são definidas por `LARGURA_FAIXA_TABELA7`. O valor padrão, `20`, gera `[0%, 20%)`, `[20%, 40%)`, `[40%, 60%)`, `[60%, 80%)` e `[80%, 100%]`. O valor precisa dividir 100 exatamente (5, 10, 20, 25 ou 50); com outros valores, como 30, os rótulos saem errados.

### Abas 9, 10 e 11: `QUESTOES_SELECIONADAS`

As três abas usam a mesma lista:

```python
QUESTOES_SELECIONADAS = ["Q48", "Q54", "Q67"]
```

| Aba | O que a lista define |
|---|---|
| 9 | Colunas de acertos, erros e taxa de acerto de cada questão da lista (Tabelas 9 e 10) |
| 10 | Padrões de acerto do aluno: quais das questões da lista ele acertou, como `{Q48}`, `{Q48, Q54}` ou `∅` (nenhuma). Cruzados com o Conceito Enade do curso (Tabelas 11 e 12) |
| 11 | Uma tabela por questão da lista, com a quantidade e o % de candidatos em cada alternativa (A a D, em branco e múltiplas escolhas) |

Para trocar as questões, edite a lista, por exemplo `QUESTOES_SELECIONADAS = ["Q35", "Q40", "Q72"]`. Pode ter qualquer quantidade de questões, de Formação Geral ou do Componente Específico.

## Aviso

Os resultados são cálculos próprios feitos a partir dos microdados públicos, seguindo as notas técnicas citadas acima. Não se trata de uma publicação oficial do Inep.
