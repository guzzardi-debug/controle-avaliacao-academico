# **ERRATA_GATILHOS_PONTUAVEIS.md**

**Natureza:** errata operacional. Documenta uma decisão explícita do projeto `Controle Avaliação` que diverge do uso original dos identificadores de gatilho definidos em `MATRIZ_DE_AVALIAÇÃO.md`. Esta decisão foi tomada pelo responsável do projeto em 2026-07-07, registrada em conversa de estruturação, e não é uma inferência automática de IA. Este documento não altera `MATRIZ_DE_AVALIAÇÃO.md`; ele documenta uma camada operacional adicional que vale apenas para o campo `Gatilho` do `Registro por Sessão` do Controle Avaliação.

---

## 1. O que muda

Na Matriz de Avaliação original, tokens como `C+2`, `D-1`, `F+2`, `P-1` são **IDs de catálogo**: cada um aponta para uma linha específica de uma tabela de gatilhos com descrição comportamental (ex.: `C+2` = "Redige ou melhora cláusula, documento, agenda, proposta de resolução ou emenda"). Nenhum desses IDs carrega, na Matriz original, um valor numérico de delta — o campo Delta de cada um está marcado `[A definir pela equipe acadêmica]` e a própria Matriz proíbe que esse valor seja inferido ou estimado (Seção 3).

No Controle Avaliação, a partir desta decisão, o campo `Gatilho` de cada eixo em `Registro por Sessão` passa a ser **pontuável**: o token digitado pela equipe é lido diretamente como um ajuste de pontos, não como uma busca em catálogo.

| Antes (Matriz de Avaliação) | Agora (Controle Avaliação, campo `Gatilho`) |
| :---- | :---- |
| `C+2` = ID do 2º gatilho de bônus de Contribuição (catálogo, sem número associado) | `C+2` = soma **2 pontos** em Contribuição |
| `C-1` = ID do 1º gatilho de penalidade de Contribuição (catálogo, sem número associado) | `C-1` = tira **1 ponto** em Contribuição |
| `D-4` = ID do 4º gatilho de penalidade de Diplomacia (catálogo, sem número associado) | `D-4` = tira **4 pontos** em Diplomacia |

Essa mudança vale só para o mecanismo interno do Controle Avaliação. Ela não redefine, corrige ou substitui o catálogo de gatilhos da Matriz de Avaliação, que continua sendo a referência qualitativa de **que tipo de comportamento** justifica o uso de um token em determinada direção (bônus ou penalidade) em determinado eixo.

## 2. Formato do token

Padrão: `<Eixo><Sinal><Número>`, onde:

- `Eixo` ∈ {`D` (Diplomacia e Respeito às Regras), `F` (Fidelidade à Função Designada), `C` (Contribuição), `P` (Participação)}
- `Sinal` ∈ {`+`, `-`}
- `Número` = inteiro

Múltiplos tokens no mesmo campo `Gatilho` de um eixo, separados por `;`. Exemplos válidos:

- `C+2`
- `C+2; C-1`
- `F+2; F-1`
- `D+1; D-1`
- `P+1; P-2`

Validação recomendada: inteiros de **-8 a +5**. Justificativa: Diplomacia e Fidelidade partem de base 8 e precisam poder cair até 0 (delta mínimo útil = -8); Contribuição e Participação partem de base 5 e precisam poder subir até 10 (delta máximo útil = +5).

## 3. Fórmula da Nota Preliminar

Por eixo, dentro de cada linha `Delegado × Sessão` de `Registro por Sessão`:

```
Nota Preliminar <Eixo> = limitar( base(<Eixo>) + soma(tokens de <Eixo> no campo Gatilho), 0, 10 )
```

Bases oficiais (Seção 1 da Matriz de Avaliação):

| Eixo | Base |
| :---- | ----: |
| Diplomacia e Respeito às Regras | 8 |
| Fidelidade à Função Designada | 8 |
| Contribuição | 5 |
| Participação | 5 |

A nota nunca fica abaixo de 0 nem acima de 10, independentemente da soma dos tokens.

### Exemplo

Registro em Contribuição:

- Evidência C: "Redigiu cláusula incorporada ao documento, mas centralizou parte da discussão."
- Gatilho C: `C+2; C-1`

Cálculo: base 5 + (+2 -1) = 6 → Nota Preliminar C = **6**.

## 4. O que a fórmula faz e o que ela não faz

- A fórmula **soma** os tokens que a equipe já escolheu digitar. Ela não decide, sugere ou infere quais tokens usar.
- A `Evidência` continua obrigatória para justificar qualquer token lançado — sem evidência descrevendo o fato observável, o lançamento do token não tem lastro acadêmico, mesmo que a fórmula calcule um número.
- O campo `Triagem` (dropdown de apoio, ver `Rubrica`) não participa da fórmula em nenhuma hipótese. Ele serve só para leitura acadêmica posterior (feedback, premiação, auditoria de padrões).
- Esta fórmula não resolve os parâmetros ainda pendentes da Matriz de Avaliação (Seção 13 de `MATRIZ_DE_AVALIAÇÃO.md`): peso relativo entre eixos, nota mínima automática por ausência, parâmetros de presidência/vice, penalidades adicionais do delegado dominador. Esses continuam `[A definir pela equipe acadêmica]` e fora do escopo desta errata.
- Direção Observacional continua sem gerar token de pontuação (delta = 0, Seção 3 da Matriz) — um registro observacional é feito só via `Evidência` e `Triagem`, sem token em `Gatilho`.

## 5. Travas preservadas

- Participação não compensa Contribuição: os campos de cada eixo são independentes; a soma de tokens de um eixo nunca transborda para outro.
- Evidência antes de nota: nenhuma Nota Preliminar deve ser lançada sem evidência textual correspondente no mesmo eixo/sessão.
- Revisão acadêmica final: a Nota Preliminar por sessão não é a nota final do delegado — a consolidação e a leitura crítica continuam na aba `Nota` e na revisão acadêmica de premiação (Seção 10 da Matriz de Avaliação).
- Terminologia: não usar a palavra "humana"; usar "revisão" ou "revisão acadêmica".

## 6. Escopo

Esta errata vale exclusivamente para o campo `Gatilho` do `Registro por Sessão` do Controle Avaliação. Ela não altera `MATRIZ_DE_AVALIAÇÃO.md`, não altera `ONTOLOGIA_ACADEMICA.md` e não altera o contrato técnico da Planilha de Notas oficial (`PLANITA_NOTAS_CNTRATO_TECNICO.md`).
