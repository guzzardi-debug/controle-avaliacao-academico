# Projeto — Controle Avaliação Acadêmico IN

Este repositório serve para construir a planilha `Controle Avaliação` da Internationali Negotia, eixo Acadêmico.

## Objetivo do projeto

Criar uma planilha auxiliar por comitê para registrar evidências acadêmicas por sessão, organizar notas por eixo e gerar uma área final de colagem compatível com a aba `Avaliações` da Planilha de Notas oficial.

A planilha não substitui a Planilha de Notas oficial. Ela serve como instrumento de acompanhamento, registro qualitativo, conferência acadêmica e preparação das notas que serão coladas na planilha oficial.

## Fontes e hierarquia

Use como fontes de conteúdo:

1. `specs/` — especificações acadêmicas e operacionais deste projeto.
2. `input/` — arquivos originais anexados, especialmente a Planilha de Notas oficial.
3. `output/` — entregas geradas.

A Planilha de Notas oficial é referência visual, operacional e de compatibilidade de colagem. Ela não define a rubrica acadêmica.

A Matriz de Avaliação e a Ontologia Acadêmica definem os eixos, bases, gatilhos, travas e distinções conceituais. Não invente deltas, pesos, critérios ou códigos rápidos quando a fonte indicar que estão pendentes.

## Escopo obrigatório da planilha

A planilha `Controle Avaliação` deve ter apenas quatro abas:

1. `Preenchimento`
2. `Rubrica`
3. `Registro por Sessão`
4. `Nota`

Não crie abas extras.

## Restrições

* A quantidade de linhas deve depender da quantidade real de delegados preenchidos na aba `Preenchimento`.
* Não usar campo `ID Delegado`.
* Não usar campo `Tipo de função`.
* Não usar a palavra `humana` em nenhum lugar da planilha. Usar `revisão` ou `revisão acadêmica`.
* Não automatizar julgamento acadêmico. A automação soma pontos de tokens já escolhidos pela equipe; ela não decide quais tokens usar nem dispensa a Evidência.
* Não inventar valores de delta para gatilhos da Matriz de Avaliação (catálogo D+1…P-7, com deltas por intensidade ainda `[A definir pela equipe acadêmica]`). Exceção registrada: no campo `Gatilho` do `Registro por Sessão`, tokens como `C+2`/`D-1` são pontuáveis por decisão explícita do projeto — ver "Lógica de pontuação por Gatilho" abaixo e `specs/ERRATA_GATILHOS_PONTUAVEIS.md`. Essa exceção não autoriza inventar novos deltas fora desse mecanismo.
* Não redesenhar livremente a lógica da Planilha de Notas oficial.

## Estrutura do Registro por Sessão

Grade fixa: uma linha por `Delegado × Sessão`, sempre considerando 5 sessões.

Bloco de identificação: `Sessão`, `Delegado`, `Representação/Função`, `Presença na sessão`, `Status do registro`.

Cada eixo (Diplomacia, Fidelidade, Contribuição, Participação) tem exatamente quatro campos próprios — nunca um campo único de evidência/nota para a linha inteira:

1. `Evidência` — texto livre com quebra automática; observável e útil para revisão acadêmica, não precisa ser curtíssimo.
2. `Gatilho` — pontuável (ver seção seguinte); aceita múltiplos tokens separados por `;`.
3. `Triagem` — dropdown de apoio à leitura acadêmica posterior; não substitui o gatilho e não calcula nota.
4. `Nota Preliminar` — calculada automaticamente a partir do campo `Gatilho` daquele eixo.

Bloco de síntese: `Síntese acumulada da sessão`, `Observação geral da sessão`.

Não criar: `Gatilho 1/2/3`, `Delta 1/2/3`, "Item de ajuste", `Nota Final` (fica só em `Nota`), "Auxílio acadêmico" (o nome correto é `Triagem`), `Código rápido` (bloqueado pela Matriz).

## Lógica de pontuação por Gatilho (decisão do projeto — 2026-07-07)

Decisão explícita do projeto, não inferência automática: no `Registro por Sessão` do Controle Avaliação, o campo `Gatilho` é pontuável. Isso diverge deliberadamente do uso original da Matriz de Avaliação, onde `C+2`, `D-1` etc. são apenas IDs de catálogo (ver `specs/ERRATA_GATILHOS_PONTUAVEIS.md` para o detalhamento completo). A Matriz de Avaliação em si não foi alterada; a divergência vale só para este mecanismo operacional.

Formato do token: `D+X`, `D-Y`, `F+X`, `F-Y`, `C+X`, `C-Y`, `P+X`, `P-Y`, onde `X`/`Y` são inteiros. Múltiplos tokens no mesmo campo, separados por `;` (ex.: `C+2; C-1`). Validação recomendada: inteiros de -8 a +5.

Fórmula por eixo: `Nota Preliminar = limitar(base do eixo + soma dos tokens do eixo, 0, 10)`. Bases: Diplomacia = 8, Fidelidade = 8, Contribuição = 5, Participação = 5.

A Evidência continua obrigatória para justificar qualquer token usado — a fórmula organiza a soma, não substitui a revisão acadêmica.

## Estrutura da aba Nota

Consolida as `Notas Preliminares` do `Registro por Sessão` para montar, por delegado, na mesma ordem de `Preenchimento`:

`D1 | D2 | D3 | D4 | D5 | F1 | F2 | F3 | F4 | F5 | C1 | C2 | C3 | C4 | C5 | P1 | P2 | P3 | P4 | P5`

Essa é a área compatível com colagem em `Avaliações!E4:X103` da Planilha de Notas oficial. Pode incluir também uma área de síntese por delegado, acumulando os textos de evidência e observação das cinco sessões, para apoiar revisão acadêmica, feedback e conferência final.

## Princípio acadêmico central

Separar Participação de Contribuição.

Participação mede presença ativa, regularidade e volume de engajamento.

Contribuição mede impacto, avanço material, proposta, cláusula, documento, acordo, destravamento, organização de solução e materialização.

Participação alta não compensa Contribuição baixa.

## Estilo de trabalho esperado

Antes de executar mudanças grandes, leia os arquivos em `specs/` e `input/`.

Quando a tarefa for de estruturação, entregue primeiro um plano de execução e uma lista de arquivos que serão criados ou alterados.

Quando a tarefa for de implementação, produza a planilha ou os arquivos necessários e depois verifique se a entrega respeita as specs.

Não faça abstrações fora do escopo solicitado.
