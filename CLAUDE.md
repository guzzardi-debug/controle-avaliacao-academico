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
* Não usar a palavra `humana` em nenhum lugar da planilha.
* Não automatizar julgamento acadêmico.
* Não inventar valores de delta para gatilhos.
* Não redesenhar livremente a lógica da Planilha de Notas oficial.

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
