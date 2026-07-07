# Relatório de Entrega — Controle Avaliação

**Arquivo entregue:** `output/Controle Avaliação.xlsx`
**Gerado por:** `tools/build_controle_avaliacao.py` (regenerável)
**Validado por:** `tools/validate_controle_avaliacao.py`
**Fontes:** `CLAUDE.md`, `specs/MATRIZ_DE_AVALIAÇÃO.md`, `specs/ONTOLOGIA_ACADEMICA.md`, `specs/ERRATA_GATILHOS_PONTUAVEIS.md`, `specs/PLANITA_NOTAS_CONTRATO_TECNICO.md`, `input/Planilha de Notas (2).xlsx`.

A planilha tem exatamente quatro abas: `Preenchimento`, `Rubrica`, `Registro por Sessão` e `Nota`.

Esta rodada ajustou usabilidade e ordenamento sobre a base já entregue: `Preenchimento` ficou compacta (sem começar na linha 10), `Registro por Sessão` passou a ordenar por sessão em vez de por delegado, a coluna de ordenamento ganhou destaque roxo institucional, a listra alternada virou linha sim/linha não, o workbook foi configurado para recalcular automaticamente ao abrir, e os gatilhos pontuáveis foram testados no Microsoft Excel de verdade — não apenas por simulação em Python.

---

## 1. Como cada aba funciona

### Preenchimento

Lista-mãe do comitê, agora compacta: título na linha 1, instrução na linha 2, cabeçalho (`Delegado`, `Representação/Função`, `Presença geral`) na linha 3, e o primeiro delegado imediatamente na linha 4 — sem nenhuma linha vazia entre o cabeçalho e o primeiro registro. A coluna de numeração aparece automaticamente quando o delegado é preenchido. A ordem desta aba comanda a ordem do `Registro por Sessão` e da aba `Nota` — por isso a instrução fixa de não inserir nem excluir linhas. Capacidade: 100 delegados, igual à Planilha de Notas oficial.

### Rubrica

Referência operacional em seis seções: (1) eixos e notas base 8/8/5/5 com a pergunta avaliativa central de cada eixo; (2) a lógica dos gatilhos pontuáveis, deixando explícito que os tokens `C+2`, `D-1` etc. **não são os IDs fixos da Matriz original** e apontando para `specs/ERRATA_GATILHOS_PONTUAVEIS.md`, com tabela de exemplos calculados; (3) a escala de interpretação 0–10 da Matriz, com a nota de leitura contextual por base; (4) as travas de registro (evidência obrigatória para 9–10 e para nota abaixo da base, Participação não compensa Contribuição, registro observacional sem token, revisão acadêmica na consolidação); (5) as opções de Triagem por eixo — esta tabela é a fonte real dos menus suspensos do Registro; (6) o padrão de evidência, com exemplo válido, exemplos inválidos e a terminologia institucional.

### Registro por Sessão

Principal área de uso durante o evento. Grade fixa de `Delegado × Sessão`, sempre 5 sessões — mas agora **ordenada por sessão**: a 1ª sessão traz todos os delegados na ordem do `Preenchimento` (linhas 4–103, com 100 delegados), depois a 2ª sessão traz todos os delegados de novo (linhas 104–203), e assim até a 5ª sessão (linhas 404–503). A coluna de ordenamento (A) reinicia a numeração em `1` no início de cada bloco de sessão e tem destaque visual roxo institucional com texto branco — o mesmo padrão da coluna de índice da Planilha de Notas oficial. Estrutura de colunas:

- **Identificação** (automática, exceto os dois últimos): Sessão, Delegado, Representação/Função, Presença na sessão (dropdown), Status do registro (dropdown: Pendente, Em andamento, Concluído, Revisado).
- **Um bloco por eixo** — Diplomacia e Respeito às Regras, Fidelidade à Função Designada, Contribuição, Participação — cada um com exatamente quatro campos: `Evidência` (texto livre com quebra automática), `Gatilho` (tokens pontuáveis), `Triagem` (dropdown qualitativo alimentado pela Rubrica; não calcula nota) e `Nota Preliminar` (fórmula).
- **Síntese**: Síntese acumulada da sessão e Observação geral da sessão (texto livre).

A listra alternada agora é simples — uma linha cinza, uma linha branca, continuamente — em vez de alternar por bloco de delegado. A coluna de ordenamento é a única exceção: permanece roxa em todas as linhas úteis, nunca listrada.

### Nota

Dividida visualmente por uma coluna separadora escura em duas áreas:

- **Conferência e síntese** (colunas A–K, não colar): delegado, função, presença geral, médias por eixo, média geral, e duas colunas de texto acumulado — todas as Evidências das 5 sessões (rotuladas por sessão e eixo, ex.: `S1 C: …`) e todas as Sínteses/Observações — para revisão acadêmica, feedback e conferência final.
- **Área de colagem** (colunas M–AF): apenas as 20 notas por delegado, na sequência `D1–D5 | F1–F5 | C1–C5 | P1–P5`, na ordem do `Preenchimento`, sem nomes nem observações. O cabeçalho indica o destino: copiar `M4:AF103` e colar **somente valores** em `'Avaliações'!E4` da Planilha de Notas oficial.

Como o `Registro por Sessão` passou a ordenar por sessão, todas as 20 fórmulas de referência por delegado na `Nota` foram recalculadas para apontar para a linha correta em cada bloco de sessão (ex.: para o delegado 1, a sessão 2 do eixo Diplomacia agora está na linha 104 do Registro, não mais na linha 5) — sem isso, a colagem sairia com números trocados entre sessões.

## 2. Como a quantidade de delegados alimenta o Registro por Sessão

Cada delegado da linha *n* do `Preenchimento` (delegado *i* = linha `3 + i`) tem uma linha fixa pré-vinculada em cada um dos 5 blocos de sessão do Registro: linha `4 + (sessão − 1) × 100 + (i − 1)`. As fórmulas de identificação só exibem conteúdo quando o delegado existe: com 10 delegados preenchidos, aparecem 50 linhas úteis (10 em cada um dos 5 blocos de sessão); com 42, aparecem 210. As linhas de delegados não preenchidos ficam totalmente em branco (sem número, sessão ou nota fantasma) — o mesmo comportamento de template da Planilha de Notas oficial. O filtro da aba permite esconder as linhas vazias com um clique.

## 3. Como os gatilhos pontuáveis calculam a Nota Preliminar

Conforme `specs/ERRATA_GATILHOS_PONTUAVEIS.md`: o campo `Gatilho` aceita tokens `D+X`/`D-Y`, `F+X`/`F-Y`, `C+X`/`C-Y`, `P+X`/`P-Y` (inteiros, recomendado de −8 a +5), múltiplos no mesmo campo separados por ponto e vírgula — até 8 tokens por campo. A fórmula:

```
Nota Preliminar = limitar( base do eixo + soma dos tokens do eixo , 0 , 10 )
```

com bases 8 (Diplomacia, Fidelidade) e 5 (Contribuição, Participação). A implementação usa apenas funções clássicas (`SUBSTITUTE`, `MID`, `TRIM`, `VALUE`, `IFERROR`, `MAX`, `MIN`), sem funções de matriz — compatível com Google Sheets (ambiente predominante, conforme o contrato técnico) e com qualquer versão do Excel. A `Triagem` não participa da fórmula em nenhuma hipótese.

**Esta rodada testou a fórmula dentro do Microsoft Excel de verdade** (não só em simulação Python): uma cópia do arquivo foi aberta no Excel via automação, os tokens abaixo foram digitados nas células reais de `Gatilho`, o Excel recalculou e os valores foram lidos de volta da planilha calculada.

| Gatilho digitado | Eixo | Nota Preliminar esperada | Nota Preliminar calculada pelo Excel |
| :--- | :--- | ---: | ---: |
| `D+1` | Diplomacia | 9 | 9 |
| `D+1; D-2` | Diplomacia | 7 | 7 |
| `F+2; F-1` | Fidelidade | 9 | 9 |
| `C+2` | Contribuição | 7 | 7 |
| `C+2; C-1` | Contribuição | 6 | 6 |
| `P-2` | Participação | 3 | 3 |
| `C+2` no campo Gatilho D (eixo errado) | Diplomacia | 8 (ignorado, fica na base) | 8 |
| `xyz` (texto inválido) | Diplomacia | 8 (ignorado, fica na base) | 8 |
| Presença na sessão = `Falta`, Gatilho `D+1` | Diplomacia | `-` | `-` |

Os 9 casos bateram exatamente com o esperado. Esse teste está automatizado em `tools/validate_controle_avaliacao.py` e roda a cada validação (em macOS com Microsoft Excel instalado); em outro ambiente, o teste é pulado com um aviso explícito, sem derrubar as demais verificações.

## 4. Como a aba Nota consolida a área de colagem

Cada célula da área de colagem referencia diretamente a `Nota Preliminar` da linha exata do Registro (delegado × sessão × eixo, já usando o mapeamento sessão-primeiro) e só exibe o valor quando ele é numérico — faltas (`"-"`) e sessões sem registro viram células vazias, que a fórmula de média da planilha oficial já trata. A sequência de 20 colunas espelha exatamente `Avaliações!E:X` da Planilha de Notas oficial: 5 sessões de Diplomacia, 5 de Fidelidade, 5 de Contribuição, 5 de Participação, com os cabeçalhos de grupo na mesma grafia da planilha oficial. A linha 4 da Nota corresponde à linha 4 de `Avaliações` (delegado 1), preservando o alinhamento 1:1 da colagem. A aba Nota deliberadamente **não tem filtro**, para impedir cópia de intervalo filtrado desalinhado.

## 5. Elementos visuais reproduzidos da Planilha de Notas oficial

Extraídos por inspeção direta do arquivo em `input/`:

| Elemento | Valor reproduzido |
| :--- | :--- |
| Faixa de título | preto `#000000`, texto branco, 20 pt |
| Cabeçalhos de grupo e de coluna | roxo institucional `#20124D`, texto branco em negrito |
| Coluna de ordenamento (Registro por Sessão) | roxo institucional `#20124D`, texto branco, em todas as linhas úteis |
| Linhas alternadas | cinza `#CCCCCC` / branco, linha sim / linha não, contínuo |
| Destaque de falta | rosa `#E6B8AF` via formatação condicional (`Falta` pinta a linha), nas três abas de uso |
| Bordas | finas; pretas separando blocos de eixo, cinza `#B7B7B7` internas |
| Fonte | Calibri, mesmos tamanhos (20/16/12/11) |
| Congelamento | Preenchimento `C4`, Registro `E4`, Nota `C4` |
| Filtros | Preenchimento e Registro por Sessão |
| Larguras e quebra automática | colunas de evidência/síntese largas com quebra de texto; notas estreitas centralizadas |
| Cor das guias | roxo institucional nas quatro abas |
| Cabeçalhos de sessão | `1ª Sessão` … `5ª Sessão`, repetidos por eixo, como na aba `Avaliações` |
| Recálculo ao abrir | `fullCalcOnLoad` ativado — Excel/Sheets recalculam tudo na abertura, sem mostrar valores antigos |

## 6. Verificações feitas

Três camadas de validação automatizada, todas passando:

**Autoteste do parser de tokens** (executa a cada build, 11 casos, simulação Python rápida): `C+2`→+2; `C+2; C-1`→+1; minúsculas; `D+1; D-2`→−1; `P-2`→−2; campo vazio→0; texto inválido→0; separador por vírgula rejeitado (documentado: usar ponto e vírgula); 8 tokens somados; token de eixo errado ignorado.

**Validação estrutural do arquivo gerado** (`tools/validate_controle_avaliacao.py`, 53 verificações): exatamente as 4 abas na ordem do `CLAUDE.md`; nenhuma célula contém termo vetado (`humana`, `ID Delegado`, `Tipo de função`, `Mesa`, `Plenária` — os dois últimos só aparecem na Rubrica como orientação de "não usar"); `Preenchimento` sem linha vazia entre cabeçalho e primeiro delegado; cabeçalhos canônicos das três abas de dados; malha de 500 linhas do Registro ordenada por sessão (conferidas as fronteiras dos 5 blocos); numeração da coluna de ordenamento reiniciando em 1 a cada sessão; coluna de ordenamento roxa com texto branco em todas as linhas úteis; listra simples linha sim/linha não; as 4 fórmulas de Nota Preliminar com base certa, clamp 0–10, coluna de gatilho certa e 8 fatias de token; 8 referências cruzadas da área de colagem conferidas célula a célula já no mapeamento sessão-primeiro (incluindo os extremos `M4→J4` e `AF103→V503`); 6 validações de dados do Registro, com os 4 dropdowns de Triagem apontando para intervalos preenchidos da Rubrica; formatação condicional de falta nas três abas; congelamentos e filtros; área de colagem contendo somente as fórmulas de nota; `fullCalcOnLoad` ativado. Adicionalmente, as **7.200 fórmulas** do arquivo foram tokenizadas uma a uma para conferir sintaxe válida (0 problemas).

**Teste real dos gatilhos no Microsoft Excel** (mesma ferramenta de validação, 9 casos): descrito na Seção 3 — os 9 casos exigidos foram digitados em células reais de uma cópia do arquivo, calculados pelo próprio Excel via automação, e conferidos célula a célula. Todos bateram com o valor esperado.

## Princípio acadêmico preservado

Participação e Contribuição permanecem eixos independentes em toda a cadeia: campos separados no Registro, fórmulas separadas, colunas separadas na colagem, e a trava "Participação não compensa Contribuição" registrada na Rubrica. A Nota Preliminar organiza a pontuação lançada pela equipe; a decisão de quais tokens usar, a evidência que os justifica e a revisão acadêmica final continuam integralmente com a equipe.
