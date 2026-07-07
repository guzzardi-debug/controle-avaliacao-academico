---

## **titulo: PLANILHA\_NOTAS\_CONTRATO\_TECNICO tipo: contrato\_tecnico\_planilha eixo: Academico fonte\_primaria: Planilha de Notas.xlsx fonte\_secundaria: PLANILHA\_NOTAS\_AUDITORIA.md status: consolidado\_com\_lacunas\_fechadas uso\_recomendado: automacoes\_codigo\_integracoes\_futuras ambiente\_predominante: Google Sheets observacao: Este contrato descreve como ler e escrever na Planilha de Notas sem sobrescrever fórmulas nem alterar sua lógica operacional.**

# **PLANILHA\_NOTAS\_CONTRATO\_TECNICO**

## **1\. Escopo**

Este documento descreve a estrutura técnica da `Planilha de Notas.xlsx` para automações, integrações, scripts, Claude Code ou manutenção técnica.

Ele não altera o processo avaliativo. Ele não propõe nova planilha. Ele não define critérios acadêmicos. Sua função é impedir que uma automação escreva em áreas calculadas, quebre fórmulas, ignore validações ou interprete errado os blocos da planilha.

A planilha é usada predominantemente em Google Sheets por staff operacional que normalmente não domina fórmulas. Portanto, qualquer automação, manutenção ou ajuste deve assumir que o usuário final só deve preencher campos de entrada, principalmente notas.

## **2\. Princípio técnico central**

A planilha deve ser tratada como template com áreas de entrada restritas.

Regra operacional:

Automação pode escrever apenas nos ranges de entrada.

Automação deve ler, mas não sobrescrever, ranges calculados.

Automação deve preservar fórmulas, mesclas, filtros, colunas ocultas, validações e formatações condicionais.

Não foi detectada proteção de aba nem proteção estrutural do arquivo. A ausência de proteção decorre de inércia histórica, não de necessidade operacional confirmada. Portanto, a automação precisa aplicar proteção lógica própria: não escrever fora dos ranges permitidos.

Em eventual versão protegida, as fórmulas devem ficar protegidas contra edição, mas visíveis para manutenção

## **3\. Abas e função técnica**

| Ordem | Aba | Estado | Range útil | Função técnica |
| ----: | :---- | :---- | :---- | :---- |
| 1 | `Preenchimento` | visível | `A1:H109` | parâmetros globais e cadastro |
| 2 | `Treinamentos` | visível | `A1:H102` | entrada de presença e cálculo de média de treinamento |
| 3 | `Avaliações` | visível | `A1:Y103` | entrada de presença e notas por eixo/sessão |
| 4 | `DPO` | visível | `A1:I102` | entrada de notas do DPO e cálculo de média ponderada |
| 5 | `NC` | visível | `A1:V103` | entrada de Noite Cultural por grupo e cálculo individual |
| 6 | `Resultados` | visível | `A1:I102` | consolidação dos blocos e nota final |
| 7 | `Classificados` | visível | `A1:G102` | saída de ranking/classificados |
| 8 | `Distribuição de pontuação` | oculta | `H9:I31` | tabela visual de pesos e gráfico |

A aba `Distribuição de pontuação` deve permanecer oculta. Ela não deve ser tratada como motor de cálculo, pois os pesos usados nas fórmulas principais estão hardcoded.

## **4\. Ranges de escrita autorizada**

Automação só deve escrever nos seguintes ranges:

| Aba | Range | Tipo de dado |
| :---- | :---- | :---- |
| `Preenchimento` | `F5` | lista: `Online`, `1`, `2` |
| `Preenchimento` | `F7` | lista: `Stand`, `Stand e Apresentação de Palco`, `Sem noite Cultural` |
| `Preenchimento` | `B10:H109` | cadastro |
| `Treinamentos` | `D3:E102` | `P` ou `F` |
| `Avaliações` | `C4:C103` | `Presente` ou `Falta` |
| `Avaliações` | `E4:X103` | números 0–10 esperados |
| `DPO` | `E3:H102` | números 0–10 esperados |
| `NC` | `P4:V29` | números 0–10 esperados por grupo |

Qualquer escrita fora desses ranges deve ser tratada como risco de quebra de template.

## **5\. Campos automáticos que não devem ser sobrescritos**

| Aba | Range/campo automático | Motivo |
| :---- | :---- | :---- |
| `Treinamentos` | `B3:C102` | puxa nome/função do cadastro |
| `Treinamentos` | `F3:G102` | converte `P/F` em nota; colunas ocultas |
| `Treinamentos` | `H3:H102` | média de treinamento |
| `Avaliações` | `B4:B103` | puxa nome do cadastro |
| `Avaliações` | `D4:D103` | puxa função do cadastro |
| `Avaliações` | `Y4:Y103` | média interna da aba |
| `DPO` | `B3:D102` | puxa nome, presença e função |
| `DPO` | `I3:I102` | média ponderada do DPO |
| `NC` | `B4:M103` | puxa dados, busca notas por grupo e calcula média |
| `NC` | `O4:O29` | lista de grupos A–Z |
| `Resultados` | `B3:I102` | consolidação integral |
| `Classificados` | `B3:G102` | saída calculada de classificados |
| `Distribuição de pontuação` | `H9:I31` | tabela de pesos/gráfico; aba oculta |

## **6\. Cabeçalhos canônicos**

### **6.1. `Preenchimento`**

Cabeçalhos de cadastro em `B9:H9`:

Nome | Série/turma | Função | Grupo Noite Cultural | E-mail | Telefone | Unidade Escolar

Campos globais:

| Célula | Cabeçalho |
| :---- | :---- |
| `F5` | Número de Treinamentos |
| `F7` | Noite Cultural |

### **6.2. `Treinamentos`**

Cabeçalhos principais:

| Célula | Cabeçalho |
| :---- | :---- |
| `B2` | Nome |
| `C2` | Função |
| `D1` | Treinamento 1 ou Online |
| `E1` | Treinamento 2 ou `-` |
| `D2:E2` | Presença |
| `H2` | Média |

### **6.3. `Avaliações`**

Cabeçalhos por grupo:

| Range | Cabeçalho |
| :---- | :---- |
| `C2` | Presença |
| `E2:I2` | Diplomacia e Respeito às regras |
| `J2:N2` | Fidelidade à Função designada |
| `O2:S2` | Contribuição |
| `T2:X2` | Participação |
| `Y3` | Média |

Cabeçalhos de sessão em `E3:X3`:

1ª Sessão | 2ª Sessão | 3ª Sessão | 4ª Sessão | 5ª Sessão

A sequência de sessões se repete para cada eixo.

### **6.4. `DPO`**

Cabeçalhos em `B2:I2`:

Nome | \[presença oculta em C\] | Função | Ortografia | Conteúdo | Apresentação | Respeito ao prazo | Média

Observação: `C` está oculta e contém presença puxada de `Avaliações`.

### **6.5. `NC`**

Cabeçalhos em `B3:M3`:

Nome | \[presença oculta em C\] | Função | Grupo | Tempo | Originalidade | Caracterização | Tempo | Originalidade | Conteúdo | Caracterização | Média

Os cabeçalhos `Tempo`, `Originalidade` e `Caracterização` se repetem. A automação deve desambiguar pelo cabeçalho pai:

| Cabeçalho pai | Range |
| :---- | :---- |
| Apresentação de Palco | `F2:H2` e entrada `P2:R2` |
| Estande | `I2:L2` e entrada `S2:V2` |

Tabela de entrada por grupo em `O3:V29`:

Grupo | Tempo | Originalidade | Caracterização | Tempo | Originalidade | Conteúdo | Caracterização

### **6.6. `Resultados`**

Cabeçalhos em `B2:I2`:

Nome | Função | Treinamentos | Avaliação 1 | Avaliação 2 | DPO | Noite Cultural | Nota final

### **6.7. `Classificados`**

Cabeçalhos em `B2:G2`:

Nome | Turma | E-mail | Nota | Telefone | Unidade Escolar

## **7\. Aliases**

Não foram detectados aliases nativos no XLSX.

Aliases abaixo são apenas normalização operacional inferida para automações e não substituem os cabeçalhos reais da planilha.

| Cabeçalho real | Alias inferido |
| :---- | :---- |
| `Série/turma` | `serie_turma` |
| `Função` | `funcao` |
| `Grupo Noite Cultural` | `grupo_noite_cultural` |
| `E-mail` | `email` |
| `Unidade Escolar` | `unidade_escolar` |
| `Diplomacia e Respeito às regras` | `diplomacia` |
| `Fidelidade à Função designada` | `fidelidade` |
| `Contribuição` | `contribuicao` |
| `Participação` | `participacao` |
| `Respeito ao prazo` | `prazo` |
| `Noite Cultural` | `nc` |
| `Nota final` | `nota_final` |

Status: inferido.

## **8\. Estratégia recomendada para localizar cabeçalhos**

Para automações, usar prioridade de localização nesta ordem:

1. localizar a aba pelo nome exato;  
2. validar o range útil esperado;  
3. localizar cabeçalhos por endereço fixo;  
4. usar texto do cabeçalho apenas como confirmação;  
5. evitar localizar colunas apenas por nome quando houver cabeçalhos duplicados.

Casos que exigem cuidado:

| Aba | Risco |
| :---- | :---- |
| `Avaliações` | cabeçalhos de sessão se repetem em quatro blocos |
| `NC` | `Tempo`, `Originalidade` e `Caracterização` aparecem em Palco e Estande |
| `DPO` | coluna `C` está oculta, mas faz parte da estrutura |
| `Treinamentos` | colunas `F:G` estão ocultas e alimentam `H` |
| `Classificados` | ranking automático ou aparentemente dinâmico precisa de revisão técnica |

## **9\. Campos de leitura recomendados**

| Objetivo | Range de leitura |
| :---- | :---- |
| Cadastro bruto | `Preenchimento!B10:H109` |
| Parâmetros globais | `Preenchimento!F5`, `Preenchimento!F7` |
| Presença de treinamento | `Treinamentos!D3:E102` |
| Média de treinamento | `Treinamentos!H3:H102` |
| Presença geral | `Avaliações!C4:C103` |
| Notas por eixo/sessão | `Avaliações!E4:X103` |
| Média interna de avaliação | `Avaliações!Y4:Y103` |
| Notas DPO | `DPO!E3:H102` |
| Média DPO | `DPO!I3:I102` |
| Entrada Noite Cultural por grupo | `NC!O4:V29` |
| Média Noite Cultural individual | `NC!M4:M103` |
| Resultados consolidados | `Resultados!B3:I102` |
| Classificados | `Classificados!B3:G102` |
| Pesos visuais | `Distribuição de pontuação!H9:I31` |

## **10\. Fórmulas críticas**

As fórmulas abaixo são exemplos por primeira linha útil. Fórmulas equivalentes se repetem nas linhas seguintes.

### **10.1. Treinamentos**

Conversão de presença em nota:

\=IF(D3="P",10,IF(D3="F",0,"-"))

Média de treinamento:

\=IF(Preenchimento\!$F$5="Online","Online",IF(Preenchimento\!$F$5=1,F3,IF(Preenchimento\!$F$5=2,(SUM(F3:G3)/2),"-")))

### **10.2. Avaliações**

Média interna da aba:

\=IF(OR(C4="falta",C4=""),"-",IF(C4="presente",IF(SUM(E4:X4)/20=0,"-",SUM(E4:X4)/20),"-"))

Observação técnica: essa média divide sempre por 20\. Já `Resultados` usa `AVERAGE`, que ignora vazios/textos. Portanto, `Avaliações!Y:Y` e `Resultados!E:F` podem divergir quando há células vazias.

Decisão operacional: toda sessão deve ser preenchida. Células vazias em notas devem ser tratadas como erro de preenchimento, pois podem gerar distorção.

Função operacional exata de `Avaliações!Y:Y`: `[A confirmar]`. Ela parece funcionar como média interna/conferência visual da aba, enquanto a nota final em `Resultados` usa médias próprias por bloco.

### **10.3. DPO**

Média ponderada do DPO:

\=IF(OR('Avaliações'\!C4="falta",'Avaliações'\!C4=""),"-",IF('Avaliações'\!C4="presente",IF(SUM(E3\*0.25,F3\*0.5,G3\*0.125,H3\*0.125)=0,"-",SUM(E3\*0.25,F3\*0.5,G3\*0.125,H3\*0.125))))

Hardcodes:

Ortografia \= 0,25

Conteúdo \= 0,50

Apresentação \= 0,125

Respeito ao prazo \= 0,125

### **10.4. Noite Cultural**

Busca por grupo:

\=IFERROR(VLOOKUP($E4,$O$4:$V$29,2),"-")

Média individual:

\=IF(OR('Avaliações'\!C4="falta",'Avaliações'\!C4=""),"-",IF('Avaliações'\!C4="presente",IF(IF(AND(F4="",G4="",H4=""),SUM(I4:L4)/4,SUM(F4:L4)/7)=0,"-",IF(AND(F4="",G4="",H4=""),SUM(I4:L4)/4,SUM(F4:L4)/7))))

Hardcodes:

Se Palco vazio: média de I:L dividida por 4\.

Se Palco preenchido: média de F:L dividida por 7\.

Se `Preenchimento!F7 = Stand`, Palco pode ser ignorado no fluxo operacional normal, pois não será utilizado. Não é necessário complicar o preenchimento se os campos de Palco não forem usados.

### **10.5. Resultados**

Avaliação 1:

\=IFERROR(AVERAGE('Avaliações'\!E4:N4),"Sem Nota")

Avaliação 2:

\=IFERROR(AVERAGE('Avaliações'\!O4:X4),"Sem Nota")

Noite Cultural:

\=IF(NC\!M4="-","-",IF(Preenchimento\!$F$7="Stand e Apresentação de Palco",(SUM(NC\!F4:L4)/7),IF(Preenchimento\!$F$7="Stand",(SUM(NC\!I4:L4)/4),"S.N.C")))

Nota final:

\=IFERROR(IF(AND(D3="online",H3="S.N.C"),1+E3\*0.25+F3\*0.25+G3\*0.1+3,IF(D3="online",1+E3\*0.25+F3\*0.25+G3\*0.1+H3\*0.3,IF(H3="S.N.C",D3\*0.1+E3\*0.25+F3\*0.25+G3\*0.1+3,D3\*0.1+E3\*0.25+F3\*0.25+G3\*0.1+H3\*0.3))),"-")

Hardcodes:

Treinamentos \= 0,10

Avaliação 1 \= 0,25

Avaliação 2 \= 0,25

DPO \= 0,10

Noite Cultural \= 0,30

Online \= 1 ponto fixo

Sem Noite Cultural \= 3 pontos fixos

Interpretação operacional:

Online \= 1 ponto fixo porque Treinamentos valem 10% da nota final.

S.N.C \= 3 pontos fixos porque Noite Cultural vale 30% da nota final.

Esses valores funcionam como compensação técnica para blocos não aplicáveis.

## **11\. Pesos e travas hardcoded**

### **11.1. Pesos hardcoded em fórmulas**

| Local | Hardcode | Significado |
| :---- | ----: | :---- |
| `Treinamentos!F:G` | `10`, `0` | presença `P` \= 10; `F` \= 0 |
| `Treinamentos!H:H` | divisão por `2` | média de dois treinamentos |
| `Avaliações!Y:Y` | divisão por `20` | média das 20 notas de eixo/sessão |
| `DPO!I:I` | `0.25`, `0.5`, `0.125`, `0.125` | pesos internos do DPO |
| `NC!M:M` | divisão por `4` ou `7` | média por tipo de Noite Cultural |
| `Resultados!E:E` | `AVERAGE(E:N)` | Avaliação 1 |
| `Resultados!F:F` | `AVERAGE(O:X)` | Avaliação 2 |
| `Resultados!I:I` | `0.1`, `0.25`, `0.25`, `0.1`, `0.3` | pesos finais |
| `Resultados!I:I` | `1` | bloco de treinamento online |
| `Resultados!I:I` | `3` | bloco sem Noite Cultural |

### **11.2. Pesos confirmados**

| Peso | Status |
| :---- | :---- |
| Treinamentos 10% | confirmado |
| Avaliação 1 25% | confirmado |
| Avaliação 2 25% | confirmado |
| DPO 10% | confirmado |
| Noite Cultural 30% | confirmado |
| DPO: Ortografia 25% | confirmado |
| DPO: Conteúdo 50% | confirmado |
| DPO: Apresentação 12,5% | confirmado |
| DPO: Respeito ao prazo 12,5% | confirmado |

### **11.3. Travas acadêmicas não codificadas**

| Trava acadêmica | Status técnico |
| :---- | :---- |
| bases 8/8/5/5 por eixo | não codificadas |
| evidência obrigatória para 9/10 | não codificada |
| evidência obrigatória abaixo da base | não codificada |
| Participação não compensa Contribuição | não codificada |
| revisão humana para premiação | não codificada |
| direção observacional sem delta | não codificada |

Automação não deve inferir essas travas a partir do XLSX. Elas pertencem à Matriz/Controle Avaliação.

## **12\. Validações e dropdowns**

Validações detectadas no XLSX:

| Aba | Range | Tipo | Valores/fórmula |
| :---- | :---- | :---- | :---- |
| `Preenchimento` | `F5` | lista | `Online,1,2` |
| `Preenchimento` | `F7` | lista | `Stand,Stand e Apresentação de Palco,Sem noite Cultural` |
| `Treinamentos` | `D3:E102` | lista | `P,F` |
| `Avaliações` | `C4:C103` | lista | `Presente,Falta` |
| `Avaliações` | `E4:X53` | decimal | tratar operacionalmente como 0–10 |

O operador técnico exato da validação decimal em `Avaliações!E4:X53` é detalhe de baixo impacto. Para uso operacional e documentação, tratar como validação 0–10.

Validações ausentes classificadas como falha:

| Aba | Range | Status |
| :---- | :---- | :---- |
| `Avaliações` | `E54:X103` | falha; deveria ter validação 0–10 |
| `DPO` | `E3:H102` | falha; deveria ter validação 0–10 |
| `NC` | `P4:V29` | falha; deveria ter validação 0–10 |

## **13\. Formatação condicional**

Regras detectadas:

| Aba | Range | Fórmula | Função |
| :---- | :---- | :---- | :---- |
| `Avaliações` | `C4:Y103` | `$C4="falta"` | marca ausência |
| `DPO` | `C3:I102` | `$C3="falta"` | marca ausência |
| `NC` | `C4:M103` | `$C4="falta"` | marca ausência |

Cores aproximadas registradas na auditoria:

| Aba | Cor aproximada |
| :---- | :---- |
| `Avaliações` | `#E6B8AF` |
| `DPO` | `#F4C7C3` |
| `NC` | `#F4C7C3` |

Padrões visuais:

| Padrão | Status |
| :---- | :---- |
| cinza/branco alternado | visual/legibilidade |
| cabeçalho roxo | visual institucional |
| alerta rosado de falta | funcional por formatação condicional |

## **14\. Colunas ocultas, linhas ocultas, mesclas e congelamento**

### **14.1. Colunas ocultas**

| Aba | Coluna(s) | Função |
| :---- | :---- | :---- |
| `Treinamentos` | `F:G` | conversão de presença em nota |
| `DPO` | `C` | presença puxada de `Avaliações` |
| `NC` | `C` | presença puxada de `Avaliações` |

### **14.2. Linhas ocultas**

| Aba | Linha | Observação |
| :---- | ----: | :---- |
| `Avaliações` | 1 | título técnico/visual |

### **14.3. Mesclas relevantes**

| Aba | Mesclas |
| :---- | :---- |
| `Preenchimento` | `A1:F1`, `H1:H8`, `A2:F2`, `A3:F3`, `A4:B4`, `A5:B5`, `A6:B6`, `A7:B7`, `A8:F8` |
| `Avaliações` | `D1:Y1`, `A2:B2`, `E2:I2`, `J2:N2`, `O2:S2`, `T2:X2` |
| `DPO` | `B1:I1` |
| `NC` | `A1:B1`, `D1:E1`, `F1:M1`, `N1:V1`, `F2:H2`, `I2:L2`, `P2:R2`, `S2:V2` |
| `Resultados` | `B1:I1` |
| `Classificados` | `A1:G1` |

### **14.4. Congelamento de painéis**

| Aba | Congelamento |
| :---- | :---- |
| `Preenchimento` | `xSplit=2`, `ySplit=9`, topo `C10` |
| `Treinamentos` | `xSplit=2`, `ySplit=2`, topo `C3` |
| `Avaliações` | `xSplit=2`, `ySplit=3`, topo `C4` |
| `DPO` | `xSplit=2`, `ySplit=2`, topo `C3` |
| `NC` | `xSplit=2`, `ySplit=3`, topo `C4` |
| `Resultados` | `xSplit=2`, `ySplit=2`, topo `C3` |
| `Classificados` | não detectado |
| `Distribuição de pontuação` | não detectado |

## **15\. Proteção**

Status técnico confirmado:

| Item | Estado |
| :---- | :---- |
| Proteção de aba | não detectada |
| Proteção estrutural do workbook | não detectada |
| Fórmulas ocultas/protegidas | não detectado efeito de proteção |
| Risco de sobrescrita acidental | alto |
| Motivo da ausência de proteção | inércia histórica; não há necessidade operacional confirmada |

Implicação para automação:

A automação deve tratar todos os campos calculados como read-only,

mesmo que o arquivo permita escrita neles.

Diretriz operacional para eventual proteção:

Bloquear fórmulas contra edição casual.

Manter fórmulas visíveis para manutenção técnica.

Liberar apenas ranges de entrada.

## **16\. Áreas frágeis**

| Área | Fragilidade |
| :---- | :---- |
| `Resultados!I:I` | fórmula final contém pesos e exceções hardcoded |
| `DPO!I:I` | pesos internos hardcoded |
| `NC!M:M` | depende de diferença entre Palco vazio e Palco preenchido |
| `Avaliações!Y:Y` | divide por 20; pode divergir das médias de `Resultados` |
| `Avaliações!E54:X103` | ausência de validação 0–10 classificada como falha |
| `DPO!E:H` | ausência de validação 0–10 classificada como falha |
| `NC!P:V` | ausência de validação 0–10 classificada como falha |
| `Classificados` | ranking automático/aparentemente dinâmico precisa de revisão técnica |
| `Distribuição de pontuação` | aba oculta exibe pesos, mas fórmulas usam hardcodes |
| colunas ocultas | podem ser ignoradas por automação ingênua |
| células mescladas | podem quebrar escrita baseada apenas em tabelas retangulares |

## **17\. Classificados e ranking**

A aba `Classificados` é usada em todos os eventos.

Status consolidado:

| Item | Status |
| :---- | :---- |
| A aba é usada operacionalmente para classificados | confirmado pelo uso informado |
| O ranking é automático ou deveria ser dinâmico | informado como aparente |
| A integridade técnica do ranking | precisa de revisão |
| Uso como decisão final sem conferência | não recomendado |
| Revisão humana acadêmica | deve ser preservada |

Automação não deve assumir que a ordenação de `Classificados` está tecnicamente correta sem teste específico.

## **18\. Pré-condições para execução de automação**

Antes de escrever no arquivo, validar:

1. o workbook contém todas as 8 abas esperadas;  
2. a aba `Distribuição de pontuação` está oculta ou, se visível, não foi alterada indevidamente;  
3. os cabeçalhos principais existem nos endereços esperados;  
4. os ranges de entrada existem;  
5. as fórmulas críticas existem nos ranges calculados;  
6. `Preenchimento!F5` contém valor aceito;  
7. `Preenchimento!F7` contém valor aceito;  
8. não há valores fora da escala 0–10 nos campos numéricos de entrada;  
9. não há células vazias em notas de sessões que deveriam estar preenchidas;  
10. não há escrita planejada em campos automáticos;  
11. há snapshot/backup antes da escrita;  
12. o ambiente de uso predominante é Google Sheets;  
13. se houver mais de 26 grupos de Noite Cultural, verificar versão alternativa `[A confirmar]`.

## **19\. Erros bloqueantes**

A automação deve bloquear execução, ou exigir revisão humana, se detectar:

| Erro | Motivo |
| :---- | :---- |
| aba esperada ausente | risco estrutural |
| cabeçalho canônico ausente ou deslocado | risco de escrita em coluna errada |
| fórmula crítica ausente em linha de referência | possível sobrescrita prévia |
| escrita planejada em range automático | risco de quebrar cálculo |
| nota fora de 0–10 | distorce médias |
| nota vazia em sessão que deveria estar preenchida | distorce médias; toda sessão deve ser preenchida |
| presença fora de `Presente/Falta` | quebra dependências |
| treinamento fora de `P/F` | quebra média de treinamento |
| valor de `F5` fora de `Online/1/2` | quebra lógica de treinamento |
| valor de `F7` fora dos três valores aceitos | quebra lógica de NC |
| grupo de Noite Cultural fora de `A–Z` em participantes com NC | pode impedir `VLOOKUP`; se exceder 26 grupos, verificar versão alternativa `[A confirmar]` |
| dados colados com número de colunas diferente do range | desalinhamento |
| `Resultados!I:I` com valor literal onde deveria haver fórmula | fórmula sobrescrita |
| `Classificados` usado como ranking automático sem teste de ordenação | risco de classificação incorreta |
| ausência de revisão humana na premiação | risco acadêmico |

## **20\. Snapshots recomendados antes de escrita**

Antes de qualquer escrita automatizada, registrar snapshot dos seguintes itens:

| Snapshot | Range |
| :---- | :---- |
| lista de abas e estado visível/oculto | workbook |
| parâmetros globais | `Preenchimento!F5`, `Preenchimento!F7` |
| cadastro | `Preenchimento!B10:H109` |
| presença de treinamentos | `Treinamentos!D3:E102` |
| notas de avaliação | `Avaliações!C4:X103` |
| notas DPO | `DPO!E3:H102` |
| notas NC | `NC!O4:V29` |
| resultados finais | `Resultados!B3:I102` |
| classificados | `Classificados!B3:G102` |
| fórmulas críticas | `Treinamentos!H3`, `Avaliações!Y4`, `DPO!I3`, `NC!M4`, `Resultados!I3` |
| estado da aba oculta | `Distribuição de pontuação` |
| validações em blocos de nota | `Avaliações`, `DPO`, `NC` |

## **21\. Exemplos de leitura**

### **21.1. Ler resultado final consolidado**

Range:

Resultados\!B3:I102

Campos:

Nome | Função | Treinamentos | Avaliação 1 | Avaliação 2 | DPO | Noite Cultural | Nota final

Uso: relatório final, conferência de cálculo, exportação para análise.

### **21.2. Ler notas por eixo**

Range:

Avaliações\!B4:Y103

Campos principais:

Nome | Presença | Função | Eixos por sessão | Média

Uso: diagnóstico acadêmico-numérico.

Observação: `Avaliações!Y:Y` parece funcionar como média interna/conferência visual, mas a nota final usa médias próprias em `Resultados`. Função operacional exata: `[A confirmar]`.

### **21.3. Ler Noite Cultural por grupo**

Range:

NC\!O4:V29

Campos:

Grupo | Palco Tempo | Palco Originalidade | Palco Caracterização | Stand Tempo | Stand Originalidade | Stand Conteúdo | Stand Caracterização

Uso: conferir entrada por grupo antes da distribuição individual.

### **21.4. Ler classificados**

Range:

Classificados\!B3:G102

Campos:

Nome | Turma | E-mail | Nota | Telefone | Unidade Escolar

Cuidado: o ranking automático/aparentemente dinâmico precisa de revisão técnica. Não assumir ordenação íntegra sem teste.

## **22\. Exemplos de escrita**

### **22.1. Cadastro em `Preenchimento!B10:H109`**

Nome	Série/turma	Função	Grupo Noite Cultural	E-mail	Telefone	Unidade Escolar

Nome do Delegado	2º ano A	Delegação X	A	[aluno@escola.com](mailto:aluno@escola.com)	61999999999	Escola X

### **22.2. Treinamentos em `Treinamentos!D3:E102`**

Treinamento 1	Treinamento 2

P	P

F	P

### **22.3. Avaliações em `Avaliações!C4:X103`**

Presença	D1	D2	D3	D4	D5	F1	F2	F3	F4	F5	C1	C2	C3	C4	C5	P1	P2	P3	P4	P5

Presente	8	8	8	8	8	8	8	8	8	8	5	5	5	5	5	5	5	5	5	5

### **22.4. DPO em `DPO!E3:H102`**

Ortografia	Conteúdo	Apresentação	Respeito ao prazo

8	9	8	10

### **22.5. Noite Cultural em `NC!P4:V29`**

Palco Tempo	Palco Originalidade	Palco Caracterização	Stand Tempo	Stand Originalidade	Stand Conteúdo	Stand Caracterização

8	9	8	7	8	9	8

## **23\. Testes mínimos após escrita**

Após qualquer escrita automatizada, validar:

| Teste | Critério mínimo |
| :---- | :---- |
| integridade de abas | 8 abas esperadas presentes |
| fórmulas críticas | fórmulas ainda existem em células de referência |
| ausência de erros visíveis | não encontrar `#REF!`, `#DIV/0!`, `#VALUE!`, `#NAME?`, `#N/A` em áreas calculadas |
| validação de escala | campos numéricos entre 0 e 10 |
| validação de presença | apenas `Presente/Falta` e `P/F` nos campos correspondentes |
| ausência de notas vazias | sessões existentes preenchidas |
| resultados gerados | `Resultados!I:I` retorna número ou `-`, não texto inesperado |
| DPO | `DPO!I:I` atualiza após escrita em `E:H` |
| NC | `NC!M:M` atualiza após escrita em `P:V` |
| classificados | `Classificados!E:E` reflete `Resultados!I:I` conforme fórmulas existentes |
| ranking | ordenação/filtragem de `Classificados` precisa ser testada antes de uso final |
| campos automáticos | nenhum campo automático foi substituído por valor literal indevido |
| filtros/mesclas/ocultos | filtros, mesclas, linhas/colunas ocultas e aba oculta preservados |
| ambiente Google Sheets | comportamento preservado após upload/edição no ambiente predominante |

## **24\. Limites e pontos `[A confirmar]`**

| Ponto | Status |
| :---- | :---- |
| existência de versões históricas com fórmulas diferentes | `[A confirmar]` |
| função operacional exata de `Avaliações!Y:Y` | `[A confirmar]` |
| versão alternativa para mais de 26 grupos de Noite Cultural | `[A confirmar]` |
| uso de versão numérica de grupos quando excede A–Z | `[A confirmar]` |
| limite acima de 100 participantes | sob demanda; não tratar como requisito padrão |
| arredondamento formal | sem problema operacional confirmado |
| aba `Distribuição de pontuação` | manter oculta; referência visual, não motor de cálculo |
| ranking em `Classificados` | precisa de revisão técnica antes de ser considerado íntegro |
