# -*- coding: utf-8 -*-
"""Valida output/Controle Avaliação.xlsx contra CLAUDE.md e specs/."""

import re
import unicodedata
import openpyxl

ARQ = "output/Controle Avaliação.xlsx"
erros = []


def check(cond, msg):
    if cond:
        print(f"  ok  {msg}")
    else:
        erros.append(msg)
        print(f"FALHA {msg}")


wb = openpyxl.load_workbook(ARQ)

# 1. exatamente as 4 abas, na ordem do CLAUDE.md
check(wb.sheetnames == ["Preenchimento", "Rubrica", "Registro por Sessão", "Nota"],
      f"abas exatas e na ordem: {wb.sheetnames}")

# 2. termos proibidos em qualquer célula (CLAUDE.md)
PROIBIDOS = ["humana", "id delegado", "tipo de função", "mesa", "plenária"]
achados = []
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str):
                low = unicodedata.normalize("NFC", c.value).lower()
                for p in PROIBIDOS:
                    if p in low:
                        # permitido: orientação da Rubrica citando o termo vetado
                        if p in ("mesa", "plenária") and "não usar" in low:
                            continue
                        achados.append((ws.title, c.coordinate, p))
check(not achados, f"nenhum termo proibido em células: {achados}")

pre = wb["Preenchimento"]
reg = wb["Registro por Sessão"]
rub = wb["Rubrica"]
nota = wb["Nota"]

# 3. cabeçalhos canônicos
check([pre.cell(row=9, column=j).value for j in range(2, 5)] ==
      ["Delegado", "Representação/Função", "Presença geral"],
      "Preenchimento: cabeçalhos B9:D9")

hdr3 = [reg.cell(row=3, column=j).value or "" for j in range(1, 25)]
esperado = ["", "Sessão", "Delegado", "Representação/Função", "Presença na sessão",
            "Status do registro"]
for l in ("D", "F", "C", "P"):
    esperado += [f"Evidência {l}", f"Gatilho {l}", f"Triagem {l}", f"Nota Preliminar {l}"]
esperado += ["Síntese acumulada da sessão", "Observação geral da sessão"]
check(hdr3 == esperado, "Registro: 24 cabeçalhos da linha 3 (4 campos exatos por eixo)")

grupos2 = [reg["A2"].value, reg["G2"].value, reg["K2"].value, reg["O2"].value,
           reg["S2"].value, reg["W2"].value]
check(grupos2 == ["Identificação", "Diplomacia e Respeito às Regras",
                  "Fidelidade à Função Designada", "Contribuição", "Participação",
                  "Síntese"], f"Registro: grupos da linha 2: {grupos2}")

hdr_nota = [nota.cell(row=3, column=j).value for j in range(13, 33)]
check(hdr_nota == ["1ª Sessão", "2ª Sessão", "3ª Sessão", "4ª Sessão", "5ª Sessão"] * 4,
      "Nota: 20 colunas de colagem com rótulos de sessão")
grupos_nota = [nota["M2"].value, nota["R2"].value, nota["W2"].value, nota["AB2"].value]
check(grupos_nota == ["Diplomacia e Respeito às regras", "Fidelidade à Função designada",
                      "Contribuição", "Participação"],
      f"Nota: grupos de eixo na grafia da planilha oficial: {grupos_nota}")

# 4. malha do Registro: 100 delegados x 5 sessões, sem lacunas
check(reg.max_row == 3 + 500, f"Registro termina na linha {reg.max_row} (3+500)")
for r_probe, i, s in [(4, 1, 1), (8, 1, 5), (9, 2, 1), (503, 100, 5)]:
    f_c = reg.cell(row=r_probe, column=3).value
    f_b = reg.cell(row=r_probe, column=2).value
    ok = (f'Preenchimento!$B${9 + i}' in f_c) and (f'"{s}ª Sessão"' in f_b)
    check(ok, f"Registro linha {r_probe} = delegado {i}, sessão {s}")

# 5. fórmulas de Nota Preliminar: base certa, clamp, colunas certas
for col, letra, base in [(10, "D", 8), (14, "F", 8), (18, "C", 5), (22, "P", 5)]:
    f = reg.cell(row=4, column=col).value
    gat = {"D": "H", "F": "L", "C": "P", "P": "T"}[letra]
    ok = (f'MIN(10,{base}+' in f and "MAX(0," in f and f'${gat}4' in f
          and f'"{letra}",""' in f and '$E4="Falta"' in f and f.count("IFERROR") == 8)
    check(ok, f"Nota Preliminar {letra}: base {base}, clamp 0-10, gatilho ${gat}, 8 tokens")

# 6. consolidação da Nota: referências exatas ao Registro
casos = [("M4", "J4"), ("N4", "J5"), ("Q4", "J8"), ("R4", "N4"), ("W4", "R4"),
         ("AB4", "V4"), ("M5", "J9"), ("M103", "J499"), ("AF103", "V503")]
for cel, ref in casos:
    f = nota[cel].value
    check(f"'Registro por Sessão'!${ref[0]}{ref[1:]}" in f and "ISNUMBER" in f,
          f"Nota!{cel} -> Registro!{ref}")

# 7. validações de dados
dvs_reg = list(reg.data_validations.dataValidation)
check(len(dvs_reg) == 6, f"Registro: 6 validações (presença, status, 4 triagens): {len(dvs_reg)}")
triagem_dvs = [dv for dv in dvs_reg if "Rubrica" in (dv.formula1 or "")]
check(len(triagem_dvs) == 4, "Registro: 4 dropdowns de Triagem apontando para a Rubrica")
for dv in triagem_dvs:
    m = re.match(r"'Rubrica'!\$C\$(\d+):\$C\$(\d+)", dv.formula1)
    ok = False
    if m:
        a, b = int(m.group(1)), int(m.group(2))
        vals = [rub.cell(row=x, column=3).value for x in range(a, b + 1)]
        ok = all(isinstance(v, str) and v for v in vals)
    check(ok, f"Triagem {dv.sqref}: intervalo {dv.formula1} com {b - a + 1} opções preenchidas")
dvs_pre = list(pre.data_validations.dataValidation)
check(any("Presente" in (dv.formula1 or "") for dv in dvs_pre),
      "Preenchimento: dropdown Presente/Falta")

# 8. formatação condicional de falta nas três abas de uso
for ws, col in [(pre, "$D"), (reg, "$E"), (nota, "$D")]:
    regras = [r for rng in ws.conditional_formatting for r in rng.rules]
    ok = any('="Falta"' in (r.formula[0] if r.formula else "") for r in regras)
    check(ok, f"{ws.title}: formatação condicional de Falta")

# 9. congelamento e filtros
check(pre.freeze_panes == "C10" and reg.freeze_panes == "E4" and nota.freeze_panes == "C4",
      f"congelamentos: {pre.freeze_panes}, {reg.freeze_panes}, {nota.freeze_panes}")
check(pre.auto_filter.ref == "A9:D109" and reg.auto_filter.ref == "A3:X503",
      f"filtros: {pre.auto_filter.ref}, {reg.auto_filter.ref}")
check(nota.auto_filter.ref is None, "Nota sem filtro (protege alinhamento da colagem)")

# 10. área de colagem limpa: M:AF só contém fórmula numérica/vazio, sem texto extra
sujas = []
for r in range(4, 104):
    for c in range(13, 33):
        v = nota.cell(row=r, column=c).value
        if not (isinstance(v, str) and v.startswith("=IF(ISNUMBER(")):
            sujas.append(nota.cell(row=r, column=c).coordinate)
check(not sujas, f"área de colagem M4:AF103 contém apenas as fórmulas de nota: {sujas[:5]}")

# 11. médias de conferência da Nota
f = nota["E4"].value
check("AVERAGE(M4:Q4)" in f, "Nota: Média D usa AVERAGE das 5 sessões de Diplomacia")
f = nota["I4"].value
check("AVERAGE(M4:AF4)" in f, "Nota: Média geral usa AVERAGE das 20 notas")

print()
if erros:
    print(f"{len(erros)} FALHA(S)")
    raise SystemExit(1)
print("Todas as verificações passaram.")
