# Controle de Calibração e Aferição: Laboratório Físico

Planilha `Controle_Calibracao_LCP.xlsm`. Substitui a planilha mensal de calibração (abas Tambor de Abrasão,
Gran. Fina Alpine, Calibração Blaine, Comparativo Fisher, Verificação Peneiradores, Compressão, Granulometria,
Umidade, Tamb Kg x Tamb 15,0Kg, Resultados, Controle).

## Abas

| Aba | Para que serve |
|---|---|
| **Lançamento** | Única aba de digitação. **Cabeçalho**: data, turno, letra e responsável. **Agenda do turno**: o que está previsto, o que já foi preenchido na tela e o último registro de cada ensaio. Abaixo, os **9 ensaios** com cálculo e situação (Conforme / Não conforme) na hora: Blaine, Tambor de abrasão (rpm), Alpine, Umidade, Compressão, Granulometria LTF x LCE, Tamboramento 5 x 15 kg, Peneiradores e Fisher |
| **Painel** | Indicadores do período: resultados registrados, conformidade, não conformidades e **aderência à agenda**. Também mostra a tabela por ensaio, os gráficos de tendência (Blaine, tambores, Alpine, umidade, compressão), a situação atual de cada equipamento, os peneiradores e a lista das não conformidades |
| BD_Afericoes (oculta) | Base de resultados, tabela `tbl_Afericoes`, **mesmas 17 colunas** da planilha anterior (Power BI) |
| BD_Limites (oculta) | **Critérios**: nominal, limites e tolerâncias, tabela `tbl_Limites` (Power BI). Editável: o Lançamento e o Painel usam o valor novo na hora |
| Config (oculta) | Responsáveis, letras, **agenda semanal** (X por dia e turno) e pasta do backup |

## Como usar

1. Confira data, turno, letra e responsável. Ao abrir, a data e o turno vêm pelo relógio.
2. Preencha só os ensaios feitos no turno. Os cálculos são automáticos:
   - rpm pelo tempo e número de voltas de cada tambor;
   - umidade da estufa pelas pesagens;
   - frações, −6,3 mm, RG e diâmetro médio da granulometria;
   - % do tamboramento 5 x 15 kg.
3. Clique em **Registrar lançamento**. Os resultados vão para a base, com o número do lançamento (`LCP-000123`), e a tela é limpa.
   - Ensaio sem dados não é registrado.
   - Se o mesmo ensaio já foi registrado na mesma data e turno, a planilha pergunta antes.
4. **Desfazer último lançamento** apaga da base o último registro, se algo foi lançado errado.
5. Peneiradores: o botão **Marcar itens vazios como OK** preenche os itens ainda vazios. Valores aceitos: OK, NÃO OK e SEM TAG; "nao ok" digitado vira NÃO OK.

Não é preciso criar um arquivo por mês: a base acumula. O Painel filtra por período.

## Critérios (BD_Limites)

Iguais aos da planilha anterior:

| Ensaio | Critério |
|---|---|
| Rotação dos tambores | 25 ± 1 rpm |
| Alpine | 87,4 a 89,4 % |
| Blaine manual e automático | 1986 a 2074 cm²/g |
| Calibração Star | 2076 a 2164 cm²/g |
| Fisher | ± 300 cm²/g |
| Compressão | 360 (FX 16) e 310 (FX 12,5) kgf/pel ± 20; velocidade 15 ± 2 mm/min |
| Granulometria | −6,3 mm ± 0,30; RG ± 0,15 (demais frações informativas) |
| Umidade | ± 0,05 |
| Tamboramento 5 x 15 kg | +6,3 mm ± 0,5 |

Voltas: 188 (66TA05, 66TA08, 66TA09) e 200 (66TA06, 66TA07), conforme a BD_Limites.

## Power BI

`tbl_Afericoes` e `tbl_Limites` mantêm nomes, colunas e ordem. A coluna **Origem** passa a identificar o lançamento:
`LCP-000123 | 19x07 | registrado em dd/mm/aaaa hh:mm`. A base já vem com os 26 resultados lançados em outubro na
planilha anterior.

**Importar planilha antiga** (no Painel): lê a aba BD_Afericoes de uma planilha mensal anterior e acrescenta os
resultados que ainda não estão na base.

**Backup:** ao salvar, uma cópia é gravada na pasta definida em Config (a mesma pasta da planilha anterior).

## Gerar

```
cd build
python3 build_calibracao.py           # ../Controle_Calibracao_LCP.xlsm
python3 build_calibracao.py --teste   # versão com teste automático (LibreOffice)
```
