# Comparação entre o módulo de compressão e o resultado do laboratório (MES)

**Foco da tela:** o módulo de compressão automática continua sendo o assunto principal da tela do PIMS. A comparação com o laboratório (MES) e o estado do sistema de amostragem entram como informação **complementar**: um painel discreto na tela principal e uma planilha para a análise detalhada.

## 1. O que existe hoje na tela PIMS PQ US04

O arquivo `PIMS_PQ_US04_atual_rev_1.atgraphic` foi lido por engenharia reversa do formato. O original não foi alterado. A lista completa de objetos está em `pq_us04_extraido/`, e `PIMS_PQ_US04_layout.png` mostra um desenho aproximado da tela, para conferência.

| Bloco da tela | Tag | Mapa | O que indica (interpretação) |
|---|---|---|---|
| Alimentação (t/h) | `406-FX-001` | IP_ANALOGMAP | Vazão de alimentação, 1.196,43 t/h no momento da gravação. Serve para saber se a usina está operando |
| Qualidade: Compressão | `M4710050020-0031-HHLFU` | IP_MESVALOR | Compressão do laboratório vinda do MES (275 daN/pelota na gravação) |
| Qualidade: Tamboramento | `M4710050020-0029-HHLFU` | IP_MESVALOR | Tamboramento (93,6) |
| Qualidade: Relação | `M4710050020-2021-HHLFU` | IP_MESVALOR | Relação (0,64). Significado a confirmar |
| Qualidade: menos 6,3 mm | `M4710050020-0026-HHLFU` | IP_MESVALOR | Fração abaixo de 6,3 mm (1,4) |
| Modo 07CR001 (Manual / Automático) | `U04-07CR001` | IP_PIDMAP03 | Modo do cortador primário |
| Modo 07CR002 (Manual / Automático) | `U04-07CR002` | IP_PIDMAP03 | Modo do cortador secundário |
| Painel Falha e sinóticos | `U04-07CR001`, `U04-07CR002`, `U04-07AL001`, `U04-07TP018`, `U04-07TP015` | IP_PIDMAP01 | Estado de operação ou falha de cortadores, alimentador e transportadores |
| Tendência | `U04-07CR001`, `U04-07AL001`, `U04-07CR002`, `U04-07TP018` | IP_PIDMAP01 | Histórico de estado dos equipamentos |
| Link | `PQ usina 04.apx` | | Atalho para o arquivo do Process Explorer na rede |

Leitura do fluxo pelo sinótico: transportador 07TP015, cortador primário 07CR001, alimentador 07AL001, cortador secundário 07CR002 e transportador 07TP018. O vínculo entre cada valor e seu rótulo foi feito pela posição na tela (por exemplo, "Compressão" está na mesma altura do campo de `-0031`). Vale conferir com um duplo clique no editor.

**Achado útil para a tela do módulo:** nesta tela, os campos de dados apontam para a tag, um **mapa** e o atributo `VAL` (por exemplo, `406-FX-001` com `IP_ANALOGMAP`). O roteiro da tela do módulo foi ajustado para seguir esse mesmo padrão.

## 2. O problema da comparação: 30 min contra 2 h

O módulo gera um resultado a cada 30 min. O MES traz um resultado a cada 2 h. Os dois recebem amostras do mesmo sistema de amostragem. A comparação justa é **por janela**:

1. Cada resultado do MES define uma janela de 2 h. Hipótese: o resultado das 10:00 representa a amostra composta de 08:00 a 10:00. Isso precisa ser confirmado, porque o horário gravado no MES pode ser o da coleta, o do início do turno ou o do lançamento. Um erro aqui desloca toda a comparação.
2. Na janela, calcula-se a **média dos ensaios do módulo** com qualidade Good, normalmente 4 ensaios.
3. A janela só entra na comparação se for **válida**:
   - o resultado do MES existe;
   - o módulo tem pelo menos 3 ensaios Good na janela;
   - os cortadores 07CR001 e 07CR002 ficaram **em automático durante toda a janela**, sem falha no sistema;
   - a alimentação `406-FX-001` está acima de um mínimo, ou seja, a usina estava operando.
4. Para cada janela válida: **diferença = média do módulo menos MES**.

Em manual, a amostra que vai para o laboratório e a que vai para o módulo podem não representar o mesmo material. Por isso essas janelas são separadas da comparação, e não apenas sinalizadas.

## 3. Como ler o resultado

| Indicador | Pergunta que responde |
|---|---|
| Viés médio (média das diferenças) | O módulo mede sistematicamente mais ou menos que o laboratório? |
| Teste t pareado (p-valor) | Esse viés é real ou pode ser acaso? |
| Limites de concordância (viés mais ou menos 1,96 DP das diferenças) | Numa janela qualquer, quanto os dois podem divergir? |
| Gráfico de Bland-Altman | A diferença muda conforme o nível de compressão? |
| Diferença esperada só pelo acaso das pelotas | Qual divergência é normal mesmo medindo o mesmo material? |

**Ordem de grandeza esperada.** O desvio padrão entre pelotas fica perto de 50 daN/pelota. Com 60 pelotas no laboratório e 4 ensaios de 20 pelotas no módulo (quantidades a confirmar), só o sorteio das pelotas já produz diferenças de até cerca de **18 daN/pelota** em 95 % das janelas, mesmo sem erro nenhum. Diferenças nessa faixa não indicam problema. Se o desvio padrão das diferenças ficar muito acima disso, há outra causa: amostras diferentes, calibração da prensa ou desalinhamento de horário.

**Quantidade de dados.** São necessárias pelo menos 30 janelas válidas (cerca de 3 dias de operação em automático) para uma primeira conclusão. Para uma avaliação de aceitação do módulo, 2 a 4 semanas são mais adequadas.

## 4. Onde cada coisa fica

| Onde | O que mostra | Depende de |
|---|---|---|
| **Tela principal do módulo** (destaque) | Carga média, desvio padrão, faixas, qualidade e tendência. Na tendência, os resultados do MES aparecem como quadrados cinza, discretos | Tags do módulo e tag do MES, já existentes |
| **Painel "Referência do laboratório"** (complementar, canto inferior direito) | Último resultado MES, média do módulo na mesma janela, diferença, modo dos cortadores 07CR001 e 07CR002 e situação da comparação | Tag do MES e modos, já existentes. Média, diferença e situação dependem de tags calculadas |
| **Planilha `comparacao_modulo_mes.xlsx`** | Análise do período: janelas válidas, viés, teste t, limites de concordância e gráficos | Exportação do PIMS para o Excel |

### Tags calculadas sugeridas (Aspen Calc, opcionais)

Nomes provisórios, a combinar com o administrador do PIMS. Cálculo disparado a cada novo resultado do MES:

| Tag sugerida | Lógica |
|---|---|
| `307-CMP-MOD-JAN` | Média de `307-GI-001` (só qualidade Good) entre o horário do MES menos 2 h e o horário do MES |
| `307-CMP-DIF` | `307-CMP-MOD-JAN` menos `M4710050020-0031-HHLFU`, somente se a janela for válida |
| `307-CMP-SITUACAO` | 1 = válida; 2 = amostragem em manual ou em falha; 3 = poucos ensaios do módulo; 4 = usina parada |

Sem essas tags, o painel mostra o resultado do MES e o modo dos cortadores, e a comparação numérica fica na planilha. **Recomendação:** começar assim e criar as tags calculadas depois de validar o método com algumas semanas de dados.

## 5. Planilha de comparação

`comparacao_modulo_mes.xlsx` sai preenchida com **dados simulados**, só para demonstrar o funcionamento. No exemplo, o módulo foi simulado medindo 8 daN/pelota a mais que o laboratório, e a planilha recupera esse viés (6,8 daN/pelota, significativo). As janelas em manual, com falha e com poucos ensaios foram separadas automaticamente.

Para usar com dados reais:

1. Exporte o mesmo período das tags do módulo, da tag do MES e dos modos e falhas da amostragem, mais a alimentação.
2. Cole nas abas `Modulo`, `MES` e `Amostragem` (células amarelas).
3. Confira os parâmetros na aba `Parametros`.
4. Leia a aba `Resumo`.

As colunas ocultas da aba `Janelas` (O a U) apoiam os gráficos. O `#N/A` nelas é proposital: serve para o Excel não desenhar pontos nas janelas inválidas.

## 6. Perguntas para fechar o método

1. O horário do resultado no MES marca o **fim**, o **início** da janela de 2 h ou o **lançamento** no sistema?
2. Qual cortador alimenta o módulo e qual alimenta o laboratório? Se for apenas o 07CR002, a validade pode considerar só esse cortador.
3. Quantas pelotas por ensaio o laboratório rompe? E o módulo?
4. Qual alimentação mínima de `406-FX-001` indica usina operando normalmente?
5. Qual é o significado de "Relação" (`M4710050020-2021-HHLFU`)?
