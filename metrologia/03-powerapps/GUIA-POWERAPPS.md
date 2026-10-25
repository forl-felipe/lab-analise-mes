# Etapa 2 · Aplicativo Power Apps

Tempo estimado: **1h30** na primeira vez.

## 2.0 O que você vai montar

Um app **novo**, chamado **Metrologia**, com 5 telas. O app da versão 1 continua funcionando até a virada: não mexa nele.

| Tela | Arquivo | Quem usa | O que faz |
|---|---|---|---|
| `scrInicio` | `tela-1-scrInicio.yaml` | Todos | Início. Laboratorista: **Agenda do turno**. Instrumentação: **Aguardando você**. Todos: lista de equipamentos com situação |
| `scrEquipamento` | `tela-2-scrEquipamento.yaml` | Todos | Abre pelo **QR Code**: LIBERADO / NÃO UTILIZAR, calibração, última aferição, justificativa e histórico |
| `scrEnsaio` | `tela-3-scrEnsaio.yaml` | Laboratorista | Formulário de **medição** (Blaine, tambor, Alpine, umidade, compressão, granulometria, tamboramento, Fisher). Calcula na hora |
| `scrVerificacao` | `tela-4-scrVerificacao.yaml` | Laboratorista | **Checklist dos peneiradores** (5 itens por peneira) |
| `scrCalibrar` | `tela-5-scrCalibrar.yaml` | Instrumentação | Registro de **calibração** |

Os formulários **não são fixos**: são montados a partir das listas **PontosAfericao** (campos) e **Limites** (critérios). Para incluir uma peneira nova, basta acrescentar uma linha na PontosAfericao. Não precisa mexer no app.

## 2.1 Criar o app

1. Em **make.powerapps.com**, confira o **ambiente** no canto superior direito (o mesmo combinado com a TI).
2. **+ Criar > Aplicativo de tela em branco**. Nome: **Metrologia**. Formato: **Telefone**. Clique em **Criar**.

## 2.2 Conectar as 9 listas

Menu da esquerda > **Dados** (cilindro) > **Adicionar dados** > **SharePoint** > escolha o site > marque:

`Acessos` · `Equipamentos` · `Ensaios` · `PontosAfericao` · `Limites` · `Afericoes` · `Lancamentos` · `Calibracoes` · `Configuracoes`

> Os nomes precisam aparecer **exatamente** assim no painel Dados.

## 2.3 Configuração do app

**Configurações > Geral > Limite de linhas de dados: 2000.**

## 2.4 Colar as telas

Para cada arquivo, **nesta ordem: 5, 4, 3, 2, 1**:

1. Abra o arquivo `.yaml` no Bloco de Notas > **Ctrl+A** > **Ctrl+C**.
2. No Power Apps, painel **Exibição em árvore** > clique com o botão direito em **Screen1** > **Colar**.

Depois das cinco:
- Apague a **Screen1**.
- Arraste **scrInicio** para o topo da árvore.
- Pode aparecer X vermelho enquanto faltam telas. Com as cinco coladas, os erros somem.

> **Plano B**, se a tela inteira não colar: crie uma tela em branco com o nome certo (ex.: `scrInicio`), selecione a tela e cole o arquivo da pasta `alternativa-somente-controles`. Depois copie o conteúdo do arquivo `...-OnVisible.txt` para a propriedade **OnVisible** da tela. Fórmulas digitadas ou coladas na barra de fórmulas do Studio em português usam `;` no lugar de `,` (veja 2.9).

## 2.5 Tela inicial pelo QR Code

Clique em **App** na árvore > propriedade **StartScreen** > cole:

Studio em português:
```
If(IsBlank(Param("equip")); scrInicio; scrEquipamento)
```
Studio em inglês:
```
If(IsBlank(Param("equip")), scrInicio, scrEquipamento)
```

## 2.6 Testar no Studio

1. Cadastre-se na lista **Acessos** como **Laboratorista**.
2. Selecione **scrInicio** e aperte **F5**. A agenda do turno deve aparecer.
3. Siga o **roteiro de testes** (`roteiro-de-testes.md`, no pacote de dados): 26 casos com o valor esperado de cada cálculo.
4. Troque seu perfil para **Instrumentação** na lista Acessos, feche o app e abra de novo: aparece **Aguardando você** e o botão **Registrar calibração**.
5. Desmarque **Ativo**: você vira **Colaborador**, só consulta.

> O perfil é lido quando o app abre. Depois de mudar a lista Acessos, feche e abra o app (ou use o botão **Atualizar** ↻ no topo do Início).

## 2.7 Publicar e compartilhar

1. **Salvar > Publicar**.
2. make.powerapps.com > Aplicativos > (…) **Metrologia** > **Compartilhar** > adicione **Todos, exceto usuários externos** como **Usuário**. Coloque mais uma pessoa como **Coproprietário** (backup).
3. (…) > **Detalhes** > copie o **Link da Web**. Ele vai no gerador de etiquetas (etapa 5) e na configuração `LinkApp` da lista Configuracoes.

## 2.8 Como o app funciona (para quem for manter)

### Perfil

```
Abre o app ──► procura o e-mail do usuário na lista Acessos (Ativo = Sim)
                 ├─ achou, Perfil = Laboratorista  ──► gblPerfil = "Laboratorista"
                 ├─ achou, Perfil = Instrumentação ──► gblPerfil = "Instrumentação"
                 └─ não achou                      ──► gblPerfil = "Colaborador"
```

O app mostra ou esconde os botões pelo `gblPerfil`. A **segurança real** é a permissão do SharePoint (grupos da etapa 1). Mesmo que alguém force um botão, a gravação é recusada.

### Turno e letra

- Turno 07x19 entre `TurnoDiaInicio` e `TurnoDiaFim` (lista Configuracoes). Antes das 7h, o ensaio pertence ao turno 19x07 **do dia anterior**.
- Letra da escala: ciclo de 8 dias a partir de `EscalaDataReferencia`, com as letras de `EscalaLetras07x19` e `EscalaLetras19x07`.
- Letra gravada na aferição: a **Letra da pessoa** na lista Acessos. Se estiver vazia, a letra da escala (igual à planilha).
- Se o laboratorista escolher outra data ou turno, o registro sai com **ForaDoTurno = Sim**.

### Agenda do turno

Ensaio **Previsto** = coluna do dia e turno marcada na lista **Ensaios** (SegD, SegN … DomN).
**Enviado** = existe um registro em **Lancamentos** com aquele ensaio, na data e turno atuais, com Situacao = Ativo.

### Cálculo e critério (igual à planilha)

| Entrada (PontosAfericao) | O laboratorista digita | O app calcula |
|---|---|---|
| `Valor` | o resultado | diferença para o Nominal |
| `Tempo` | minutos e segundos | rotação = Fator (voltas) × 60 ÷ tempo em segundos |
| `ComparativoReferencia` | a leitura do analisador | diferença para a estufa (digitada ou calculada das pesagens) |
| `Comparativo` | os dois valores | diferença entre eles |
| `Massa` | as massas das frações (granulometria e tamboramento) | os indicadores, com as mesmas fórmulas da planilha |
| `Checklist` | 5 itens por peneira | qualquer NÃO OK reprova; SEM TAG não reprova |

**Conforme:** com limites, LimInferior ≤ valor ≤ LimSuperior. Com tolerância, |valor − referência| ≤ Tolerância. Sem critério, Informativo.
Essas regras foram conferidas contra **todos os 400 registros de outubro** (`relatorio-validacao.md`).

### Envio

1. **Revisar e enviar**: monta a lista de resultados, conta os não conformes e abre o painel de revisão.
2. Se houver não conforme, **causa e descrição são obrigatórias**.
3. **Confirmar envio** grava uma linha por resultado em **Afericoes** e um resumo em **Lancamentos** (mesmo código `LCP-...`).
4. O fluxo **F1** atualiza a situação dos equipamentos e avisa a instrumentação.
5. **Desfazer** (no Início, no mesmo uso do app): apaga os resultados do último envio e marca o lançamento como Desfeito.

### Situação do equipamento (mesma regra no app, no fluxo diário e no Power BI)

1. Calibração **recusada** → Bloqueado
2. **Manutenção** solicitada → Bloqueado
3. Calibração **vencida** → Bloqueado
4. Última aferição **não conforme** (e sem calibração aprovada depois dela) → Bloqueado
5. Sem data de calibração → Liberado · Sem data de calibração
6. Vence em até `DiasAvisoCalibracao` dias → Liberado · Calibração vence em breve
7. Senão → Liberado

## 2.9 Vírgula e ponto e vírgula

- **Colando YAML**: o Power Apps converte sozinho.
- **Digitando ou colando na barra de fórmulas** com o Studio em português: troque `,` por `;` e `;` (entre comandos) por `;;`.

## 2.10 Problemas comuns

| Sintoma | Causa provável | O que fazer |
|---|---|---|
| X vermelho em `Pessoa.Email` | A coluna Pessoa de Acessos não é do tipo **Pessoa ou Grupo** | Recrie a coluna com o tipo certo |
| X vermelho em `.Value` (Perfil, Letra, Formulario, Entrada, TipoRegistro) | Coluna criada como texto, não como **Escolha** | Recrie como Escolha com as opções do guia |
| A agenda aparece vazia | Lista Ensaios sem as colunas do dia marcadas, ou Ativo = Não | Confira a grade SegD … DomN |
| Todo mundo aparece como Colaborador | E-mail da pessoa diferente do login | Escolha a pessoa pelo seletor (não digite) |
| "Não foi possível gravar os resultados" | Pessoa sem permissão na lista Afericoes | Fluxo F0 ou grupo **Metrologia - Laboratoristas** |
| Equipamento continua bloqueado depois do lançamento | O fluxo F1 ainda não rodou | Aguarde 1 a 2 minutos e use ↻ Atualizar |

## 2.11 Manutenção do dia a dia (sem mexer no app)

| Quero | Onde |
|---|---|
| Dar ou tirar acesso | Lista **Acessos** |
| Mudar uma tolerância | Lista **Limites** |
| Incluir ou trocar uma peneira, prensa ou tambor | Lista **PontosAfericao** (e **Equipamentos**) |
| Mudar a agenda de um ensaio | Lista **Ensaios** (colunas SegD … DomN) |
| Mudar a escala de letras, horário de turno ou massas de ensaio | Lista **Configuracoes** |
| Cadastrar a data de calibração | App, perfil Instrumentação (ou lista **Equipamentos**) |
