# Etapa 2 · Aplicativo Power Apps

Tempo estimado: **1h30** na primeira vez.

## 2.0 O que você vai montar

Um app **novo**, chamado **Metrologia**, com 5 telas. O app da versão 1 continua funcionando até a virada: não mexa nele.

| Tela | Arquivo | Quem usa | O que faz |
|---|---|---|---|
| `scrInicio` | `tela-1-scrInicio.yaml` | Todos | Laboratorista e instrumentação: botão **Ler QR Code** (scanner) e campo **Digite a TAG** (computador). Laboratorista: agenda do turno (só leitura) e Desfazer. Instrumentação: quantas calibrações vencidas e vencendo. Colaborador: orientação para usar a câmera |
| `scrEquipamento` | `tela-2-scrEquipamento.yaml` | Laboratorista e instrumentação | Ficha **só daquele equipamento**: LIBERADO / NÃO UTILIZAR, calibração, última aferição. Laboratorista: histórico de aferições e **Lançar**. Instrumentação: **3 últimas calibrações com as observações** e **Registrar calibração**. Peneirador: posições e peneiras. Peneira: onde está instalada e troca de posição |
| `scrEnsaio` | `tela-3-scrEnsaio.yaml` | Laboratorista | Formulário de **medição** (Blaine, tambor, Alpine, umidade, compressão, granulometria, tamboramento, Fisher). Calcula na hora |
| `scrVerificacao` | `tela-4-scrVerificacao.yaml` | Laboratorista | **Checklist dos peneiradores**: tudo começa OK, botão **Marcar todos OK**, toque no item para corrigir. Só as peneiras do peneirador lido no QR |
| `scrCalibrar` | `tela-5-scrCalibrar.yaml` | Instrumentação | Registro de **calibração** (o certificado vai para a imagem de consulta; a observação aparece para o próximo técnico) |
| Leitor de QR | `leitor-qr-OnScan.txt` | Laboratorista e instrumentação | Controle **inserido à mão** na scrInicio (seção 2.5): o YAML ainda não traz esse controle |

**Regra do app: um QR Code, um equipamento.** Não existe lista nem busca de equipamentos. O QR da etiqueta é o endereço da imagem do equipamento (`…/ConsultaQR/TAG.png`): a câmera do celular abre a imagem (colaborador comum, sem app); o **scanner do app** lê o mesmo QR, tira a TAG do nome do arquivo e abre só aquele equipamento, com os botões do perfil de quem leu. Os formulários de lançamento abrem só com os campos do equipamento lido.

Os formulários **não são fixos**: são montados a partir das listas **PontosAfericao** (campos) e **Limites** (critérios). Para incluir uma peneira nova, basta acrescentar uma linha na PontosAfericao. Não precisa mexer no app.

## 2.1 Criar o app

1. Em **make.powerapps.com**, troque para o **ambiente de produção** liberado pela TI (canto superior direito). **Não use o Sandbox:** ele é individual e a TI remove o compartilhamento. O pedido é feito pelo ServiceNow ("tornar o aplicativo um produto"). Se a TI pedir que o app fique numa **solução**, crie por Soluções > a solução combinada > + Novo > Aplicativo > Aplicativo de tela.
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

> **Plano B**, se a tela inteira não colar: crie uma tela em branco com o nome certo (ex.: `scrInicio`), selecione a tela e cole o arquivo da pasta `alternativa-somente-controles`. Depois copie o conteúdo do arquivo `...-OnVisible.txt` para a propriedade **OnVisible** da tela. Fórmulas digitadas ou coladas na barra de fórmulas do Studio em português usam `;` no lugar de `,` (veja 2.11).

## 2.5 Leitor de QR Code (inserir à mão, 5 minutos)

O leitor de código de barras não vem no YAML. Você insere uma vez, por cima do cartão azul **Ler QR Code** da tela Início. Tudo o que precisa está no arquivo **`leitor-qr-OnScan.txt`** desta pasta.

1. Selecione a tela **scrInicio**. Menu **Inserir > Mídia > Leitor de código de barras** (em inglês: *Insert > Media > Barcode reader*).
2. Renomeie o controle para **`bcrIniScanner`** (exatamente assim: a fórmula usa esse nome).
3. Copie as propriedades do arquivo (X, Y, Width, Height, Visible, Text, Fill, BorderThickness). O controle fica **transparente** por cima do cartão azul; quem toca no cartão abre a câmera.
4. Propriedade **OnScan**: cole a fórmula do arquivo (há uma versão para o Studio em português e outra em inglês).
5. Na árvore, confira que **bcrIniScanner** ficou **acima** dos controles `recIniScan…` (o último inserido fica por cima).

**O que a fórmula faz:** lê o texto do QR (por exemplo `https://…/ConsultaQR/BAL-01.png`), tira a TAG do nome do arquivo (`BAL-01`), confere se o equipamento existe e está ativo e abre a ficha. Também aceita QR com o link do app (`…&equip=TAG`) ou só a TAG. QR de outra coisa mostra "Este QR Code não é de um equipamento cadastrado".

> **Câmera no celular:** use o app **Power Apps** instalado (Play Store / App Store, login da empresa) ou o navegador do celular. Na primeira vez, o celular pede permissão para a câmera: toque em **Permitir**.

## 2.6 Tela inicial por link (botões dos e-mails)

Os e-mails dos fluxos têm o botão **Abrir no app** com `…&equip=TAG`. Para o app abrir direto na ficha, clique em **App** na árvore > propriedade **StartScreen** > cole:

Studio em português:
```
If(IsBlank(Param("equip")); scrInicio; scrEquipamento)
```
Studio em inglês:
```
If(IsBlank(Param("equip")), scrInicio, scrEquipamento)
```

## 2.7 Testar no Studio

1. Cadastre-se na lista **Acessos** como **Laboratorista**.
2. Selecione **scrInicio** e aperte **F5**. A agenda do turno deve aparecer, sem itens clicáveis.
3. **Digite uma TAG** no campo da tela Início e toque em **Abrir**: abre a ficha daquele equipamento (é o caminho do computador). TAG que não existe mostra aviso.
4. **Scanner:** publique (2.8), abra o app no celular, toque em **Ler QR Code** e aponte para uma etiqueta impressa (ou para o QR na tela do gerador de etiquetas).
5. Siga o **roteiro de testes** (`roteiro-de-testes.md`, no pacote de dados): 26 casos com o valor esperado de cada cálculo.
6. Troque seu perfil para **Instrumentação** na lista Acessos, feche o app e abra de novo: o Início mostra o scanner e as contagens de calibração; a ficha mostra as **últimas observações** e **Registrar calibração**.
7. Desmarque **Ativo**: você vira **Colaborador**. O Início orienta a usar a câmera do celular (a consulta do colaborador é pela imagem, não pelo app).

> O perfil é lido quando o app abre. Depois de mudar a lista Acessos, feche e abra o app (ou use o botão **Atualizar** ↻ no topo do Início).

## 2.8 Publicar e compartilhar

1. **Salvar > Publicar**.
2. make.powerapps.com > Aplicativos > (…) **Metrologia** > **Compartilhar** > adicione **só quem registra**: laboratoristas e instrumentação (as pessoas da lista Acessos), como **Usuário**. O colaborador comum não precisa do app: ele consulta pela câmera do celular (imagem no SharePoint). Coloque mais uma pessoa como **Coproprietário** (backup).
3. (…) > **Detalhes** > copie o **Link da Web** e cole na configuração `LinkApp` da lista Configuracoes (é o botão "Abrir no app" dos e-mails). As etiquetas **não** usam esse link: elas apontam para as imagens da biblioteca ConsultaQR.
4. Para quem usa no celular: instale o app **Power Apps** (Play Store / App Store), entre com a conta da empresa e marque o **Metrologia** como favorito. Também funciona pelo navegador.

## 2.9 Celular e computador

É **o mesmo app**. No computador ele abre pelo navegador (Link da Web) ou pelo Teams (+ Adicionar um app > Power Apps > Metrologia), com o mesmo login.

| | Celular | Computador |
|---|---|---|
| Abrir um equipamento | **Ler QR Code** (câmera) | **Digite a TAG** exata e **Abrir** |
| Formato da tela | Tela cheia | O mesmo formato de celular, centralizado |
| Registro gravado como | `ModoLeitura = QR` | `ModoLeitura = Digitado` |

Não existe lista nem busca no computador: para abrir, digita-se a TAG exata. Assim a regra "um equipamento por vez" continua valendo, e a gestão consegue ver no Power BI o que foi lançado sem a leitura da etiqueta.

## 2.10 Como o app funciona (para quem for manter)

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

A agenda é **só um lembrete**: os itens não abrem nada. Para lançar, o laboratorista vai até o equipamento e lê o QR dele.

- Ensaio **previsto** = coluna do dia e turno marcada na lista **Ensaios** (SegD, SegN … DomN). Só os previstos aparecem.
- Cada ensaio lista seus equipamentos (TAGs da PontosAfericao; nos peneiradores, um item por peneirador). **✓** = já existe aferição daquele equipamento no turno atual.
- Situação do ensaio: **FEITO** (todos os equipamentos lançados), **2 DE 5** (parcial) ou **FALTA**.
- Na ficha do equipamento, o laboratorista vê o aviso **PREVISTO · FALTA** ou **FEITO NESTE TURNO ✓**.
- Coleções: `colAfTurno` (aferições do turno), `colAgendaUn` (um item por equipamento previsto) e `colAgenda` (um item por ensaio).
- A aderência oficial (Power BI) continua a regra da planilha, por ensaio.

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

### Ficha do equipamento (o que cada perfil vê)

| Perfil | Botões no rodapé |
|---|---|
| Colaborador | nenhum ("Modo consulta") |
| Laboratorista | um botão por ensaio da coluna `EnsaioAfericao` (ex.: `TAM;T515` = dois botões). Peneira instalada: **Verificar as peneiras do <peneirador>** |
| Instrumentação | **Registrar calibração**. Peneira: também **Instalar em posição**. A lista mostra as **3 últimas calibrações** com resultado, certificado e a observação do técnico (coleção `colCalHist`) |

O botão voltar da ficha leva ao Início (agenda ou "Leia o QR"). Nas telas de lançamento, voltar leva de volta à ficha do equipamento.

### Peneiradores e peneiras

- **QR do peneirador** → ficha com as posições e a situação de cada peneira (só leitura: tocar não abre outra ficha) → **Verificar peneiras** abre o checklist só com aquele conjunto.
- **QR da peneira** → ficha da peneira (calibração, onde está, verificações). Laboratorista: **Verificar o peneirador** dela. Instrumentação: **Registrar calibração** e **Instalar em posição**.
- **Checklist**: os itens ficam numa coleção (`colFormVe`, colunas I1 a I5). Cada toque troca o valor; **Marcar todos OK** usa `UpdateIf`. O checklist mostra e envia só as peneiras do peneirador lido.
- Posições vazias (sem peneira) não aparecem no checklist.

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

## 2.11 Vírgula e ponto e vírgula

- **Colando YAML**: o Power Apps converte sozinho.
- **Digitando ou colando na barra de fórmulas** com o Studio em português: troque `,` por `;` e `;` (entre comandos) por `;;`.

## 2.12 Problemas comuns

| Sintoma | Causa provável | O que fazer |
|---|---|---|
| X vermelho em `Pessoa.Email` | A coluna Pessoa de Acessos não é do tipo **Pessoa ou Grupo** | Recrie a coluna com o tipo certo |
| X vermelho em `.Value` (Perfil, Letra, Formulario, Entrada, TipoRegistro) | Coluna criada como texto, não como **Escolha** | Recrie como Escolha com as opções do guia |
| A agenda aparece vazia | Lista Ensaios sem as colunas do dia marcadas, ou Ativo = Não | Confira a grade SegD … DomN |
| Laboratorista não vê o botão Lançar | Coluna `EnsaioAfericao` do equipamento vazia | Preencha com o código do ensaio (ex.: `TAM`) |
| Todo mundo aparece como Colaborador | E-mail da pessoa diferente do login | Escolha a pessoa pelo seletor (não digite) |
| "Não foi possível gravar os resultados" | Pessoa sem permissão na lista Afericoes | Fluxo F0 ou grupo **Metrologia - Laboratoristas** |
| Equipamento continua bloqueado depois do lançamento | O fluxo F1 ainda não rodou | Aguarde 1 a 2 minutos e use ↻ Atualizar |
| Tocar em **Ler QR Code** não abre a câmera | Leitor não inserido, nome diferente de `bcrIniScanner`, ou ele ficou atrás do cartão | Seção 2.5. Na árvore, o leitor tem de ficar acima dos controles `recIniScan…` |
| "Este QR Code não é de um equipamento cadastrado" | A imagem tem nome diferente da TAG, ou o equipamento está com Ativo = Não | Confira o nome do arquivo na ConsultaQR (`TAG.png`) e a lista Equipamentos |
| A câmera abre, mas não lê | Pouca luz, etiqueta muito pequena ou suja | Aproxime até o QR ocupar o quadro. Etiqueta de peneira: imprima em boa resolução (seção 5 do guia de etiquetas) |
| O celular não pede permissão da câmera | Permissão negada antes | Configurações do celular > Apps > Power Apps (ou o navegador) > Permissões > Câmera |

## 2.13 Manutenção do dia a dia (sem mexer no app)

| Quero | Onde |
|---|---|
| Dar ou tirar acesso | Lista **Acessos** |
| Mudar uma tolerância ou o **valor padrão de uma amostra nova** (ex.: compressão) | Lista **Limites** (`Nominal`, `Tolerancia`, `LimInferior`, `LimSuperior`). Vale no próximo lançamento |
| Incluir ou trocar uma peneira, prensa ou tambor | Lista **PontosAfericao** (e **Equipamentos**) |
| Mudar a agenda de um ensaio | Lista **Ensaios** (colunas SegD … DomN) |
| Mudar a escala de letras, horário de turno ou massas de ensaio | Lista **Configuracoes** |
| Cadastrar a data de calibração | App, perfil Instrumentação (ou lista **Equipamentos**) |
| Atualizar a imagem que a câmera abre | **Gerador de imagens de consulta** + biblioteca **ConsultaQR** (Carregar > mesmo nome > Substituir) |
