# Etapa 1 · SharePoint (a base de dados)

> **Resumo:** são **9 listas** no mesmo site. Duas já existem desde a versão 1 (**Equipamentos** e **Calibracoes**):
> nelas você só **acrescenta colunas**. As outras sete são novas.
> A lista **Calibradores** da versão 1 é substituída pela **Acessos**. Mantenha a antiga até o app novo entrar no ar e depois apague.

## 1.0 Duas formas de fazer

| Caminho | Quem faz | Tempo |
|---|---|---|
| **A. Script automático** `criar-estrutura.ps1` | TI ou quem tiver PnP.PowerShell | 5 minutos (cria listas, colunas, índices, grupos, permissões e importa os dados) |
| **B. Manual**, seguindo este guia | Você | 1h30 a 2h |

O resultado é o mesmo. Se for pelo **caminho A**, pule para a seção **1.4 Acessos**, que sempre é manual.

```
./criar-estrutura.ps1 -SiteUrl https://EMPRESA.sharepoint.com/sites/SITE -ClientId <id do app do PnP> -ImportarDados -PastaCsv <pasta>\csv
```

## 1.1 Regras de ouro (caminho manual)

1. **Nome da coluna exatamente como na tabela:** sem acento, sem espaço, com as mesmas maiúsculas. O app procura esses nomes.
2. **Nunca renomeie a coluna Título.** Em cada lista ela guarda uma coisa diferente (está indicado em cada tabela).
3. Crie a coluna com **+ Adicionar coluna > tipo > nome**. Para Sim/Não, Escolha e Data, ajuste as opções antes de salvar.
4. **Índices:** depois de criar as colunas, vá em **Configurações da lista > Colunas indexadas > Criar um novo índice** e indexe as colunas marcadas com **Indexar**. Sem isso, o app fica lento quando a lista passar de 5.000 itens.
5. Todas as listas no **mesmo site** (o mesmo do app atual).

## 1.2 Listas que já existem: acrescentar colunas

### `Equipamentos` (já existe)

Cadastro único de equipamentos: calibração (instrumentação) e situação da última aferição (laboratório).

Crie **somente as colunas que ainda não existem**. As da versão 1 continuam iguais.

**Título (Title):** TAG do equipamento (única; é o que vai no QR Code) · **Impor valores exclusivos: Sim** (e indexar)

| Nome exato da coluna | Tipo no SharePoint | Configuração | Para que serve |
|---|---|---|---|
| `Descricao` | Linha única de texto |  | Ex.: Tambor de abrasão |
| `TipoEquipamento` | Linha única de texto |  | Categoria: Balança, Tambor, Peneira, Prensa... |
| `FabricanteModelo` | Linha única de texto |  |  |
| `NumeroSerie` | Linha única de texto |  |  |
| `Local` | Linha única de texto |  | Área ou bancada |
| `Grupo` | Linha única de texto |  | Agrupador. Para peneiras: o peneirador (Circulares, Peneirador 05...) |
| `Conjunto` | Linha única de texto | **Indexar** | Peneiras: TAG do peneirador onde a peneira está instalada agora (vazio = fora de uso / em estoque) |
| `EnsaioAfericao` | Linha única de texto |  | Código do ensaio de aferição deste equipamento (ex.: TAM). Vazio = não é aferido. Mais de um: separar com ponto e vírgula. |
| `ControlaCalibracao` | Sim/Não | Padrão: Sim | Sim = tem calibração periódica controlada |
| `PeriodicidadeMeses` | Número | Padrão: `12` | Intervalo entre calibrações |
| `UltimaCalibracao` | Data e hora · **Incluir hora: Não** |  |  |
| `ProximaCalibracao` | Data e hora · **Incluir hora: Não** |  |  |
| `UltimoResultado` | Linha única de texto |  | Aprovada ou Recusada |
| `UltimoResponsavel` | Linha única de texto |  |  |
| `MotivoRecusa` | Várias linhas de texto (texto sem formatação) |  |  |
| `ManutencaoSolicitada` | Sim/Não | Padrão: Não |  |
| `UltimaAfericao` | Data e hora · **Incluir hora: Não** |  | Gravada pelo fluxo F1 |
| `UltimaAfericaoResultado` | Linha única de texto |  | Conforme ou Não conforme (fluxo F1) |
| `UltimaAfericaoLancamento` | Linha única de texto |  | Código do lançamento (fluxo F1) |
| `Status` | Linha única de texto |  | Situação calculada pela rotina diária (F2): Liberado / Bloqueado · motivo |
| `ObservacaoCadastro` | Várias linhas de texto (texto sem formatação) |  | Pendências de cadastro (ex.: TAG a confirmar) |
| `Ativo` | Sim/Não | Padrão: Sim |  |

### `Calibracoes` (já existe)

Histórico de calibrações (já existe da versão 1).

Crie **somente as colunas que ainda não existem**. As da versão 1 continuam iguais.

**Título (Title):** TAG do equipamento

| Nome exato da coluna | Tipo no SharePoint | Configuração | Para que serve |
|---|---|---|---|
| `EquipamentoID` | Número | **Indexar** |  |
| `DataCalibracao` | Data e hora · **Incluir hora: Não** |  |  |
| `Responsavel` | Linha única de texto |  |  |
| `Local` | Linha única de texto |  |  |
| `Resultado` | Linha única de texto |  |  |
| `MotivoRecusa` | Várias linhas de texto (texto sem formatação) |  |  |
| `ManutencaoSolicitada` | Sim/Não | Padrão: Não |  |
| `ProximaCalibracao` | Data e hora · **Incluir hora: Não** |  |  |
| `Certificado` | Linha única de texto |  |  |
| `Observacoes` | Várias linhas de texto (texto sem formatação) |  |  |

## 1.3 Listas novas

Para cada uma: **+ Novo > Lista > Lista em branco**, digite o nome exato e crie as colunas.

### `Ensaios`

Catálogo dos ensaios e agenda semanal (mesma grade da aba Config da planilha).

**Quem edita:** Gestor do sistema

**Título (Title):** Código do ensaio (BLA, TAM, ALP, UMI, COM, GRA, T515, PEN, FIS) · **Impor valores exclusivos: Sim** (e indexar)

| Nome exato da coluna | Tipo no SharePoint | Configuração | Para que serve |
|---|---|---|---|
| `Nome` | Linha única de texto | Obrigatória | Nome exibido no app |
| `NomeBD` | Linha única de texto | Obrigatória | Nome gravado na coluna Ensaio das aferições (o mesmo da planilha, para o Power BI) |
| `Formulario` | Escolha | Opções: `Medição`, `Checklist` · Padrão: `Medição` | Medição (valores) ou Checklist (peneiradores) |
| `Frequencia` | Linha única de texto |  | Texto exibido no app (ex.: Seg, qua e sáb · noite) |
| `Ordem` | Número |  | Ordem na agenda |
| `Instrucoes` | Várias linhas de texto (texto sem formatação) |  | Aparece no topo do formulário |
| `Ativo` | Sim/Não | Padrão: Sim |  |
| `SegD` | Sim/Não | Padrão: Não | Previsto na segunda · turno 07x19 (dia) |
| `SegN` | Sim/Não | Padrão: Não | Previsto na segunda · turno 19x07 (noite) |
| `TerD` | Sim/Não | Padrão: Não | Previsto na terça · turno 07x19 (dia) |
| `TerN` | Sim/Não | Padrão: Não | Previsto na terça · turno 19x07 (noite) |
| `QuaD` | Sim/Não | Padrão: Não | Previsto na quarta · turno 07x19 (dia) |
| `QuaN` | Sim/Não | Padrão: Não | Previsto na quarta · turno 19x07 (noite) |
| `QuiD` | Sim/Não | Padrão: Não | Previsto na quinta · turno 07x19 (dia) |
| `QuiN` | Sim/Não | Padrão: Não | Previsto na quinta · turno 19x07 (noite) |
| `SexD` | Sim/Não | Padrão: Não | Previsto na sexta · turno 07x19 (dia) |
| `SexN` | Sim/Não | Padrão: Não | Previsto na sexta · turno 19x07 (noite) |
| `SabD` | Sim/Não | Padrão: Não | Previsto na sábado · turno 07x19 (dia) |
| `SabN` | Sim/Não | Padrão: Não | Previsto na sábado · turno 19x07 (noite) |
| `DomD` | Sim/Não | Padrão: Não | Previsto na domingo · turno 07x19 (dia) |
| `DomN` | Sim/Não | Padrão: Não | Previsto na domingo · turno 19x07 (noite) |

### `PontosAfericao`

Cada linha é um campo do formulário de um ensaio: qual equipamento, qual parâmetro, que tipo de entrada e qual critério usar. É o equivalente às linhas da aba Lançamento. Nos peneiradores, cada linha é uma POSIÇÃO (peneirador + malha) e a coluna TAG diz qual peneira está nela.

**Quem edita:** Gestor do sistema. Instrumentação troca a peneira de uma posição pelo app

**Título (Title):** Código do ponto (ex.: TAM-01) · **Impor valores exclusivos: Sim** (e indexar)

| Nome exato da coluna | Tipo no SharePoint | Configuração | Para que serve |
|---|---|---|---|
| `Ensaio` | Linha única de texto | Obrigatória · **Indexar** | Código do ensaio (lista Ensaios) |
| `Ordem` | Número |  | Ordem no formulário |
| `Secao` | Linha única de texto |  | Subtítulo dentro do formulário (ex.: FX -16,0 +12,5 mm, Peneirador 05) |
| `Conjunto` | Linha única de texto |  | Peneiradores: TAG do peneirador desta posição. O QR Code do peneirador abre só as posições dele |
| `TAG` | Linha única de texto |  | TAG do equipamento (lista Equipamentos). Vazio quando não há equipamento cadastrado |
| `EquipamentoBD` | Linha única de texto |  | Texto gravado na coluna Equipamento das aferições (igual à planilha) |
| `Parametro` | Linha única de texto | Obrigatória |  |
| `Unidade` | Linha única de texto |  |  |
| `Entrada` | Escolha | Obrigatória · Opções: `Valor`, `Tempo`, `Comparativo`, `ComparativoReferencia`, `Massa`, `Checklist` | Como o laboratorista informa o resultado |
| `ChaveLimite` | Linha única de texto |  | Chave do critério na lista Limites. Vazio = informativo |
| `Fator` | Número |  | Tempo: número de voltas do tambor. Massa (granulometria): diâmetro médio da fração |
| `RotuloA` | Linha única de texto |  | Rótulo do 1º campo |
| `RotuloB` | Linha única de texto |  | Rótulo do 2º campo (Tempo, Comparativo, Massa) |
| `TipoRegistro` | Escolha | Opções: `Calibração`, `Verificação`, `Comparativo` | Coluna Tipo das aferições |
| `Ativo` | Sim/Não | Padrão: Sim |  |

### `Limites`

Critérios de aceitação. É AQUI que se define nominal, limites e tolerância de cada ensaio. O app usa o valor novo imediatamente.

**Quem edita:** Gestor do sistema

**Título (Title):** Chave do critério (ex.: TAMB, BLAM, COM16) · **Impor valores exclusivos: Sim** (e indexar)

| Nome exato da coluna | Tipo no SharePoint | Configuração | Para que serve |
|---|---|---|---|
| `Ensaio` | Linha única de texto |  | Nome do ensaio (como na planilha) |
| `Parametro` | Linha única de texto |  |  |
| `Unidade` | Linha única de texto |  |  |
| `Nominal` | Número |  | Valor de referência |
| `LimInferior` | Número |  | Se houver limites, Conforme quando LimInferior <= valor <= LimSuperior |
| `LimSuperior` | Número |  |  |
| `Tolerancia` | Número |  | Sem limites: Conforme quando |valor - referência| <= tolerância |
| `Fonte` | Várias linhas de texto (texto sem formatação) |  | De onde veio o critério / quem definiu |

### `Afericoes`

Um resultado por linha. Mesmo conteúdo da aba BD_Afericoes (o Power BI lê esta lista).

**Quem edita:** App (laboratoristas). Ninguém edita à mão.

**Título (Title):** Código do lançamento (ex.: LCP-261024-213501-MA) · **indexar**

| Nome exato da coluna | Tipo no SharePoint | Configuração | Para que serve |
|---|---|---|---|
| `Ensaio` | Linha única de texto | **Indexar** | Nome do ensaio (NomeBD) |
| `TipoRegistro` | Linha única de texto |  | Calibração / Verificação / Comparativo |
| `DataEnsaio` | Data e hora · **Incluir hora: Não** | **Indexar** |  |
| `Equipamento` | Linha única de texto |  | Como na planilha (ex.: analisador x Estufa) |
| `TAG` | Linha única de texto | **Indexar** | TAG do equipamento, para ligar com a lista Equipamentos |
| `Parametro` | Linha única de texto |  |  |
| `Unidade` | Linha única de texto |  |  |
| `Valor` | Número |  |  |
| `Referencia` | Número |  |  |
| `Diferenca` | Número |  |  |
| `LimInferior` | Número |  | Critério vigente no momento do registro |
| `LimSuperior` | Número |  |  |
| `Tolerancia` | Número |  |  |
| `Resultado` | Linha única de texto |  | Conforme / Não conforme / Informativo |
| `Responsavel` | Linha única de texto |  |  |
| `ResponsavelEmail` | Linha única de texto |  |  |
| `Letra` | Linha única de texto |  |  |
| `Turno` | Linha única de texto |  | Turno do ensaio: 07x19 ou 19x07 |
| `Observacao` | Várias linhas de texto (texto sem formatação) |  |  |
| `Causa` | Linha única de texto |  | Causa provável (quando não conforme) |
| `ItensNaoOK` | Linha única de texto |  | Peneiradores: itens marcados NÃO OK |
| `RegistradoEm` | Data e hora · **Incluir hora: Sim** |  |  |
| `TurnoRegistro` | Linha única de texto |  |  |
| `ForaDoTurno` | Linha única de texto |  | Sim quando lançado depois do turno do ensaio |
| `Origem` | Linha única de texto |  | App ou Planilha (importado) |

### `Lancamentos`

Um resumo por envio (substitui a aba Historico). Dispara o fluxo F1.

**Quem edita:** App (laboratoristas)

**Título (Title):** Código do lançamento (o mesmo das aferições) · **Impor valores exclusivos: Sim** (e indexar)

| Nome exato da coluna | Tipo no SharePoint | Configuração | Para que serve |
|---|---|---|---|
| `Ensaio` | Linha única de texto | **Indexar** | Nome do ensaio (NomeBD) |
| `DataEnsaio` | Data e hora · **Incluir hora: Não** | **Indexar** |  |
| `Turno` | Linha única de texto |  |  |
| `Letra` | Linha única de texto |  |  |
| `Responsavel` | Linha única de texto |  |  |
| `ResponsavelEmail` | Linha única de texto |  |  |
| `QtdResultados` | Número |  |  |
| `QtdNaoConformes` | Número |  |  |
| `TAGsNaoConformes` | Várias linhas de texto (texto sem formatação) |  |  |
| `Causa` | Linha única de texto |  |  |
| `Observacao` | Várias linhas de texto (texto sem formatação) |  |  |
| `RegistradoEm` | Data e hora · **Incluir hora: Sim** |  |  |
| `ForaDoTurno` | Linha única de texto |  |  |
| `Situacao` | Linha única de texto | Padrão: `Ativo` | Ativo ou Desfeito |
| `Origem` | Linha única de texto |  |  |

### `Configuracoes`

Parâmetros gerais: escala das letras, horário dos turnos, massas de ensaio, e-mails dos relatórios.

**Quem edita:** Gestor do sistema

**Título (Title):** Chave da configuração · **Impor valores exclusivos: Sim** (e indexar)

| Nome exato da coluna | Tipo no SharePoint | Configuração | Para que serve |
|---|---|---|---|
| `Valor` | Linha única de texto | Obrigatória |  |
| `Descricao` | Várias linhas de texto (texto sem formatação) |  |  |

## 1.4 Acessos: é aqui que se define quem é quem

> **Esta lista é o único lugar onde você define o tipo de acesso de cada pessoa.**

| Perfil no app | Como a pessoa ganha | O que ela pode fazer |
|---|---|---|
| **Colaborador** | Automático: qualquer pessoa da empresa que **não** está nesta lista, ou que está com Ativo = Não | Consultar equipamentos, validade e histórico |
| **Laboratorista** | Linha nesta lista com **Perfil = Laboratorista** e Ativo = Sim | Consultar e **lançar aferições** |
| **Instrumentação** | Linha nesta lista com **Perfil = Instrumentação** e Ativo = Sim | Consultar e **registrar calibrações** |

Crie a lista **Acessos** com estas colunas:

**Título (Title):** Nome da pessoa (só para leitura humana)

| Nome exato da coluna | Tipo no SharePoint | Configuração | Para que serve |
|---|---|---|---|
| `Pessoa` | Pessoa ou Grupo · seleção múltipla: Não | Obrigatória | A pessoa, escolhida pelo seletor de pessoas (evita erro de digitação de e-mail) |
| `Perfil` | Escolha | Obrigatória · Opções: `Laboratorista`, `Instrumentação` | Laboratorista: lança aferições. Instrumentação: registra calibrações. |
| `Letra` | Escolha | Opções: `A`, `B`, `C`, `D`, `ADM` | Letra da escala da pessoa (A, B, C, D ou ADM). Vai gravada em cada aferição. |
| `Ativo` | Sim/Não | Padrão: Sim | Desmarque para tirar o acesso sem apagar o histórico |
| `Observacao` | Linha única de texto |  | Livre |

**Para dar acesso a alguém:** abra a lista Acessos > **+ Novo** > em **Pessoa**, digite o nome e escolha no seletor > escolha o **Perfil** e a **Letra** > Salvar.
A aba **Acessos** da carga inicial tem os nomes e as letras da planilha, para você conferir.

**Para tirar o acesso:** desmarque **Ativo**. O histórico da pessoa continua.

**Permissão no SharePoint (automática):** o fluxo **F0 · Sincronizar acessos** (etapa 3) coloca e tira a pessoa do grupo certo do SharePoint sempre que você muda esta lista. Assim você mantém **um lugar só**. Sem o fluxo, adicione a pessoa também no grupo correspondente (seção 1.6).

## 1.5 Limites: é aqui que se define cada tolerância

> **Toda tolerância, limite e valor nominal de ensaio fica na lista `Limites`.** É a mesma tabela da aba BD_Limites da planilha.

| Situação | Preencha | O app considera Conforme quando |
|---|---|---|
| Faixa (ex.: rotação entre um mínimo e um máximo) | `LimInferior` e/ou `LimSuperior` (e `Nominal` para mostrar a diferença) | LimInferior ≤ valor ≤ LimSuperior |
| Diferença máxima (ex.: analisador x estufa) | `Tolerancia` (e `Nominal` quando a referência é fixa) | \|valor − referência\| ≤ Tolerância |
| Sem critério | nada | grava como **Informativo** |

**Para mudar uma tolerância:** abra a lista Limites > clique na linha do critério > altere `Tolerancia` > Salvar. **Vale a partir do próximo lançamento.** Cada aferição grava o critério que estava valendo naquele momento, então o histórico não muda.

**Qual ponto usa qual critério:** a coluna `ChaveLimite` da lista **PontosAfericao** liga cada campo do formulário a uma linha da Limites.

**Voltas dos tambores e diâmetros da granulometria** ficam na coluna `Fator` da lista **PontosAfericao**. **Massas de ensaio do tamboramento** ficam na lista **Configuracoes**.

## 1.6 Permissões (segurança real)

O app esconde os botões conforme o perfil, mas quem garante que ninguém grave por fora é a permissão do SharePoint.

1. **Configurações do site > Permissões do site > Criar grupo**: crie os dois grupos abaixo, com permissão **Leitura** no site.
2. No grupo **Visitantes** do site, inclua **Todos, exceto usuários externos**. Assim todos leem.
3. Em cada lista da tabela: **Configurações da lista > Permissões para esta lista > Parar de herdar permissões > Conceder permissões**: grupo + **Colaborar**.

| Lista | Colaborador (Visitantes) | Metrologia - Laboratoristas | Metrologia - Instrumentação | Você (Proprietários) |
|---|---|---|---|---|
| Acessos | Leitura | Leitura | Leitura | Controle total |
| Equipamentos | Leitura | Leitura | Colaborar | Controle total |
| Ensaios | Leitura | Leitura | Leitura | Controle total |
| PontosAfericao | Leitura | Leitura | Colaborar | Controle total |
| Limites | Leitura | Leitura | Leitura | Controle total |
| Afericoes | Leitura | Colaborar | Leitura | Controle total |
| Lancamentos | Leitura | Colaborar | Leitura | Controle total |
| Calibracoes | Leitura | Leitura | Colaborar | Controle total |
| Configuracoes | Leitura | Leitura | Leitura | Controle total |

> **Os fluxos do Power Automate** devem ser criados por uma conta **Proprietária** do site (a sua). Eles atualizam Equipamentos e Status mesmo quando quem lançou é laboratorista.

## 1.7 Carregar os dados iniciais (caminho manual)

O arquivo **carga-inicial.xlsx** tem uma aba por lista, com as colunas na mesma ordem.

1. Abra a lista no SharePoint > **Editar no modo de grade**.
2. Confira se a ordem das colunas na grade é igual à da aba. Se não for, arraste as colunas da grade.
3. No Excel, selecione da linha 2 até a última (sem cabeçalho) e copie.
4. Clique na primeira célula vazia da grade e cole (Ctrl+V). Aguarde salvar e clique em **Sair do modo de grade**.

**Ordem:** Configuracoes > Limites > Ensaios > PontosAfericao > Equipamentos > Lancamentos > Afericoes (400 linhas: cole em 2 ou 3 partes) > Acessos (pelo seletor, um a um).

## 1.8 Cores nas listas

- **Equipamentos > coluna Status > Formatar esta coluna > Modo avançado**: cole `formatacao/equipamentos-status.json`.
- **Afericoes > coluna Resultado**: cole `formatacao/afericoes-resultado.json`.

## 1.9 Conferência final

- [ ] As 9 listas estão no mesmo site, com os nomes exatos
- [ ] Índices criados (Equipamentos.Title, Afericoes.Title/DataEnsaio/TAG/Ensaio, Lancamentos.Title/DataEnsaio/Ensaio, PontosAfericao.Ensaio, Calibracoes.EquipamentoID)
- [ ] Grupos criados e permissões aplicadas
- [ ] Configuracoes: `EmailsRelatorioDiario` e `LinkApp` preenchidos
- [ ] Limites: 13 critérios conferidos com o gestor
- [ ] Equipamentos: TAGs repetidas e datas de calibração conferidas com a instrumentação (veja `ObservacaoCadastro`)
- [ ] Equipamentos da versão 1 (os que já estavam na lista): `ControlaCalibracao` = Sim e `Ativo` = Sim
- [ ] Acessos: laboratoristas e instrumentação cadastrados pelo seletor
