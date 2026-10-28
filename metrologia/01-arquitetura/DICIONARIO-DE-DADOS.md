# Dicionário de dados

Gerado de `ferramentas/modelo.py`. Uma tabela por lista do SharePoint.

## Acessos

Define quem é Laboratorista e quem é Instrumentação. Quem não está aqui (ou está inativo) entra como Colaborador, somente consulta.

**Quem edita:** Gestor do sistema (dono do site)

**Título (Title):** Nome da pessoa (só para leitura humana)

| Nome exato da coluna | Tipo no SharePoint | Configuração | Para que serve |
|---|---|---|---|
| `Pessoa` | Pessoa ou Grupo · seleção múltipla: Não | Obrigatória | A pessoa, escolhida pelo seletor de pessoas (evita erro de digitação de e-mail) |
| `Perfil` | Escolha | Obrigatória · Opções: `Laboratorista`, `Instrumentação` | Laboratorista: lança aferições. Instrumentação: registra calibrações. |
| `Letra` | Escolha | Opções: `A`, `B`, `C`, `D`, `ADM` | Letra da escala da pessoa (A, B, C, D ou ADM). Vai gravada em cada aferição. |
| `Ativo` | Sim/Não | Padrão: Sim | Desmarque para tirar o acesso sem apagar o histórico |
| `Observacao` | Linha única de texto |  | Livre |

## Equipamentos (existente)

Cadastro único de equipamentos: calibração (instrumentação) e situação da última aferição (laboratório).

**Quem edita:** Instrumentação (calibração) e gestor (cadastro). A situação da aferição é atualizada pelo fluxo F1.

**Título (Title):** TAG do equipamento (única; é o nome da imagem do QR Code: TAG.png) · **Impor valores exclusivos: Sim** (e indexar)

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
| `UltimoCertificado` | Linha única de texto |  | Número do certificado da última calibração (gravado pelo app). Vai na imagem de consulta |
| `MotivoRecusa` | Várias linhas de texto (texto sem formatação) |  |  |
| `ManutencaoSolicitada` | Sim/Não | Padrão: Não |  |
| `UltimaAfericao` | Data e hora · **Incluir hora: Não** |  | Gravada pelo fluxo F1 |
| `UltimaAfericaoResultado` | Linha única de texto |  | Conforme ou Não conforme (fluxo F1) |
| `UltimaAfericaoLancamento` | Linha única de texto |  | Código do lançamento (fluxo F1) |
| `Status` | Linha única de texto |  | Situação calculada pela rotina diária (F2): Liberado / Bloqueado · motivo |
| `ObservacaoCadastro` | Várias linhas de texto (texto sem formatação) |  | Pendências de cadastro (ex.: TAG a confirmar) |
| `Ativo` | Sim/Não | Padrão: Sim |  |

## Ensaios

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

## PontosAfericao

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

## Limites

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

## Afericoes

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
| `ModoLeitura` | Linha única de texto |  | Como o equipamento foi aberto no app: QR (scanner do app), Digitado (TAG digitada, ex.: no computador) ou Link (botão de um e-mail) |

## Lancamentos

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
| `ModoLeitura` | Linha única de texto |  | Como o equipamento foi aberto no app: QR (scanner do app), Digitado (TAG digitada, ex.: no computador) ou Link (botão de um e-mail) |

## Calibracoes (existente)

Histórico de calibrações (já existe da versão 1).

**Quem edita:** App (instrumentação)

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
| `Observacoes` | Várias linhas de texto (texto sem formatação) |  | Observação do técnico. As 3 últimas aparecem para a instrumentação na ficha do equipamento |
| `ModoLeitura` | Linha única de texto |  | Como o equipamento foi aberto no app: QR (scanner do app), Digitado (TAG digitada, ex.: no computador) ou Link (botão de um e-mail) |

## Configuracoes

Parâmetros gerais: escala das letras, horário dos turnos, massas de ensaio, e-mails dos relatórios.

**Quem edita:** Gestor do sistema

**Título (Title):** Chave da configuração · **Impor valores exclusivos: Sim** (e indexar)

| Nome exato da coluna | Tipo no SharePoint | Configuração | Para que serve |
|---|---|---|---|
| `Valor` | Linha única de texto | Obrigatória |  |
| `Descricao` | Várias linhas de texto (texto sem formatação) |  |  |
