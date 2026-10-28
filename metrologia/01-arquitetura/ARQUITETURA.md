# Arquitetura do sistema Metrologia

Sistema de **calibração e aferição** do laboratório físico. Substitui a planilha de aferição e integra o controle de calibração por QR Code da versão 1.

## 1. Visão geral

```
                      ┌──────────────────────────────────────────────────────┐
  Etiqueta QR         │ câmera do celular ─► imagem TAG.png no SharePoint     │  Colaborador comum
  (endereço da        │                      (biblioteca ConsultaQR)          │  (só consulta)
   imagem TAG.png) ───┤                                                      │
                      │ scanner do app ───► APP POWER APPS "Metrologia"      │  Laboratorista: aferições
                      │ (ou TAG digitada)    ficha só daquele equipamento    │  Instrumentação: calibrações
                      └──────────────────────────────┬───────────────────────┘
                                                     │ grava / lê
              ┌──────────────────────────────────────▼─────────────┐
              │      SHAREPOINT (9 listas + biblioteca ConsultaQR)  │
              │  Configuração: Acessos · Ensaios · PontosAfericao   │
              │                Limites · Configuracoes              │
              │  Cadastro:     Equipamentos                         │
              │  Registros:    Afericoes · Lancamentos · Calibracoes│
              └───────┬──────────────────────────────┬─────────────┘
                      │                              │
          ┌───────────▼───────────┐      ┌───────────▼────────────┐
          │      POWER BI         │      │     POWER AUTOMATE      │
          │  gestão: conformidade,│      │ F0 acessos → grupos     │
          │  aderência por letra, │      │ F1 alerta de NC         │
          │  tendência, calibração│      │ F2 situação + relatório │
          └───────────────────────┘      │    + aviso "venceu ontem"│
                                         │ F3 calibração registrada│
                                         └────────────────────────┘
```

## 2. Princípios de projeto

1. **O app só lança e consulta.** Toda a gestão (indicadores, filtros, histórico) fica no Power BI. O app continua simples para quem está na bancada.
2. **Configuração em listas, não no app.** Tolerâncias, agenda, campos dos formulários, escala de letras e acessos são editados nas listas. Incluir uma peneira ou mudar um limite **não exige** mexer no app.
3. **Uma regra, quatro lugares, um teste.** A regra de aceitação e a de situação do equipamento estão no app (Power Fx), no fluxo diário, no Power BI e em `ferramentas/regras.py`. O script Python é conferido contra os dados reais da planilha.
4. **Mesmas colunas da planilha.** A lista Afericoes tem o conteúdo da aba BD_Afericoes, e o Power BI recebe as mesmas colunas (`tbl_Afericoes`, `tbl_Limites`). O painel atual é aproveitado.
5. **Histórico imutável.** Cada aferição grava o critério que valia no momento (limites e tolerância). Mudar um limite depois não altera resultados antigos.
6. **Um QR Code, um equipamento, duas leituras.** Cada etiqueta tem um QR só, com o endereço da imagem do equipamento no SharePoint. A **câmera do celular** abre a imagem (consulta rápida, para qualquer pessoa). O **scanner do app** lê o mesmo QR e abre a ficha daquele equipamento (laboratorista e instrumentação). O app não tem lista nem busca: para ver outro equipamento, lê-se outro QR (ou, no computador, digita-se a TAG exata).
7. **Só ferramentas já licenciadas.** SharePoint, Power Apps e Power Automate com conectores padrão. Power BI Pro para compartilhar o painel.

## 3. Perfis e acessos

| Perfil | Quem | Como entra | No app | Permissão no SharePoint |
|---|---|---|---|---|
| **Colaborador** | Qualquer pessoa da empresa | Automático: não está na lista Acessos, ou está com Ativo = Não | **Não precisa do app.** Lê o QR com a câmera do celular e vê a **imagem** do equipamento no SharePoint: situação da calibração e validade | Leitura em tudo |
| **Laboratorista** | Quem executa os ensaios | Lista **Acessos** · Perfil = Laboratorista | Scanner do app (ou TAG digitada no computador) abre a ficha: situação da calibração e campos da aferição. Agenda do turno só de leitura; justificativa; desfazer | Leitura + **Colaborar** em Afericoes e Lancamentos |
| **Instrumentação** | Técnicos de calibração | Lista **Acessos** · Perfil = Instrumentação | Scanner do app (ou TAG digitada) abre a ficha com a última calibração e as **3 últimas observações** dos técnicos; registra a calibração; troca peneira de posição. Pendências no e-mail diário | Leitura + **Colaborar** em Calibracoes, Equipamentos e PontosAfericao |

**Um lugar só:** a lista **Acessos**. O fluxo **F0** sincroniza a lista com os grupos do SharePoint. O app lê a lista para mostrar os botões; o grupo garante que a gravação só acontece para quem pode.

Administração (listas de configuração, fluxos, app e **troca das imagens de consulta**) é dos **Proprietários do site**. Não é um quarto perfil do app.

**Modo de leitura:** cada aferição, lançamento e calibração grava em `ModoLeitura` como o equipamento foi aberto: `QR` (scanner), `Digitado` (TAG digitada, normalmente no computador) ou `Link` (botão de um e-mail). Na gestão, dá para ver o que foi registrado sem a leitura da etiqueta.

## 4. Modelo de dados

```
 Ensaios (1) ──< PontosAfericao (n) >── (1) Limites          configuração
                      │ TAG
                      ▼
               Equipamentos (1) ──< Afericoes (n) >── (1) Lancamentos     registros
                      │                 (Title = código do lançamento)
                      └──< Calibracoes (n)
 Acessos   ·   Configuracoes                                     apoio
```

| Lista | Papel | Volume esperado |
|---|---|---|
| Acessos | Quem é Laboratorista / Instrumentação | ~30 |
| Equipamentos | Cadastro único, calibração, situação | ~60 |
| Ensaios | 9 ensaios + agenda semanal (14 colunas Sim/Não) | 9 |
| PontosAfericao | Campos de cada formulário (≈ linhas da aba Lançamento) | ~70 |
| Limites | Critérios de aceitação (= BD_Limites) | ~13 |
| Afericoes | Um resultado por linha (= BD_Afericoes) | ~5.000 por ano |
| Lancamentos | Um resumo por envio | ~1.000 por ano |
| Calibracoes | Histórico de calibrações | ~100 por ano |
| Configuracoes | Escala, turnos, massas, e-mails, endereço das imagens | ~14 |
| **ConsultaQR** (biblioteca) | Uma imagem por equipamento (`TAG.png`), aberta pela câmera | ~60 imagens |

Detalhe de todas as colunas: [DICIONARIO-DE-DADOS.md](DICIONARIO-DE-DADOS.md).

## 5. Regras de negócio

### 5.1 Aceitação (igual à BD_Limites da planilha)

- **Com limites:** Conforme quando `LimInferior ≤ valor ≤ LimSuperior` (um dos lados pode estar vazio).
- **Com tolerância:** Conforme quando `|valor − referência| ≤ Tolerância`, com arredondamento de 6 casas.
- **Sem critério:** Informativo.
- **Referência:** o Nominal do critério; nos comparativos, o segundo valor; nos analisadores de umidade, a estufa.

### 5.2 Cálculos por tipo de entrada

| Entrada | Ensaios | Cálculo |
|---|---|---|
| Valor | Blaine, Alpine, Compressão | direto |
| Tempo | Tambor de abrasão | rpm = voltas × 60 ÷ (min × 60 + s) |
| ComparativoReferencia | Umidade | analisador − estufa. Estufa digitada ou `(P.inicial − P.final) ÷ (P.inicial − Tara) × 100` |
| Comparativo | Fisher | A − B |
| Massa | Granulometria LTF x LCE | % por fração; −6,3 = frações 8+9+10; RG = (f3+f4) ÷ f5; −16+12,5 = f3+f4; −16+9 = f3+f4+f5; −16+8 = f3..f6; diâmetro médio = Σ(% × diâmetro) ÷ 100 |
| Massa | Tamboramento 5 x 15 kg | +6,3 % = massa ÷ (massa de ensaio ÷ 100); −0,5 % = 100 − (+6,3 % + −6,3+0,5 %) |
| Checklist | Peneiradores | 5 itens: qualquer **NÃO OK** reprova; **SEM TAG** e **N/A** não reprovam (fórmula da planilha) |

### 5.3 Não conformidade

- Causa provável e descrição do que foi verificado são **obrigatórias** para enviar.
- O equipamento fica **Bloqueado** até uma aferição conforme ou uma calibração aprovada **posterior**.
- A instrumentação recebe e-mail na hora (F1).

### 5.4 Situação do equipamento (ordem de prioridade)

1. Calibração recusada → **Bloqueado**
2. Manutenção solicitada → **Bloqueado**
3. Calibração vencida → **Bloqueado**
4. Última aferição não conforme, sem calibração aprovada depois → **Bloqueado**
5. Sem data de calibração → Liberado · Sem data de calibração
6. Vence em até N dias (`DiasAvisoCalibracao`) → Liberado · Calibração vence em breve
7. Liberado

### 5.5 Turno, letra e "fora do turno"

- Turno 07x19 das 7h às 19h. Antes das 7h, o ensaio é do 19x07 do dia anterior.
- Letra da escala: ciclo de 8 dias (`EscalaDataReferencia`, `EscalaLetras07x19`, `EscalaLetras19x07`).
- Letra gravada: a da pessoa (Acessos). Se vazia, a da escala.
- Lançamento com data ou turno diferentes do atual → `ForaDoTurno = Sim`.

### 5.6 Aderência à agenda (Power BI)

Mesma regra da macro da planilha: **No turno**, **No dia** (outro turno fora da rotina), **Com atraso** (antes da próxima data da rotina), **Não feito**. Turno em andamento e futuro não contam. Conferido contra o painel de outubro: 91 previstos, 83 realizados.

## 5A. Consulta pela câmera: a imagem no SharePoint

O colaborador comum não usa o app. O QR Code da etiqueta é o endereço `PastaImagensQR/TAG.png`; a câmera do celular abre a imagem do equipamento, no mesmo padrão visual do app.

| Situação na imagem | Quando | Cor |
|---|---|---|
| **CALIBRADO** · válida até dd/mm/aaaa | Há data de próxima calibração e ela não passou | Verde |
| **NÃO UTILIZAR** · calibração vencida | A data de próxima calibração já passou | Vermelho |
| **A CONFIRMAR** | Equipamento sem data de calibração cadastrada | Cinza |
| **SEM CALIBRAÇÃO PERIÓDICA** | Equipamento que não tem calibração controlada (escolhido à mão no gerador) | Azul |

**Regra decidida:** a imagem mostra NÃO UTILIZAR **somente com a calibração vencida**. Uma aferição não conforme **não** muda a imagem: ela bloqueia o equipamento no app, para a equipe, que faz nova aferição ou chama a calibração.

**Ciclo de atualização (manual, uns 3 minutos por equipamento):**
```
Instrumentação registra calibração APROVADA ──► F3: e-mail para você com a linha pronta
Calibração vence (dia seguinte à validade)  ──► F2 parte E: e-mail com a linha NÃO UTILIZAR
        │
        ▼
gerador-imagens-consulta.html (cola a linha) ──► baixa TAG.png ──► ConsultaQR › Carregar › mesmo nome › Substituir
```
- Substituir mantém o mesmo endereço: a etiqueta continua valendo. A imagem antiga fica no **histórico de versões** (registro de auditoria).
- Toda imagem traz "**Imagem atualizada em**" e a frase "Se a data de validade já passou, não use o equipamento": mesmo uma imagem esquecida não engana.
- Carga inicial de todas as imagens: exibição **Imagens de consulta** da lista Equipamentos > modo de grade > copiar as linhas > colar no gerador > **Baixar todas (.zip)**.
- Calibração **recusada** não gera imagem nova (o e-mail do F3 avisa).
- O SharePoint pede login da conta da empresa para abrir a imagem.

**Por que não automatizar a imagem:** o Power Automate com conectores padrão não desenha imagens. Automatizar exigiria conector pago ou serviço externo. Com poucas calibrações por semana, o passo manual é simples e não custa nada.

## 5B. Peneiradores e peneiras

São dezenas de peneiras distribuídas em alguns conjuntos (peneiradores), verificadas várias vezes por semana. O desenho separa **posição** de **peneira**:

```
 Peneirador A (PEN-A)     ← etiqueta de PENEIRADOR (laboratorista)
   ├─ posição 19 mm   ──►  PN-0001  ← etiqueta de PENEIRA (instrumentação)
   ├─ posição 16 mm   ──►  PN-0002
   └─ ...
```

| Conceito | Onde fica | O que significa |
|---|---|---|
| **Peneirador** (conjunto) | Lista **Equipamentos**, `TipoEquipamento = Peneirador` | A máquina ou o conjunto de peneiras circulares. Tem etiqueta própria |
| **Posição** | Lista **PontosAfericao**, ensaio PEN: uma linha por peneirador + malha | Lugar fixo no peneirador. `Conjunto` = TAG do peneirador; `TAG` = peneira que está nela agora |
| **Peneira** | Lista **Equipamentos**, `TipoEquipamento = Peneira` | Item calibrado pela instrumentação. `Conjunto` = peneirador onde está instalada (vazio = fora de uso) |

### Duas etiquetas, dois públicos

| Etiqueta | Onde cola | Quem usa | O que abre |
|---|---|---|---|
| **Peneirador** (100 × 50 mm) | Na máquina ou no suporte das peneiras circulares | Laboratorista (scanner do app) | Situação de todas as peneiras daquele conjunto e o botão **Verificar as peneiras**, que abre o checklist já filtrado |
| **Peneira** (mini 50 × 25 mm) | No aro de cada peneira | Instrumentação (scanner do app) e qualquer um pela câmera | Calibração da peneira, onde está instalada, as últimas observações. Botões **Registrar calibração** e **Instalar em posição** |

As duas etiquetas seguem a regra geral: o QR é o endereço da imagem (`TAG.png`). Pela câmera, abre a imagem; pelo scanner do app, abre a ficha, e o app decide o que mostrar pelo tipo do equipamento.

### Verificação rápida (checklist)

1. Todas as peneiras **começam com os 5 itens em OK**. Se a calibração está vencida no cadastro, o item Calibração já vem **NÃO OK**.
2. Botão **✔ Marcar todos OK**, igual ao botão da planilha.
3. Para corrigir, **toque no item**: OK → NÃO OK → N/A. No item Calibração: OK → NÃO OK → SEM TAG → N/A.
4. O topo mostra o peneirador lido no QR e quantas peneiras estão não conformes. Botão **Só não conformes** para revisar antes de enviar.
5. Mostra e envia só as peneiras do peneirador lido no QR. Para verificar outro peneirador, lê-se o QR dele. A regra continua a da planilha: qualquer NÃO OK reprova; SEM TAG e N/A não reprovam.

### Troca de peneira

Quando uma peneira sai e outra entra no lugar, a **instrumentação** lê a etiqueta da peneira nova > **Instalar em posição** > escolhe a posição (o app sugere as posições da mesma malha). O app:
- coloca a peneira nova na posição (PontosAfericao);
- marca a peneira antiga como fora de uso (Conjunto vazio);
- esvazia a posição antiga da peneira nova, se ela vinha de outro lugar.

O histórico não se perde: cada verificação fica gravada com a TAG da peneira que estava na posição naquele dia. As alterações em PontosAfericao ficam no histórico de versões do SharePoint.

## 6. Fluxos principais

**Lançamento de aferição**
```
Laboratorista vê na Agenda do turno o que falta ─► vai até o equipamento e lê o QR
  ─► Lançar (formulário só com os campos daquele equipamento) ─► digita os valores (app calcula)
  ─► Revisar e enviar (justificativa se NC) ─► Afericoes (n linhas) + Lancamentos (1 linha)
  ─► F1: atualiza Equipamentos.UltimaAfericao* e Status ─► e-mail à instrumentação se houver NC
```

**Leitura do QR Code**
```
Câmera do celular ─► …/ConsultaQR/TAG.png ─► imagem (CALIBRADO / NÃO UTILIZAR)            qualquer pessoa
App › Ler QR Code ─► mesmo endereço ─► app tira a TAG do nome do arquivo ─► scrEquipamento  equipe
App › Digite a TAG (computador, sem câmera) ─► scrEquipamento (registro marcado "Digitado")
  Laboratorista: situação da calibração + Lançar · Instrumentação: últimas observações + Registrar calibração
Botão "Abrir no app" dos e-mails ─► link com ?equip=TAG ─► scrEquipamento (marcado "Link")
```

**Calibração**
```
Instrumentação ─► scanner do app (ou TAG digitada) ─► vê as 3 últimas observações ─► Registrar calibração
  ─► Calibracoes + Equipamentos (datas, resultado, certificado, status)
  ─► F3: e-mail para o gestor com a linha do gerador ─► imagem nova na ConsultaQR (Substituir)
```

**Rotina diária 07:00 (F2)**
```
Recalcula o Status de todos os equipamentos ─► agenda de ontem (feito x pendente)
  ─► NC de ontem ─► equipamentos bloqueados ou vencendo ─► e-mail
  ─► parte E: calibrações que venceram ontem ─► e-mail "atualize a imagem" (linhas NÃO UTILIZAR)
```

## 7. Onde se configura cada coisa

| O quê | Onde |
|---|---|
| Quem é laboratorista ou instrumentação | Lista **Acessos** |
| **Tolerância, limites e nominal de cada ensaio** | Lista **Limites** |
| Campos de cada formulário (equipamentos, parâmetros, tipo de entrada, critério) | Lista **PontosAfericao** |
| Voltas dos tambores, diâmetro das frações | `Fator` em **PontosAfericao** |
| Dias e turnos de cada ensaio | Lista **Ensaios** (SegD … DomN) |
| Escala de letras, horários dos turnos, massas do tamboramento, aviso de vencimento, e-mails | Lista **Configuracoes** |
| Cadastro e calibração de equipamentos | Lista **Equipamentos** e app (instrumentação) |
| **Valor padrão de uma amostra nova** (ex.: compressão) | Lista **Limites**: `Nominal`, `Tolerancia`, `LimInferior`, `LimSuperior`. Vale no próximo lançamento |
| Imagem que a câmera abre | Biblioteca **ConsultaQR** (gerador de imagens + Substituir) |
| Endereço das imagens e quem recebe o aviso de imagem | Lista **Configuracoes**: `PastaImagensQR`, `EmailsAvisoImagem` |

## 8. Capacidade e desempenho

- ~5.000 aferições por ano. Afericoes passa de 5.000 itens em cerca de 1 ano: por isso Title, DataEnsaio, TAG e Ensaio são **indexadas** desde o início.
- O app só consulta Afericoes com filtros delegáveis (TAG, data, código do lançamento). As listas de configuração (pequenas) são carregadas uma vez ao abrir o app.
- O Power BI lê as listas inteiras. Não há limite de 5.000 para ele.

## 9. Segurança, privacidade e continuidade

- Os dados ficam no tenant Microsoft 365 da empresa (SharePoint).
- Versionamento ligado nas listas (histórico de alterações) e Lixeira do SharePoint (93 dias).
- Fluxos criados por um Proprietário do site. Ter **mais um Proprietário/Coproprietário** do app e dos fluxos.
- App no **ambiente oficial** combinado com a TI (não no Sandbox).
- O repositório GitHub deste projeto **não deve conter** dados do laboratório. Os dados da carga inicial ficam no pacote entregue à parte.

## 10. Decisões pendentes

| Decisão | Impacto | Responsável |
|---|---|---|
| TAG correta das peneiras que aparecem em dois peneiradores (lista no resumo da carga) | Etiquetas e histórico | Laboratório |
| TAG da peneira sem identificação na planilha | Etiqueta | Laboratório |
| TAG oficial do conjunto de peneiras circulares (provisória: PN-CIRCULARES) | Etiqueta do peneirador | Laboratório |
| Datas de calibração de todos os equipamentos | Situação Liberado/Bloqueado | Instrumentação |
| Perfil das pessoas com letra ADM | Lista Acessos | Gestor |
| Ambiente do app (produção) | Publicação | TI |
| Licença Power BI Pro para quem vê o painel | Compartilhamento | TI |

## 11. Próximas evoluções (fase 2)

- **Certificado anexado** na calibração (fluxo "Power Apps (V2)" com arquivo, gravando na biblioteca Certificados).
- **Imagem de consulta automática** (conector Word ou serviço de imagem), se o volume de calibrações crescer.
- **App no Teams** do laboratório.
- **Fechamento de manutenção** com etapas (aberta, em manutenção, concluída, recalibrado).
