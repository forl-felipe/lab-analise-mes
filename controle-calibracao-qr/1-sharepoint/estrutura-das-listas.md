# Etapa 1: SharePoint (banco de dados)

Tempo estimado: **30 a 40 minutos** se for manual, **2 minutos** com o script `criar-listas.ps1`.

## 1.1 Criar o site

1. Acesse `https://SUAEMPRESA.sharepoint.com` > **+ Criar site** > **Site de equipe**.
2. Nome sugerido: **Laboratório - Controle de Calibração**.
3. Se já existe um site do laboratório, pode usar ele mesmo.

## 1.2 Criar as 3 listas

> **REGRA DE OURO:** crie cada coluna **exatamente com o nome da tabela** (sem acento e sem espaço).
> O app usa esses nomes. Se você criar como "Última calibração", o app não encontra a coluna.
> **Não renomeie a coluna "Título" (Title).** Nas três listas ela já existe e guarda a TAG ou o e-mail.

Para criar: **+ Novo > Lista > Lista em branco**. Depois use **+ Adicionar coluna** para cada linha das tabelas.

### Lista `Equipamentos` (a planilha central)

| Coluna | Tipo no SharePoint | Configuração | Exemplo |
|---|---|---|---|
| Título *(já existe)* | Linha de texto | Em configurações da coluna: **Impor valores exclusivos = Sim** | `BAL-01` (TAG) |
| Descricao | Linha única de texto | | Balança Prix |
| TipoEquipamento | Linha única de texto | Não use "Tipo": esse nome já existe escondido no SharePoint | Balança |
| FabricanteModelo | Linha única de texto | | Prix 3 Fit |
| NumeroSerie | Linha única de texto | | 123456 |
| Local | Linha única de texto | | Laboratório - Bancada 1 |
| PeriodicidadeMeses | Número | 0 casas decimais, padrão **12** | 6 |
| UltimaCalibracao | Data e hora | **Incluir hora = Não** | 10/10/2026 |
| ProximaCalibracao | Data e hora | **Incluir hora = Não** | 10/04/2027 |
| UltimoResultado | Linha única de texto | | Aprovada / Recusada |
| UltimoResponsavel | Linha única de texto | | João Silva |
| MotivoRecusa | Várias linhas de texto | Texto **sem** formatação | Desvio na linearidade |
| ManutencaoSolicitada | Sim/Não | Padrão **Não** | |
| Status | Linha única de texto | Recebe a formatação de cores (1.4) | Válido |
| Ativo | Sim/Não | Padrão **Sim** | Desmarque para "aposentar" um equipamento |

### Lista `Calibracoes` (histórico, uma linha por calibração)

| Coluna | Tipo no SharePoint | Configuração |
|---|---|---|
| Título *(já existe)* | Linha de texto | Guarda a TAG |
| EquipamentoID | Número | 0 casas decimais. Em **Configurações da lista > Colunas indexadas**, indexe esta coluna |
| DataCalibracao | Data e hora | Incluir hora = Não |
| Responsavel | Linha única de texto | |
| Local | Linha única de texto | |
| Resultado | Linha única de texto | |
| MotivoRecusa | Várias linhas de texto | Texto sem formatação |
| ManutencaoSolicitada | Sim/Não | Padrão Não |
| ProximaCalibracao | Data e hora | Incluir hora = Não |
| Certificado | Linha única de texto | |
| Observacoes | Várias linhas de texto | Texto sem formatação |

### Lista `Calibradores` (quem pode registrar calibração)

| Coluna | Tipo | Uso |
|---|---|---|
| Título *(já existe)* | Linha de texto | **E-mail corporativo** da pessoa (ex.: `joao.silva@empresa.com`) |
| Nome | Linha única de texto | Nome da pessoa |

Para dar acesso de calibração a alguém, basta adicionar o e-mail aqui. Quem não está na lista abre o app em **modo visualização**.

## 1.3 Permissões (a segurança de verdade)

O app esconde o botão "Registrar calibração" de quem não está em `Calibradores`. Mas quem garante a segurança é a permissão do SharePoint:

| Grupo | Equipamentos | Calibracoes | Calibradores |
|---|---|---|---|
| Instrumentação (ex.: grupo "Instrumentação - Calibração") | **Editar** | **Editar** | Leitura |
| Responsável/gestor do laboratório | Controle total | Controle total | **Editar** |
| Demais colaboradores (ex.: "Todos, exceto usuários externos") | **Leitura** | **Leitura** | **Leitura** |

Como configurar em cada lista: **Configurações (engrenagem) > Configurações da lista > Permissões para esta lista > Parar de herdar permissões > Conceder permissões**.

> Todos que vão usar o app precisam de pelo menos **Leitura** nas 3 listas. Sem isso, o app abre vazio para a pessoa.

## 1.4 Cores na coluna Status (gestão visual)

1. Na lista **Equipamentos**, clique no cabeçalho da coluna **Status** > **Configurações da coluna** > **Formatar esta coluna**.
2. Clique em **Modo avançado**, apague o conteúdo e cole o conteúdo de `formatacao-coluna-status.json`.
3. Clique em **Salvar**.

A coluna passa a mostrar o status **calculado na hora** a partir das datas: Válido (verde), Próximo (amarelo, vence em até 30 dias), Vencido (vermelho claro), Recusada (vermelho), Manutenção solicitada (azul), Sem registro (cinza).

## 1.5 Carregar os equipamentos

1. Abra `modelo-equipamentos.xlsx` e preencha os equipamentos (uma linha por equipamento).
2. No SharePoint, abra a lista **Equipamentos** > **Editar no modo de grade**.
3. Copie as linhas do Excel (sem cabeçalho) e cole na grade. A ordem das colunas da visão precisa ser a mesma do Excel.
4. Se a coluna **Ativo** der erro ao colar, deixe ela fora: o padrão já é "Sim".

## 1.6 Visões úteis (opcional)

Crie visões na lista Equipamentos para o dia a dia:
- **Vencendo nos próximos 30 dias:** filtro `ProximaCalibracao` é menor ou igual a `[Hoje]+30`.
- **Vencidos:** filtro `ProximaCalibracao` é menor que `[Hoje]`.
- **Recusados / em manutenção:** filtro `UltimoResultado = Recusada` OU `ManutencaoSolicitada = Sim`.

Você também pode exportar qualquer visão para Excel (**Exportar > Exportar para Excel**) para auditorias.
