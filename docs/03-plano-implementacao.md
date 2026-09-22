# Plano de implementação

Ordem importa: cada fase depende da anterior. A Fase 1 inteira acontece **sem
tocar em nenhum visual** — ao final dela o painel atual já mostra números certos.

---

## Fase 1 · Fundação (Power Query + modelo)

### 1.1 Parâmetro e fonte única
1. Criar o parâmetro `Parametro_CaminhoBase` (Texto) com a URL do `.xlsx`.
2. Criar `Fonte_Base` → **desmarcar "Habilitar carga"**.
3. Criar as funções, todas com carga desmarcada:
   `fx_Tabela` · `fx_Config` · `fx_NormalizarTag` · `fx_ParseData` · `fx_QualidadeData`

> Efeito: 12 downloads do mesmo arquivo por atualização passam a 1.

### 1.2 Substituir as consultas existentes
Para cada uma, abrir **Editor Avançado** e colar o conteúdo do arquivo `.pq`:

| Consulta | Arquivo | Efeito principal |
|---|---|---|
| `tbl_Calibracao` | `10-` | parser de 6 formatos · status metrológico · guarda de nulo |
| `tbl_Equipamentos` | `11-` | chave normalizada · coluna `Laboratorio` alinhada |
| `tbl_Inspecoes` | `12-` | **remove 195 linhas em branco** · tipa `Qtd_NC` e `Data` |
| `tbl_Medicoes_Operacionais` | `13-` | converte decimais com vírgula em número |
| `tbl_Notas` | `14-` | coluna `Situação` com o domínio real |
| `Tbl_Paradas` | `15-` | chave normalizada |
| `tbl_Intervencoes` | `16-` | corrige `DataHora`, que estava sempre nula |
| `tbl_Grupos` | `17-` | coluna `Domínio` separando planta de laboratório |
| `tbl_Meta_Inspecoes` | `18-` | remove as 4 linhas em branco |
| `tbl_Config` | `20-` | passa a ser lida por `fx_Config` |

### 1.3 Consultas novas
`19-tbl_Sobressalentes` · `21-DimLaboratorio` · `22-DePara_Inspecao_Equipamento`

### 1.4 Desligar a data/hora automática
*Arquivo ▸ Opções ▸ Arquivo Atual ▸ Carregar Dados ▸ desmarcar "Data/hora automática"*

Remove `DateTableTemplate_…` e os 8 `LocalDateTable_…` do modelo.

### 1.5 Tabelas calculadas
Criar `DimCalendario` (`dax/01-`) e recriar `DimTecnicos` (`dax/02-`).
Renomear a coluna `[Date]` para `[Data]` e marcar como **tabela de datas**.

### 1.6 Relacionamentos

**Corrigir o que existe**
- `tbl_Notas → tbl_Equipamentos` : trocar direção **Ambos → Única**
- `tbl_Calibracao → tbl_Equipamentos` : **EXCLUIR.** Zero de 214 linhas casam —
  o relacionamento existe, mas nunca filtrou nada. Ver D1.

**Criar**

| De | Para | Card. | Obs. |
|---|---|---|---|
| `tbl_Calibracao[Laboratorio]` | `DimLaboratorio[Laboratorio]` | M:1 | a ponte real |
| `tbl_Equipamentos[Laboratorio]` | `DimLaboratorio[Laboratorio]` | M:1 | a ponte real |
| `tbl_Inspecoes[Responsavel]` | `DimTecnicos[Responsavel]` | M:1 | conserta o gráfico por responsável |
| `tbl_Meta_Inspecoes[Responsavel]` | `DimTecnicos[Responsavel]` | M:1 | conserta `Cumprimento Meta` |
| `tbl_Calibracao[DataVencimento]` | `DimCalendario[Data]` | M:1 | ativo |
| `tbl_Intervencoes[Data]` | `DimCalendario[Data]` | M:1 | ativo |
| `Tbl_Paradas[Data]` | `DimCalendario[Data]` | M:1 | ativo |
| `tbl_Inspecoes[Data]` | `DimCalendario[Data]` | M:1 | ativo |
| `tbl_Notas[Abertura]` | `DimCalendario[Data]` | M:1 | ativo |
| `tbl_Calibracao[UltimaCalibData]` | `DimCalendario[Data]` | M:1 | **inativo** |
| `tbl_Notas[Atendimento]` | `DimCalendario[Data]` | M:1 | **inativo** |

Depois de preencher o de-para (D3):
`tbl_Inspecoes[TagKey Inspecao] → DePara[TagKey Inspecao] → tbl_Equipamentos[TagKey]`

### 1.7 Medidas
1. Mover as 9 medidas de `_Medidas_Inspecoes` para `_Medidas` (propriedade
   **Tabela Inicial**, na exibição de Modelo — não quebra visual nenhum).
   Excluir a tabela `_Medidas_Inspecoes` vazia.
2. Ocultar a coluna `[Value]` de `_Medidas`.
3. Excluir: `DAXQtd Equipamentos Distintos`, `DAXInstrumentos Sem Data`,
   `DAXQtd Inspecoes` — conferi no JSON do relatório, nenhuma está em uso.
4. Repontar o gráfico de Criticidade (página INSPEÇÕES) para
   `[Inspecoes Realizadas]` e excluir `Qtd Inspecoes por Criticidade`.
5. Aplicar `dax/03-medidas.dax`: substituir as existentes, criar as novas,
   definir **Pasta de Exibição** e **Formato** de cada uma.

### 1.8 Ocultar do painel de campos
Todas as chaves (`TagKey`, `TagKey Inspecao`, `ID_*`), `tbl_Config`,
`DePara_Inspecao_Equipamento`, `Ordem Status`, `Ordem Metrologico`,
e as colunas `Ordem` das dimensões.
Ordenar por coluna: `StatusAtual` por `Ordem Status`,
`Status Metrologico` por `Ordem Metrologico`, `Laboratorio` por `Ordem`.

---

## Fase 2 · Identidade visual

1. *Exibir ▸ Temas ▸ Procurar temas* → `tema/samarco-tema.json`
2. Remover a formatação manual visual-a-visual (hoje um card tem raio de borda
   8px e sombra customizada e o card ao lado não tem nada) — com o tema
   aplicado, o padrão vem de graça.
3. Corrigir as **sobreposições** — existem nas 4 páginas:
   VISÃO GERAL 1px · CALIBRAÇÕES 17px · MANUTENÇÃO 15px · INSPEÇÕES 31px e 9px.
4. Aplicar a grade de 1920×1080 do mockup (ver `README.md`).
5. Padronizar cabeçalho (hoje 150/196/166/175px em cada página) e navegação
   (hoje em x=1320 em três páginas e x=1044 na quarta).

---

## Fase 3 · Expansão

- Páginas novas: EQUIPAMENTOS · PARADAS · SOBRESSALENTES · NOTAS/OMs · RELATÓRIOS
- Disponibilidade com a regra série/redundância — **exige** parada no grão de
  equipamento com início e fim. Hoje `Tbl_Paradas` tem "Horas Parado" agregado
  e 1 linha.
- Painel de alertas de sobressalentes — **exige** `Qtd. Atual`, `Estoque Mín.`
  e `Compatível com` preenchidos.
- Padronizar nomes de tabela (`Tbl_Paradas` → `tbl_Paradas`, prefixos
  `Dim`/`f`). Deixado para cá porque as páginas serão refeitas de qualquer forma.

---

## Pendências na origem (planilha)

Nenhuma delas se resolve no Power BI:

| # | Item | Onde | Impacto |
|---|---|---|---|
| 1 | `ID_Inspecao` contém o literal `yyyy` | `tbl_Inspeções` | inspeções e medições nunca cruzam |
| 2 | 3 equipamentos inspecionados sem cadastro | `tbl_Equipamentos` | Cross Belt, Bomba de Polpa, Divisor Rotativo |
| 3 | De-para de inspeção não preenchido | novo | 5 linhas |
| 4 | `Crítico` = "Sim" nas 34 linhas | `tbl_Equipamentos` | card "Críticos" repete o total |
| 5 | `Status Operacional` = "Operante" nas 34 | `tbl_Equipamentos` | coluna sem poder de filtro |
| 6 | 10 datas de vencimento não interpretáveis | `tbl_Calibracao` | `NOVA`, `nova`, `CALIBRAÇÃO`, `CALIBR` |
| 7 | 80 datas de última calibração sem dia | `tbl_Calibracao` | formato `mm/aa` |
| 8 | Última calibração com data futura (até 15/07/2027) | `tbl_Calibracao` | impossível |
| 9 | 2 TAGs duplicadas | `tbl_Calibracao` | `66BL98`, `66PS01` |
| 10 | 1 `NOK` fora do domínio (deveria ser `N/OK`) | `tbl_Calibracao` | |
| 11 | `tbl_Notas` vazia | `Notas Manutenção` | 2 cards sem base |
| 12 | Sobressalentes: 3 colunas vazias | `tbl_Sobressalentes` | painel de alertas inviável |
| 13 | Lista de equipamentos do formulário não vem do mestre | formulário | o de-para volta a dessincronizar |
