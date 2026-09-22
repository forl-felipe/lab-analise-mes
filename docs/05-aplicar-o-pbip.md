# Como aplicar o modelo corrigido (.pbip)

O que está em `Gestão equioamentos power pbip/…SemanticModel/` substitui a pasta
`Gestão equioamentos power pbip.SemanticModel` do seu projeto.

---

## Passos

1. **Feche o Power BI Desktop.** O arquivo não pode estar aberto.
2. **Guarde a pasta atual** — renomeie para
   `Gestão equioamentos power pbip.SemanticModel.BACKUP`.
3. Descompacte o ZIP no mesmo lugar. A pasta nova tem o nome exato da antiga.
4. Abra o `.pbip` normalmente.
5. Na primeira atualização o Power BI vai pedir as credenciais do SharePoint —
   é esperado, porque as consultas foram reescritas.

## Como conferir que deu certo

| Onde | O que esperar |
|---|---|
| Painel de campos | `_Medidas` com 8 pastas (00 Cabeçalho … 99 Auxiliares) |
| Painel de campos | `_Medidas_Inspecoes` **não existe mais** |
| Exibição de Modelo | 16 tabelas, sem nenhuma `LocalDateTable_…` |
| `DimCalendario` | 19 colunas, relacionada a 5 tabelas |
| Card `% Conformidade` | **57,0%** (era 63,5%) |
| Cards `OMs Abertas` e `Lead Time Médio` | **em branco** (eram `0`) |
| Tabela da página CALIBRAÇÕES | continua preenchida (o filtro com emoji foi preservado) |

## O que mudou no modelo

**16 tabelas** (eram 22, sendo 9 tabelas de data automáticas e uma tabela de
medidas duplicada).

**18 relacionamentos** (eram 16, dos quais 9 serviam só às tabelas de data):

- removido `tbl_Calibracao[Tag] → tbl_Equipamentos[Tag]` — não casava **nenhuma**
  das 214 linhas
- removidos `tbl_Inspecoes[Equip_Tag]` e `tbl_Medicoes_Operacionais[Equip_Tag]`
  para `tbl_Equipamentos` — também zero correspondências; agora passam pelo de-para
- `tbl_Notas → tbl_Equipamentos` deixou de ser bidirecional
- criados: os dois para `DimLaboratorio` (a ponte real entre os domínios),
  `tbl_Meta_Inspecoes → DimTecnicos` (conserta `Cumprimento Meta`),
  sete para `DimCalendario` (dois inativos, para usar com `USERELATIONSHIP`)
  e os três da ponte de inspeções

**59 medidas** em uma tabela só, todas com pasta e formato.

**26 colunas novas em `tbl_Calibracao`**, com destaque para `Status Metrologico`,
`Em Risco` e as duas colunas de qualidade de data.

## Se algo der errado

Restaure a pasta `.BACKUP` e me diga a mensagem exata que apareceu. Os arquivos
são texto — dá para corrigir e regerar.

## Ainda pendente

- A pasta `.Report` não foi tocada. As sobreposições de visual e a grade do
  mockup são a Fase 2.
- Repontar o gráfico de Criticidade (página INSPEÇÕES) para
  `[Inspecoes Realizadas]` e excluir `Qtd Inspecoes por Criticidade`, que só
  continua existindo para não quebrar aquele visual.
- As 13 pendências da planilha, em `docs/03-plano-implementacao.md`.
