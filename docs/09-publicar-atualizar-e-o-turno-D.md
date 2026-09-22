# 09 — Turno D, publicação e atualização automática

## 1. Por que só aparecem Turno A, B e C

Fui conferir no arquivo em vez de supor. Nas 200 linhas de `tbl_Inspeções`:

| Turno | Linhas |
|---|---|
| Turno A | 3 |
| Turno B | 1 |
| Turno C | 1 |
| *(célula vazia)* | **195** |

E na aba **Listas**, a coluna `Turno` tem quatro valores: `Turno A`,
`Turno B`, `Turno C` e `Adm`. **"Turno D" não existe em lugar nenhum da
planilha** — nem nos registros, nem na lista de domínio.

O mesmo vale para os técnicos: só 5 das 200 linhas têm `Responsavel`
preenchido, com 4 nomes distintos.

Ou seja: o painel não está filtrando nada. Ele mostra exatamente o que
existe. Não há nenhuma lista de turnos escrita no código — conferi com
busca em todo o Power Query e em todo o DAX.

**O que fazer na planilha:**

1. Na aba **Listas**, coluna `Turno`, acrescentar `Turno D` (e decidir se
   `Adm` continua na mesma lista ou vira uma categoria à parte).
2. Preencher `Turno` e `Responsavel` nas linhas de inspeção que estão em
   branco.

Feito isso, `Turno D` aparece sozinho no gráfico na primeira atualização.
Não precisa mexer no painel.

### Um detalhe que estava atrapalhando

As 195 células "vazias" de `Turno` e `Responsavel` **não estão vazias**:
contêm um **espaço não separável** (U+00A0), que é o que o SharePoint
grava quando uma célula é limpa por fórmula ou colada de outro lugar.

Para o Power Query isso é conteúdo, não vazio. O efeito: a linha entrava
no modelo e o "espaço" virava uma categoria em branco no gráfico, ao lado
de Turno A, B e C.

Corrigido nesta versão: as 9 tabelas que vêm da planilha agora convertem
texto composto só de espaço (comum, não separável, tabulação, quebra de
linha) em `null` antes de decidir se a linha é vazia.

## 2. Mudanças no SharePoint chegam sozinhas ao Power BI?

**Valores: sim.** Trocar o nome de um técnico, acrescentar um turno, mudar
um status — tudo isso entra na próxima atualização, sem tocar no painel.
O modelo lê a planilha inteira a cada atualização; não há nada fixo no
código.

**Estrutura: não.** Estas quatro mudanças quebram a consulta e exigem
ajuste no Power Query:

- renomear uma **coluna** (ex.: `Responsavel` virar `Responsável`);
- renomear ou apagar uma **aba** / tabela nomeada;
- renomear a **tabela do Excel** dentro da aba (é pelo nome dela que o
  Power Query encontra os dados — `tbl_Inspeções`, `Calibração` etc.);
- mover o arquivo de pasta ou renomeá-lo.

Regra prática: **pode mexer à vontade nas linhas; avise antes de mexer nos
cabeçalhos.**

Se o arquivo mudar de pasta, a correção é em um lugar só: o parâmetro
`Parametro_CaminhoBase` (Página Inicial ▸ Transformar dados ▸ Gerenciar
Parâmetros). Foi para isso que ele existe.

## 3. Compartilhar com outras pessoas

O `.pbip` é formato de trabalho, de desenvolvedor. Para compartilhar, o
painel precisa ir para o **Power BI Service** (o site).

1. **Publicar**: Power BI Desktop ▸ Página Inicial ▸ **Publicar** ▸
   escolher um workspace.
   Publique em um **workspace de equipe**, nunca no "Meu workspace" —
   o que está no Meu workspace é pessoal e não dá para passar adiante.
2. **Licença**: quem publica e quem consome precisam de **Power BI Pro**
   (ou o workspace estar em capacidade Premium/Fabric, caso em que os
   leitores podem usar licença gratuita). Vale confirmar com a TI da
   Samarco o que a conta já tem.
3. **Dar acesso**: no workspace, em **Gerenciar acesso**, adicione as
   pessoas como *Visualizador*. Para um grupo grande, o melhor caminho é
   publicar um **App** a partir do workspace: o pessoal recebe um link
   limpo e só enxerga o que você escolher mostrar.
4. **Permissão da planilha**: quem for ver o painel **não** precisa de
   acesso ao `.xlsx`. Os dados ficam dentro do modelo publicado. Quem
   precisa de acesso é a **credencial que atualiza** (item 4 abaixo).

## 4. Atualização automática

Como a base está no **SharePoint Online** e a consulta usa `Web.Contents`
com uma URL fixa em parâmetro, **não é preciso gateway**. A nuvem lê a
nuvem.

No Power BI Service, no workspace, ache o **modelo semântico** "Gestão
equioamentos power pbip" ▸ ⋯ ▸ **Configurações**:

1. **Credenciais da fonte de dados** ▸ *Editar credenciais* ▸ método
   **OAuth2** ▸ entrar com uma conta que enxergue o arquivo no
   SharePoint. Nível de privacidade: **Organizacional**.
   Use de preferência uma **conta de serviço**, não a sua: se a sua senha
   expirar ou você sair da área, a atualização para.
2. **Atualização agendada** ▸ ligar ▸ definir fuso **(UTC-03:00)
   Brasília** e os horários.
   Sugestão para este painel: **06:00 e 13:00**, dias úteis. Os dados são
   de inspeção e calibração — não mudam de minuto a minuto.
3. Limite: **8 atualizações por dia** em Pro, 48 em Premium/Fabric.
4. Marque **"Enviar email de falha de atualização"** para o seu endereço.
   É assim que você fica sabendo que a planilha mudou de estrutura sem
   descobrir pelo painel vazio numa reunião.

O cartão **"Última Atualização"** do cabeçalho mostra a hora da última
atualização bem-sucedida (`UTCNOW()` menos 3 horas). Se ele estiver
velho, a atualização falhou.

## 5. O tamanho do cartão "Última Atualização"

O cartão foi reduzido nesta versão, de 480×56 para 300×48, e encostado à
direita do cabeçalho.

O motivo de o texto mudar de tamanho junto é que o visual **Cartão** ajusta
a fonte automaticamente para o número preencher a caixa. Para fixar a
fonte e aí sim poder redimensionar livremente:

1. Clique no cartão.
2. Painel **Formatar** (ícone de pincel) ▸ **Visual**.
3. Abra **Valor de destaque** (*Callout value*).
4. Desligue **Ajuste automático de tamanho** / defina **Tamanho da fonte**
   em, por exemplo, 14.
5. Em **Rótulo** (*Label*), dá para desligar o texto "Última Atualização"
   se quiser só a data.

Não deixei isso pronto no arquivo de propósito: `Valor de destaque` é uma
propriedade de formatação que eu não consegui confirmar em nenhum arquivo
real, e nesta conversa eu já quebrei o menu duas vezes chutando nome de
propriedade. Se preferir, ajuste em **uma** página, me mande a pasta
`.Report` de volta e eu replico o formato exato nas outras oito — foi
assim que acertamos os botões de navegação.
