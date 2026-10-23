# Etapa 2: aplicativo Power Apps

Tempo estimado: **40 a 60 minutos** na primeira vez.

O app tem 3 telas:

| Tela | Para quê | Quem vê |
|---|---|---|
| `scrLista` | Todos os equipamentos com status em cores, busca e quadros de contagem (Válidos, Próximos, Vencidos, Recusados, Manutenção) | Todos |
| `scrEquipamento` | Tela que abre ao ler o QR Code: dados do equipamento, status, datas, motivo da recusa e histórico | Todos |
| `scrCalibrar` | Formulário de registro da calibração (o do seu esboço) | Só quem está na lista `Calibradores` |

Como funcionam os 2 acessos com o mesmo QR Code:

```
QR Code -> link do app + ?equip=BAL-01
             |
             v
     scrEquipamento (todos veem o status e o histórico)
             |
             |-- e-mail está em "Calibradores"?  SIM -> botão "Registrar calibração" -> scrCalibrar
             '-- NÃO -> aviso "Modo visualização"
```

---

## 2.1 Criar o app em branco

1. Acesse **https://make.powerapps.com** (confira no canto superior direito se o ambiente é o da empresa).
2. **+ Criar > Aplicativo em branco > Aplicativo de tela em branco**.
3. Nome: **Controle de Calibração**. Formato: **Telefone**. Clique em **Criar**.

## 2.2 Conectar as listas do SharePoint

1. No menu da esquerda, clique no ícone **Dados** (cilindro) > **+ Adicionar dados** > procure **SharePoint**.
2. Escolha a conexão, selecione o **site** da etapa 1 e marque as 3 listas: `Equipamentos`, `Calibracoes` e `Calibradores`.
3. Clique em **Conectar**.

> Os nomes das fontes de dados precisam aparecer exatamente como `Equipamentos`, `Calibracoes` e `Calibradores`.

## 2.3 Colar as telas prontas

Os arquivos `tela-*.yaml` têm as telas completas: controles, cores, posições e todas as fórmulas.

**Para cada arquivo, nesta ordem: `tela-3-scrCalibrar.yaml`, `tela-2-scrEquipamento.yaml`, `tela-1-scrLista.yaml`:**

1. Abra o arquivo no Bloco de Notas, selecione tudo (Ctrl+A) e copie (Ctrl+C).
2. No Power Apps Studio, abra o painel **Exibição em árvore** (Tree view), clique com o botão direito em **App** (ou numa tela existente) e escolha **Colar** (ou só pressione Ctrl+V com a árvore selecionada).
3. A tela nova aparece na árvore com o nome certo (`scrCalibrar`, `scrEquipamento`, `scrLista`).

Depois de colar as três:
- **Apague a tela vazia original** (`Screen1`).
- Arraste `scrLista` para o topo da árvore.
- Alguns erros (X vermelho) podem aparecer enquanto as telas não estão todas coladas. Eles somem quando as três estiverem no app.
- Em `scrLista`, clique nos 3 pontinhos da tela > **Executar OnVisible**. Isso cria a coleção `colEquip` e tira os avisos da galeria.

### Se a colagem da tela inteira não funcionar

Use a pasta `alternativa-somente-controles`:
1. **+ Nova tela > Em branco** e renomeie para `scrLista` (ou `scrEquipamento`, `scrCalibrar`).
2. Selecione a tela, copie o arquivo `...-controles.yaml` correspondente e cole com Ctrl+V. Todos os controles entram de uma vez.
3. Copie a propriedade **OnVisible** e o **Fill** da tela, do arquivo `tela-*.yaml` (bloco `OnVisible:` logo no início), e cole na propriedade da tela.
   Atenção: veja o item 2.5 sobre o separador `;` no Studio em português.

## 2.4 Configurar o App (tela inicial pelo QR Code)

Na árvore, clique em **App** e, na propriedade **StartScreen**, cole:

Studio em **português** (separador `;`):
```
If(IsBlank(Param("equip")); scrLista; scrEquipamento)
```
Studio em **inglês** (separador `,`):
```
If(IsBlank(Param("equip")), scrLista, scrEquipamento)
```

Assim, quem abre o app pelo QR Code cai direto no equipamento. Quem abre pelo menu cai na lista geral.

## 2.5 Sobre vírgula e ponto e vírgula

- **Colando YAML** (itens 2.3 e 2.4): o Power Apps converte sozinho, não precisa fazer nada.
- **Digitando ou colando uma fórmula direto na barra de fórmulas** com o Studio em português: troque `,` por `;` e `;` (encadeamento) por `;;`.
  Exemplo: `Set(a, 1); Set(b, 2)` vira `Set(a; 1);; Set(b; 2)`.

## 2.6 Configurações recomendadas

**Configurações > Geral:**
- **Limite de linhas de dados**: 2000. Evita cortes caso a lista cresça.

**Configurações > Exibição:**
- Desmarque **Dimensionar para caber** se quiser que o app use a tela inteira do celular na vertical.

## 2.7 Testar no Studio

1. Adicione seu e-mail na lista `Calibradores`.
2. Selecione `scrLista` e aperte **F5** (Visualizar): os equipamentos aparecem com as cores.
3. Toque num equipamento > **Registrar calibração** > preencha > **Enviar / salvar os dados**.
4. Confira no SharePoint: deve existir uma linha nova em `Calibracoes`, e o equipamento em `Equipamentos` deve estar com as datas atualizadas.
5. Teste uma **Recusada** com motivo e "Solicitar manutenção" marcado. O status deve ficar azul (Manutenção solicitada).
6. Tire seu e-mail de `Calibradores`, feche e abra o app: deve aparecer "Modo visualização".

## 2.8 Publicar e compartilhar

1. **Arquivo > Salvar > Publicar esta versão**.
2. Em make.powerapps.com > **Aplicativos** > (…) do app > **Compartilhar**:
   - adicione o grupo **"Todos, exceto usuários externos"** (ou o grupo do laboratório) como **Usuário**;
   - adicione a instrumentação como **Usuário** (ou **Coproprietário** para quem vai manter o app).
   - Ao compartilhar, o Power Apps avisa que as pessoas precisam de permissão nas listas. Isso foi feito na etapa 1.3.
3. Em (…) > **Detalhes**, copie o **Link da Web**. É ele que vai no gerador de etiquetas (etapa 4).

## 2.9 Testar o link do QR Code

Cole no navegador do celular o Link da Web acrescentando `&equip=BAL-01` no final
(se o link não tiver `?`, use `?equip=BAL-01`). O app deve abrir direto na tela da BAL-01.

---

## Personalizações fáceis

| Quero... | Onde mexer |
|---|---|
| Mudar o prazo de "Próximo" de 30 para 15 dias | Procure `DateAdd(Today(), 30` nas 3 telas e no JSON da coluna Status (`2592000000` = 30 dias em milissegundos) |
| Mudar os motivos de recusa | `drpCalMotivo` > propriedade **Items** |
| Mudar cores | Propriedades **Fill**/**Color** dos controles (azul principal `RGBA(18, 59, 99, 1)`) |
| Usar um grupo do Entra ID em vez da lista Calibradores | `scrEquipamento` > **OnVisible** > troque o `LookUp(Calibradores...)` por `Office365Groups.ListGroupMembers("ID-DO-GRUPO").value` (precisa adicionar o conector Office 365 Groups) |
