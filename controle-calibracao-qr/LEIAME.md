# Controle de Calibração por QR Code

**Laboratório de Controle de Embarque.** Etiqueta com QR Code em cada equipamento. Ao ler o código:
- a **instrumentação** registra a calibração;
- os **demais colaboradores** consultam o status;
- a **planilha central** (lista do SharePoint) se atualiza sozinha.

Tudo usa ferramentas do Microsoft 365 que a empresa já tem: **SharePoint + Power Apps + Power Automate**.

---

## Arquitetura

```
 [Etiqueta QR]  --leitura-->  [App Power Apps]  --grava-->  [SharePoint]
  BAL-01                      - visualização (todos)          - Equipamentos  (planilha central)
  link ...?equip=BAL-01       - registro (instrumentação)     - Calibracoes   (histórico)
                                                              - Calibradores  (quem pode calibrar)
                                                                    |
                                                          [Power Automate]
                                                          - alerta de recusa/manutenção
                                                          - status diário + e-mail de vencimentos
```

## Conteúdo do pacote e ordem de execução

| Etapa | Pasta | O que tem | Tempo |
|---|---|---|---|
| 1 | `1-sharepoint/` | `estrutura-das-listas.md` (passo a passo), `criar-listas.ps1` (cria tudo automaticamente, opcional), `formatacao-coluna-status.json` (cores), `modelo-equipamentos.xlsx` (carga dos equipamentos) | 30-40 min |
| 2 | `2-powerapps/` | `COMO-MONTAR-O-APP.md` + **3 telas prontas** em YAML para colar no Power Apps Studio | 40-60 min |
| 3 | `3-power-automate/` | `fluxos.md`: os 2 fluxos passo a passo, com as expressões prontas | 30 min |
| 4 | `4-etiquetas-qr/` | `gerador-etiquetas.html`: abre no navegador, gera e imprime as etiquetas (funciona offline) | 10 min |

## Por que não existe um arquivo ".msapp" pronto, como um .pbix?

O arquivo do Power Apps fica **amarrado às conexões do ambiente da sua empresa** (o site e as listas do *seu* SharePoint). Um arquivo gerado fora do seu ambiente abriria com todas as conexões quebradas.

O caminho que a Microsoft suporta para trazer código pronto é **colar as telas em YAML** no Studio. Você cria o app em branco, conecta as 3 listas e cola as 3 telas. O resultado é o app completo: layout, cores, validações e gravação. Nada precisa ser desenhado à mão.

## Regras de negócio implementadas

| Regra | Onde |
|---|---|
| **Próxima calibração** = data da calibração + `PeriodicidadeMeses` (padrão 12) | App |
| **Status:** Manutenção solicitada > Recusada > Sem registro > Vencido (próxima < hoje) > Próximo (até 30 dias) > Válido | App, coluna do SharePoint e fluxo diário |
| Recusa **exige motivo**. "Outro" exige descrição | App |
| Data da calibração **não pode ser futura** | App |
| Cada calibração vira **uma linha nova** no histórico (nada é sobrescrito) | App > `Calibracoes` |
| O equipamento recebe última/próxima data, resultado, responsável e status | App > `Equipamentos` |
| Só quem está em `Calibradores` vê o botão de registrar | App |
| Só quem tem permissão de edição consegue gravar (segurança real) | SharePoint |
| Recusa ou manutenção dispara e-mail para a manutenção | Fluxo 1 |
| E-mail diário com vencidos e a vencer em 30 dias | Fluxo 2 |

## Checklist de implantação

- [ ] Aprovação da TI para o site SharePoint, o app e os fluxos (ver rodapé do esboço da ideia)
- [ ] Site e 3 listas criados com os nomes exatos
- [ ] Permissões: instrumentação = Editar; demais = Leitura
- [ ] Equipamentos carregados (TAG única para cada um)
- [ ] E-mails da instrumentação na lista `Calibradores`
- [ ] App montado, testado (visualização e registro) e **publicado**
- [ ] App **compartilhado** com todos os usuários
- [ ] Link da Web copiado e testado com `&equip=TAG`
- [ ] Fluxos 1 e 2 criados e testados
- [ ] Etiquetas impressas (material resistente a limpeza e calor perto de estufas) e coladas
- [ ] Treinamento rápido da instrumentação (5 min: ler QR > Registrar > Salvar)

## Próximas melhorias (fase 2)

- **Anexar o certificado de calibração** (PDF/foto) no registro.
- **Painel Power BI** em cima das listas: % em dia, recusas por tipo, histórico.
- **Fixar o app no Teams** do laboratório (Teams > + Aplicativos > Power Apps).
- **Fluxo de fechamento da manutenção**: a manutenção marca "concluída" e a instrumentação é avisada para recalibrar.
- **Leitor no próprio app**: controle *Barcode reader* para ler outras etiquetas sem sair do app.
