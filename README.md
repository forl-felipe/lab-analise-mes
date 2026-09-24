# Painel de Gestão de Equipamentos e Calibrações — Samarco

Power BI para gestão de equipamentos de amostragem, calibrações,
inspeções, paradas, intervenções, notas/OMs e sobressalentes dos
laboratórios LCP, LCE e LDP.

## Baixar e instalar

1. Baixe o repositório inteiro:
   **https://github.com/forl-felipe/lab-analise-mes/archive/refs/heads/claude/funny-dirac-ic9mx0.zip**
2. Extraia com o botão direito no .zip → "Extrair tudo…", destino `C:\PBI`.
   Dentro vai haver uma pasta chamada
   **`Painel Samarco`** — é o projeto inteiro, com o
   `.pbip` e as duas subpastas já com os nomes certos.
3. **Apague** a pasta antiga do projeto inteira e ponha esta no lugar.
   Apagar, não copiar por cima: ver `Painel Samarco/LEIA-ME.txt`.
4. Duplo clique no `.pbip` e clique em **Atualizar**.
5. Confira: rodapé com `· v16` e o logo da Samarco na barra lateral.
6. Aferições: cole o link da pasta do SharePoint na consulta `tbl_Afericoes`
   (passo a passo em `docs/13-v15-painel-gerencial.md`).

## Onde fica cada coisa

| Pasta | Conteúdo |
|---|---|
| `Painel Samarco/` | **o projeto pronto para instalar** |
| `docs/` | diagnóstico, decisões de arquitetura, guias e respostas |
| `powerquery/` | as consultas M, comentadas, para leitura |
| `dax/` | calendário, dimensão de técnicos e as 59 medidas |
| `tema/` | tema do Power BI e a paleta institucional (`samarco-cores.css`) |
| `entrega/` | o zip só do projeto, se preferir baixar assim |
| `afericoes/` | planilha de calibração dos operadores com as abas ocultas `BD_Afericoes` e `BD_Limites` |

## Documentação

| Documento | Responde |
|---|---|
| `docs/01-diagnostico.md` | o que havia de errado no arquivo original |
| `docs/02-decisoes-arquitetura.md` | as 10 decisões de modelagem e por quê |
| `docs/03-plano-implementacao.md` | as pendências que são da planilha |
| `docs/04` a `06` | instalação passo a passo |
| `docs/07-icones-e-navegacao.md` | menu lateral, ícones e o que quebrou no caminho |
| `docs/08-identidade-visual.md` | paleta Samarco e por que as cores de gráfico são ajustadas |
| `docs/09-publicar-atualizar-e-o-turno-D.md` | turno D, publicação, compartilhamento e atualização agendada |
| `docs/11-integracao-microsoft-forms.md` | como o formulário do Forms alimenta o painel |
| `docs/12-afericoes.md` | aferições e comparativos: a aba oculta, as regras e as decisões |
| `docs/13-v15-painel-gerencial.md` | v15/v16: Visão Geral gerencial, Equipamentos & Paradas, Aferições, metas e fontes |

## Pendências que são da planilha, não do painel

- Coluna `Crítico` com "Sim" nas 34 linhas → Equipamentos Críticos = Total.
- Aba "Notas Manutenção" vazia (só cabeçalho) → indicadores de OM em 0.
- `Turno` e `Responsavel` preenchidos em 5 de 200 linhas de inspeção;
  "Turno D" não existe na aba Listas.
- Sobressalentes com `Qtd. Atual` e `Estoque Mín.` em 1 nas 27 linhas.

Lista completa em `docs/03-plano-implementacao.md`.
