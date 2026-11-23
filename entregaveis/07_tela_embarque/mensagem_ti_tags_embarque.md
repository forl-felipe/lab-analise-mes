**Assunto:** Tags de resultados do embarque sem valor no PIMS (UBU)

Prezados,

O Laboratório de Controle de Embarque (LCE) está colocando em uso uma tela no PIMS (Aspen Process Explorer) para acompanhar a qualidade do embarque. Das 15 tags de resultado do embarque, apenas 4 recebem valores do MES. As outras 11 existem no servidor UBU, mas estão sem valor (`???` no Tag Browser) e sem descrição de origem.

**Tags que funcionam** (mapa `IP_MesValor`, descrição com o vínculo do MES):

| Tag | Descrição no PIMS |
|---|---|
| `M0620010000-0002-TNLQU` | U00-09TR002_PQ_Fe |
| `M0620010000-0004-TNLQU` | U00-09TR002_PQ_SiO2 |
| `M0620010000-0005-TNLQU` | U00-09TR002_PQ_Al2O3 |
| `M0620010000-0013-TNLQU` | U00-09TR002_PQ_PPC |

**Tags sem valor** (mapa `IP_AnalogMap`, valor `???`):

| Tag | Análise | Código MES |
|---|---|---|
| `M0620010000-0006-TNLQU` | CaO (%) | 620010000-6-TN-LQU |
| `M0620010000-0007-TNLQU` | MgO (%) | 620010000-7-TN-LQU |
| `M0620010000-0008-TNLQU` | P (%) | 620010000-8-TN-LQU |
| `M0620010000-0018-TNLQU` | B2 (%) | 620010000-18-TN-LQU |
| `M0620010000-0115-TNLQU` | Mn (%) | 620010000-115-TN-LQU |
| `M0620010000-0001-TNLFU` | H2O (%) | 620010000-1-TN-LFU |
| `M0620010000-0026-TNLFU` | -6,3 mm (%) | 620010000-26-TN-LFU |
| `M0620010000-0029-TNLFU` | Tamboramento +6,3 mm (%) | 620010000-29-TN-LFU |
| `M0620010000-0031-TNLFU` | Compressão 16,0 mm (kgf/pelota) | 620010000-31-TN-LFU |
| `M0620010000-0035-TNLFU` | Compressão 12,5 mm (kgf/pelota) | 620010000-35-TN-LFU |
| `M0620010000-2021-TNLFU` | Relação granulométrica | 620010000-2021-TN-LFU |

**Solicitação**

1. Configurar as 11 tags acima para receber os resultados do MES da mesma forma que as 4 que funcionam (mapa `IP_MesValor`, vínculo com a análise correspondente do embarque), com histórico.
2. Confirmar se, depois da configuração, as tags manterão o mapa `IP_AnalogMap` ou passarão para `IP_MesValor`. A tela foi preparada com o mapa atual de cada tag; se o mapa mudar, ajustamos.
3. Verificar o histórico das tags de Fe, SiO2, Al2O3 e PPC: o valor atual aparece, mas a tendência dos últimos dias está vazia.

Fico à disposição para validar os primeiros resultados.

Atenciosamente,

Felipe
Laboratório de Controle de Embarque (LCE), UBU
