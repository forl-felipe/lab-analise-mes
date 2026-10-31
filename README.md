# lab-analise-mes
Análise de dados do laboratório.

## Projeto: Preditor de TML (Transportable Moisture Limit)

Relatório completo: [`reports/RELATORIO_TML.md`](reports/RELATORIO_TML.md)

```
data/raw/            planilhas originais (não versionadas – contêm nomes de clientes/navios)
data/processed/      base consolidada + auditoria de qualidade (geradas pelo ETL)
src/etl.py           consolida os históricos 2025/2026 e audita inconsistências
src/tml_physics.py   física do ensaio PFD80 (índice de vazios <-> TML)
src/predictor.py     preditor híbrido v0 (treino, previsão, limite inferior 95 %)
src/analysis.py      raio-X da base, benchmark de modelos e figuras
models/              coeficientes treinados (JSON)
reports/             relatório, métricas e figuras
```

### Reproduzir
```bash
pip install -r requirements.txt
# coloque as planilhas em data/raw/Historico_TML_2025.xlsx e Historico_TML_2026.xlsx
python src/etl.py
python src/predictor.py train
python src/analysis.py
python src/predictor.py predict --ppc 3.60 --rho 4.690 --sio2 2.10 --p 0.061 --embarque 1
```
