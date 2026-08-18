# Automação de Notas Fiscais em PDF

Este projeto é um script em **Python** que automatiza o processamento de notas fiscais em PDF.  
Ele extrai informações importantes como **número da nota, data de emissão, quantidade e código BIP**, classifica cada documento em **OK**, **Suspeito** ou **Erro**, e gera um relatório em **Excel** com filtros, colunas ajustadas e cores para facilitar a análise.

## 🚀 Funcionalidades
- Leitura automática de PDFs com `pdfplumber`
- Extração de dados relevantes via expressões regulares
- Classificação das notas com base em regras de validação
- Geração de planilha Excel organizada com cores para cada status
- Log de processamento para auditoria

## 📦 Requisitos
- Python 3.9+
- Bibliotecas:
  - `pdfplumber`
  - `pandas`
  - `openpyxl`

Instale com:
```bash
pip install pdfplumber pandas openpyxl
