# Automação de Notas Fiscais em PDF

Projeto desenvolvido em **Python para automatizar um processo de trabalho relacionado ao processamento e análise de notas fiscais em PDF**.

O script automatiza tarefas que anteriormente exigiam análise manual, extraindo informações importantes como **número da nota, data de emissão, quantidade e código BIP**, validando os dados e classificando cada documento como **OK**, **Suspeito** ou **Erro**.

Ao final do processamento, o sistema gera automaticamente um **relatório em Excel**, com filtros, formatação e indicadores visuais para facilitar a análise dos resultados.

## 🚀 Funcionalidades

* Leitura automática de PDFs com `pdfplumber`
* Extração de dados utilizando expressões regulares
* Validação e classificação automática das notas
* Identificação de possíveis inconsistências
* Geração automática de relatório em Excel
* Formatação da planilha para facilitar a análise
* Registro do processamento para auditoria

## 🛠️ Tecnologias

* **Python**
* **pdfplumber**
* **Pandas**
* **OpenPyXL**
* Expressões Regulares (Regex)

## 📦 Requisitos

* Python 3.9+

Instale as dependências:

```bash
pip install pdfplumber pandas openpyxl
```

## 🎯 Objetivo

O objetivo do projeto é **reduzir tarefas manuais, aumentar a padronização do processo e facilitar a identificação de inconsistências nos documentos**, utilizando automação com Python.
