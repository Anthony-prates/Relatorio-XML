# Relatório de Andamentos (XML → XLSX)

A partir de uma situação real, marcada por acesso limitado, surgiu esta solução simples com o objetivo de agilizar o processo de geração de relatórios.

Script simples para consolidar arquivos XML de andamentos em uma planilha XLSX ordenada e sem duplicatas.

## Requisitos

- Python 3.10+ recomendado
- Pacotes: `xmltodict`, `pandas`, `openpyxl`

Instale com:

```bash
pip install -r requirements.txt
```

## Como usar

1. Ajuste as variáveis em `backend/xml_relatorio.py`:
   - `arquivos_xml`: caminhos completos dos XMLs a processar.
   - `responsaveis`: mapeamento de códigos para nomes.
2. Execute:

```bash
python backend/xml_relatorio.py
```

3. O arquivo `backend/pasta_XLSX/andamentos.xlsx` será criado/atualizado.

## Estrutura

- `backend/xml_relatorio.py`: lógica principal.
- `backend/pasta_XML/`: coloque aqui os XMLs de entrada.
- `backend/pasta_XLSX/`: saída XLSX gerada.
