# py-excel-api
JSON to excel API made with python

## Installation
```
pip3 install -r requirements.txt
```

## Usage
```
python3 run.py
```
The API will be running at `http://localhost:5000`

## Routes
* `exportDefault`: `POST` - Receives a JSON object and return a .xlsx file
```json
{
    "currencyFormat": ["receita", "despesas"],
    "title": "Relatório Semanal",
    "data": [
        {
            "data": "08/05/2023",
            "estoque": 221,
            "receita":  823,
            "despesas": 699
        },
        {
            "data": "10/05/2023",
            "estoque": 152,
            "receita":  1020,
            "despesas": 820
        },
        {
            "data": "12/05/2023",
            "estoque": 261,
            "receita":  793,
            "despesas": 930
        },
        {
            "data": "14/05/2023",
            "estoque": 201,
            "receita":  1003,
            "despesas": 769
        },
    ]
}
```
* `exportTabs`: `POST` - Receives a JSON array of objects and return a .xlsx file with separated sheets (tabs)
```json
[
    {
        "currencyFormat": ["receita", "despesas"],
        "title": "Relatório Semanal",
        "data": [
            {
                "data": "08/05/2023",
                "estoque": 221,
                "receita":  823,
                "despesas": 699
            },
            {
                "data": "10/05/2023",
                "estoque": 152,
                "receita":  1020,
                "despesas": 820
            },
            {
                "data": "12/05/2023",
                "estoque": 261,
                "receita":  793,
                "despesas": 930
            },
            {
                "data": "14/05/2023",
                "estoque": 201,
                "receita":  1003,
                "despesas": 769
            }
        ]
    },
    {
        "currencyFormat": ["valor"],
        "title": "Movimentações",
        "data": [
            {
                "data": "23/04/2023",
                "tipo": "DINHEIRO",
                "lancamentos": "1379 - FULANO DE TAL",
                "valor": 838196.12
            },
            {
                "data": "23/04/2023",
                "tipo": "DINHEIRO",
                "lancamentos": "1376 - FULANO DE TAL",
                "valor": 31008.92
            },
            {
                "data": "23/04/2023",
                "tipo": "DINHEIRO",
                "lancamentos": "1376 - FULANO DE TAL",
                "valor": 81601.03
            }
        ]
    }
]
```
