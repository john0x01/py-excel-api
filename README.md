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

### Request Example
```js
async function request() {
    const response = await fetch('http://localhost:5000/export', {
        method: 'POST',
        mode: 'cors',
        body: JSON.stringify(dataBody),
        headers: {
            'Content-Type': 'application/json'
        }
    })

    return response
}

request()
    .then(response => response.blob())
    .then(blob => {
        const url = window.URL.createObjectURL(new Blob([blob]));
        const a = document.createElement('a');
        a.href = url;
        a.download = 'arquivo.xlsx';
        document.body.appendChild(a);
        a.click();
        a.remove();
    })
    .catch(error => console.error(error));
```

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
* `exportCompositions`: `POST` - Receives an object and return a .xlsx formatting category (black background)
```json
{
    "currencyFormat": [
        "FEV/23",
        "MAR/23",
        "MAI/23",
        "JUN/23",
        "Total"
    ],
    "categoryFormat": [
        "001",
        "001",
        "001",
        "001"
    ],
    "data": [
        {
            "Código": "001",
            "Composições": "CUSTOS",
            "%": "62.5%",
            "FEV/23": 1133.34,
            "MAR/23": 566.67,
            "Total": 1700.01
        },
        {
            "Código": "SN0002",
            "Composições": "FRETE",
            "%": "18.4%",
            "FEV/23": 333.33529411764704,
            "MAR/23": 166.66764705882352,
            "Total": 500.0029411764706
        },
        {
            "Código": "IA90",
            "Composições": "nova",
            "%": "44.1%",
            "FEV/23": 800.0047058823529,
            "MAR/23": 400.00235294117647,
            "Total": 1200.0070588235294
        },
        {
            "Código": "001",
            "Composições": "CUSTOS",
            "%": "2.9%",
            "FEV/23": 0,
            "MAR/23": 80,
            "Total": 80
        },
        {
            "Código": "1118",
            "Composições": "comp cc 2/7 1",
            "%": "2.9%",
            "FEV/23": 0,
            "MAR/23": 80,
            "Total": 80
        },
        {
            "Código": "001",
            "Composições": "CUSTOS",
            "%": "4.4%",
            "FEV/23": 0,
            "MAR/23": 0,
            "MAI/23": 120,
            "Total": 120
        },
        {
            "Código": "0.01",
            "Composições": "Receitas",
            "%": "4.4%",
            "FEV/23": 0,
            "MAR/23": 0,
            "MAI/23": 120,
            "Total": 120
        },
        {
            "Código": "001",
            "Composições": "CUSTOS",
            "%": "30.2%",
            "FEV/23": 0,
            "MAR/23": 0,
            "MAI/23": 0,
            "JUN/23": 820.8275862068965,
            "Total": 820.8275862068965
        },
        {
            "Código": "0.01",
            "Composições": "Receitas",
            "%": "29.8%",
            "FEV/23": 0,
            "MAR/23": 0,
            "MAI/23": 0,
            "JUN/23": 810.9090909090909,
            "Total": 810.9090909090909
        },
        {
            "Código": "1116",
            "Composições": "comp cc 2/5 2",
            "%": "0.4%",
            "FEV/23": 0,
            "MAR/23": 0,
            "MAI/23": 0,
            "JUN/23": 9.91849529780564,
            "Total": 9.91849529780564
        }
    ]
}
```
* `exportWithChildren`: `POST` - Receives an object and return a .xlsx formatting parents (black background) and children
```json
{
    "title": "Cotação Itens",
    "parentFormat": [
        "FULANO DE TAL",
        "FORNECEDOR 123"
    ],
    "data": [
        {
            "Material": "FULANO DE TAL",
            "Observação": "QTD: 2",
            "Quantidade": "TOTAL BRUTO: R$158,00",
            "Preço": "DESCONTOS/ACRESCIMOS: R$ 0,00",
            "Total": "TOTAL: R$158,00",
        },
        {
            "Material": "1 -  | OLEO ATF TIPO A",
            "Observação": "",
            "Quantidade": "4",
            "Preço": "12",
            "Total": 48,
        },
        {
            "Materialrial": "2 -  | ÓLEO 15W40",
            "Observação": "",
            "Quantidade": "5",
            "Preço": "22",
            "Total": 110,
        },
        {
            "Material": "FORNECEDOR 123",
            "Observação": "QTD: 1",
            "Quantidade": "TOTAL BRUTO: R$40,00", 
            "Preço": "DESCONTOS/ACRESCIMOS: R$ 0,00",  
            "Total": "TOTAL: R$40,00"
        },
        {
            "Material": "1 -  | OLEO ATF TIPO A",
            "Observação": "",
            "Quantidade": "4",
            "Preço": "12",
            "Total": 48
        },
        {
            "Materialrial": "2 -  | ÓLEO 15W40",
            "Observação": "",
            "Quantidade": "5",
            "Preço": "22",
            "Total": 110
        }
    ]
}
```
* `exportSuppliers`: `POST` - Receives an object and return a .xlsx formatting materials by suppliers
```json
{
    "title": "Cotação itens",
    "currencyFormat": ["Preço", "Total"],
    "suppliers": [
      {
        "name": "HIDRAU TORQUE - GHT SP",
        "total": 18843.40
      },
      {
        "name": "ENGEPEÇAS GO", 
        "total": 15490.60,
      },
      {
        "name": "GEOMAQ SP",
        "total": 31240.56
      }
    ],
    "renameCols": {
      "MARCA_2": "MARCA", 
      "Preço_2": "Preço", 
      "Total_2": "Total",
      "MARCA_3": "MARCA", 
      "Preço_3": "Preço", 
      "Total_3": "Total"
    },
    "data": [
      {
        "Material": "COROA DE GIRO",
        "qnt.": "1",
        "MARCA": "FORTRACTOR GOLD",
        "Preço": 11364.09,
        "Total": 11364.09,
        "MARCA_2": "HP",
        "Preço_2": 15364.09,
        "Total_2": 15364.09,
        "MARCA_3": "",
        "Preço_3": 0,
        "Total_3": 0
      },
      {
        "Material": "EIXO PINHÃO",
        "qnt.": "1",
        "MARCA": "IMPORTADO",
        "Preço": 21364.09,
        "Total": 21364.09,
        "MARCA_2": "ITR-ITALIA",
        "Preço_2": 14364.09,
        "Total_2": 14364.09,
        "MARCA_3": "",
        "Preço_3": 4590,
        "Total_3": 4590
      },
      {
        "Material": "RETENTOR",
        "qnt.": "2",
        "MARCA": "BLUMAQ",
        "Preço": 196.92,
        "Total": 382.08,
        "MARCA_2": "WORLD GHT",
        "Preço_2": 6364.09,
        "Total_2": 6364.09,
        "MARCA_3": "BLUMAQ",
        "Preço_3": 420,
        "Total_3": 840
      },
    ]
}
```
