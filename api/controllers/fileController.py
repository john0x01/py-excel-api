from flask import send_file
import pandas as pd
import xlsxwriter
import io


def export_excel(response):
    try:
        df = pd.DataFrame(data=response["data"])
        
        excel_file = io.BytesIO()

        # Escreve o DataFrame no arquivo Excel
        with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Relatório')
            workbook = writer.book
            worksheet = writer.sheets["Relatório"]

            currency_format = workbook.add_format({"num_format": "R$#.##0"})   
            index = 0     
            for row_to_format in response["currencyFormat"]:
                for key in response["data"][0]:
                    if(key == row_to_format):
                        worksheet.set_column(index, index, 24, currency_format)
                    else:
                        worksheet.set_column(index, index, 24)
                    index += 1
                        

        # Retorna o arquivo Excel em memória como um anexo
        excel_file.seek(0)
        
        return send_file(
            excel_file,
            attachment_filename='data.xlsx',
            as_attachment=True,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except:
        return {'message': 'Erro na requisição JSON: Informe "data[]" e "currencyFormat[]" '}, 400
