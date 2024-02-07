from flask import jsonify, send_file
import pandas as pd
import xlsxwriter
import io

def export_with_children(response):
    try:
        df = pd.DataFrame(data=response["data"])
        
        excel_file = io.BytesIO()

        # Escreve o DataFrame no arquivo Excel
        with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
            # sheet_name = response["title"] if response["title"] else 'Dados'
            sheet_name = 'Dados'
            df.to_excel(writer, index=False, sheet_name=sheet_name)
            workbook = writer.book
            worksheet = writer.sheets[sheet_name]

            currency_format = workbook.add_format({"num_format": "R$0.00"})   
            parent_format = workbook.add_format({"bg_color": "#171717", 'font_color': "#FFFFFF"})

            row_index = 1
            for row in response["data"]:
                if(row['Material'] in response["parentFormat"]):
                    worksheet.conditional_format(row_index, 0, row_index, len(row.keys()), { 'type': 'no_blanks', 'format': parent_format})
                else:
                    worksheet.set_row(row_index, 12)
                row_index = row_index + 1

            index = 0
            for column in response["data"][0]:
                if(column in response["currencyFormat"]):
                    worksheet.set_column(index, index, 24, currency_format)
                else:
                    worksheet.set_column(index, index, 24)
                index += 1
            


        # Retorna o arquivo Excel em memória como um anexo
        excel_file.seek(0)
        
        return send_file(
            excel_file,
            download_name="data.xlsx",
            as_attachment=True,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        # return jsonify'messagee': 'Erro ao gravar excel','error': str(e)), 400
        return jsonify({
            "message": "Erro ao gravar Excel",
            "error": str(e),
            "status": "400"
        }, 400)

