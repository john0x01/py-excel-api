from flask import jsonify, send_file
import pandas as pd
import xlsxwriter
import io

def export_compositions(request):
    try:
        df = pd.DataFrame(data=request["data"])
        
        excel_file = io.BytesIO()

        # Escreve o DataFrame no arquivo Excel
        with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
            sheet_name = "Relatório Composição"
            start_row = 0
            try:
                if request["header"]:
                    try:
                        if request["header"]["title"]:
                            has_title = True
                            start_row = 1
                    except:
                        has_title = False
                    try:
                        if request["header"]["interval"]:
                            has_date = True
                            start_row = 2
                    except:
                        has_date = False
            except:
                pass
            df.to_excel(writer, index=False, sheet_name=sheet_name, startrow=start_row)
            workbook = writer.book
            worksheet = writer.sheets[sheet_name]

            currency_format = workbook.add_format({"num_format": "R$0.00"})   
            category_format = workbook.add_format({"bg_color": "#171717", 'font_color': "#FFFFFF"})

            title_format = workbook.add_format({
                'bold': True, 
                'font_size': '16', 
            })
            title_format.set_align('center')
            title_format.set_align('vcenter')

            interval_format = workbook.add_format({
                'font_size': '12', 
            })

            row_index = start_row + 1
            for row in request["data"]:
                if(row['Código'] in request["categoryFormat"]):
                    worksheet.conditional_format(row_index, 0, row_index, len(row.keys()), { 'type': 'no_blanks', 'format': category_format})
                else:
                    worksheet.set_row(row_index, 12)
                row_index = row_index + 1

            col_index = 0
            for column in request["data"][0]:
                if(column in request["currencyFormat"]):
                    worksheet.set_column(col_index, col_index, 24, currency_format)
                else:
                    worksheet.set_column(col_index, col_index, 24)
                col_index += 1
            
            if has_title:
                page_title = request["header"]["title"]
                worksheet.set_row(0, 32)
                worksheet.write(0, 0, page_title, title_format)
                worksheet.merge_range(0, 0, 0, col_index, page_title, title_format)

            if has_date:
                page_date = request["header"]["interval"]
                worksheet.write(1, 0, "Período: " + page_date , interval_format)
                worksheet.set_row(1, 18)
                worksheet.merge_range(1, 0, 1, col_index, "Período: " + page_date, interval_format)                        


        # Retorna o arquivo Excel em memória como um anexo
        excel_file.seek(0)
        
        return send_file(
            excel_file,
            download_name="data.xlsx",
            as_attachment=True,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        return jsonify({
            "message": "Erro ao gravar Excel",
            "error": str(e),
            "status": "400"
        }, 400)
  