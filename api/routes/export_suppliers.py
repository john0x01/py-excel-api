from flask import jsonify, send_file
import pandas as pd
import xlsxwriter
import io

def export_suppliers(response):
    
    try:
        data = response["data"]
        df = pd.DataFrame(data=data)
        df.rename(columns=response["renameCols"], inplace=True)
        
        excel_file = io.BytesIO()

        with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
            startrow = 1
            sheet_name = 'Relatório Cotação itens tela'
            df.to_excel(writer, index=False, sheet_name=sheet_name, startrow=startrow)
            workbook = writer.book
            worksheet = writer.sheets[sheet_name]

            total_format = workbook.add_format({
                "num_format": "R$0.00",
                "bg_color": "#171717",
                "bold": True,
                'font_color': "#FFFFFF"
            })
            currency_format = workbook.add_format({"num_format": "R$0.00"})
            title_format_odd = workbook.add_format({'bg_color': '#538DD6'})
            title_format_odd.set_align('center')
            title_format_odd.set_align('vcenter')
            title_format_even = workbook.add_format({'bg_color': '#C5D9F1'})
            title_format_even.set_align('center')
            title_format_even.set_align('vcenter')

            lastRow = len(data) + 2
            supplierIndex = 2
            for supplier in response["suppliers"]:
                if(supplierIndex % 2 == 0):
                    worksheet.merge_range(
                        0, 
                        supplierIndex, 
                        0, 
                        supplierIndex + 2, 
                        supplier["name"], 
                        title_format_odd
                    )
                else:
                    worksheet.merge_range(
                        0, 
                        supplierIndex, 
                        0, 
                        supplierIndex + 2, 
                        supplier["name"], 
                        title_format_even
                    )
                worksheet.write(lastRow, supplierIndex + 2, supplier["total"], total_format)
                supplierIndex += 3
                
          
            columnCount = 0
            for column in response["data"][0]:
                if column in response["currencyFormat"]:
                    worksheet.set_column(columnCount, columnCount, 18, currency_format)
                elif column in response["renameCols"].keys():
                    if response["renameCols"][column] in response["currencyFormat"]:
                        worksheet.set_column(columnCount, columnCount, 18, currency_format)
                    else:
                        worksheet.set_column(columnCount, columnCount, 18)
                else:
                    worksheet.set_column(columnCount, columnCount, 18)
                columnCount += 1
            
            


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