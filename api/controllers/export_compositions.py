def export_compositions(request_data, pd, io, xlsxlwriter, send_file):
    try:
        df = pd.DataFrame(data=request_data["data"])
        
        excel_file = io.BytesIO()

        # Escreve o DataFrame no arquivo Excel
        with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
            sheet_name = "Relatório Composição"
            df.to_excel(writer, index=False, sheet_name=sheet_name)
            workbook = writer.book
            worksheet = writer.sheets[sheet_name]

            currency_format = workbook.add_format({"num_format": "R$0.00"})   
            category_format = workbook.add_format({"bg_color": "#171717", 'font_color': "#FFFFFF"})

            row_index = 1
            for row in request_data["data"]:
                if(row['Código'] in request_data["categoryFormat"]):
                    worksheet.conditional_format(row_index, 0, row_index, len(row.keys()), { 'type': 'no_blanks', 'format': category_format})
                else:
                    worksheet.set_row(row_index, 12)
                row_index = row_index + 1

            index = 0
            for column in request_data["data"][0]:
                if(column in request_data["currencyFormat"]):
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
    except Exception as e:
        return {
            'message': 'Erro ao gravar excel',
            'error': str(e)

            }, 400