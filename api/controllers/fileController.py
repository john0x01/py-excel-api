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
            sheet_name = response["title"] if response["title"] else 'Dados'
            df.to_excel(writer, index=False, sheet_name=sheet_name)
            workbook = writer.book
            worksheet = writer.sheets[sheet_name]

            currency_format = workbook.add_format({"num_format": "R$0.00"})   
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
            attachment_filename='data.xlsx',
            as_attachment=True,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        return {
            'message': 'Erro ao gravar excel',
            'error': str(e)

            }, 400

def export_tabs(response):
    try:
        data_frame_array = []
        for obj in response:
            data_frame_array.append(pd.DataFrame(data=obj['data']))
        
        excel_file = io.BytesIO()

        counter = 0
        # Escreve o DataFrame no arquivo Excel
        with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
            for df in data_frame_array:
                startrow = 0
                if response[counter]["header"]:
                    try:
                        if response[counter]["header"]["title"]:
                            has_title = True
                            startrow = 1
                    except:
                        has_title = False
                    try:
                        if response[counter]["header"]["date"]:
                            has_date = True
                            startrow = 3
                    except:
                        has_date = False

                try: 
                    sheet_name = response[counter]["title"] if response[counter]["title"] else 'Dados ' + str(counter + 1)
                except:
                    sheet_name = 'Dados'+ str(counter + 1)
                
                df.to_excel(writer, index=False, sheet_name=sheet_name, startrow=startrow)
                workbook = writer.book
                worksheet = writer.sheets[sheet_name]

                currency_format = workbook.add_format({"num_format": "R$0.00"})

                title_format = workbook.add_format({
                    'bold': True, 
                    'font_size': '16', 
                    'bg_color': '#FABF8F', 
                })
                title_format.set_align('center')
                title_format.set_align('vcenter')

                date_format = workbook.add_format({
                    'bold': True, 
                    'font_size': '14', 
                    'bg_color': '#FABF8F'
                })

                columnCount = 0
                for column in response[counter]["data"][0]:
                    if(column == "FILIAL / CENTRO CUSTO"):
                        column_width = 98
                    else:
                        column_width = 18

                    if(column in response[counter]["currencyFormat"]):
                        worksheet.set_column(columnCount, columnCount, column_width, currency_format)
                    else:
                        worksheet.set_column(columnCount, columnCount, column_width)
                    columnCount += 1

                if has_title:
                    page_title = response[counter]["header"]["title"]
                    worksheet.set_row(0, 32)
                    worksheet.write(0, 0, page_title, title_format)
                    worksheet.merge_range(0, 0, 0, columnCount - 1, page_title, title_format)

                if has_date:
                    page_date = response[counter]["header"]["date"]
                    worksheet.write(1, 0, "PERÍODO DE PROGRAMAÇÃO", date_format)
                    worksheet.write(2, 0, page_date, date_format)
                    worksheet.merge_range(1, 0, 1, 1, "PERÍODO DE PROGRAMAÇÃO", date_format)                        
                    worksheet.merge_range(2, 0, 2, 1, page_date, date_format)    
                counter = counter + 1

                        

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

def export_compositions(response):
    try:
        df = pd.DataFrame(data=response["data"])
        
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
            for row in response["data"]:
                if(row['Código'] in response["categoryFormat"]):
                    worksheet.conditional_format(row_index, 0, row_index, len(row.keys()), { 'type': 'no_blanks', 'format': category_format})
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
            attachment_filename='data.xlsx',
            as_attachment=True,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        return {
            'message': 'Erro ao gravar excel',
            'error': str(e)

            }, 400
    
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
            attachment_filename='data.xlsx',
            as_attachment=True,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        return {
            'message': 'Erro ao gravar excel',
            'error': str(e)

            }, 400