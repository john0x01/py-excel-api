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
        return {'message': 'Erro na requisição JSON: Informe "data": [], "title": string e "currencyFormat": [] '}, 400

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
                try: 
                    sheet_name = response[counter]["title"] if response[counter]["title"] else 'Dados ' + str(counter + 1)
                except:
                    sheet_name = 'Dados'+ str(counter + 1)
                
                df.to_excel(writer, index=False, sheet_name=sheet_name)
                workbook = writer.book
                worksheet = writer.sheets[sheet_name]

                currency_format = workbook.add_format({"num_format": "R$0.00"})  
                index = 0
                for column in response[counter]["data"][0]:
                    print(f"Analisando coluna {column}")
                    if(column in response[counter]["currencyFormat"]):
                        worksheet.set_column(index, index, 24, currency_format)
                    index += 1
                # print(f"Formatos {response[counter]['currencyFormat']}") 
                # for row_to_format in response[counter]["currencyFormat"]:
                #     print(row_to_format)
                #     index = 0     
                #     for key in response[counter]["data"][0]:
                #         print(f"Analisando coluna {key}")
                #         if(key == row_to_format):
                #             print(f"É igual a {row_to_format}")
                #             worksheet.set_column(index, index, 24, currency_format)
                #         else:
                #             print(f"Não é igual a {row_to_format}")
                #             worksheet.set_column(index, index, 24)
                #         index += 1
                counter = counter + 1
                        

        # Retorna o arquivo Excel em memória como um anexo
        excel_file.seek(0)
        
        return send_file(
            excel_file,
            attachment_filename='data.xlsx',
            as_attachment=True,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except:
        return {
            'message': 'Erro na requisição JSON: Informe "data[]" e "currencyFormat[]" '

            }, 400