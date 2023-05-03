from flask import Blueprint, jsonify, request, send_file
import pandas as pd
import xlsxwriter
import io

api_bp = Blueprint('api', __name__)

@api_bp.route('/export', methods=['POST'])
def post_json():
    data = request.get_json()
    data_list = [data]
    df = pd.DataFrame(data_list)
    excel_file = io.BytesIO()

    # Escreve o DataFrame no arquivo Excel
    with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)

    # Retorna o arquivo Excel em memória como um anexo
    excel_file.seek(0)
    return send_file(
        excel_file,
        attachment_filename='data.xlsx',
        as_attachment=True,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

    # return send_file(output, attachment_filename='output.xlsx', as_attachment=True)