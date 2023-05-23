from flask import Blueprint, request
import json
from api.controllers.fileController import export_excel, export_tabs, export_compositions

api_bp = Blueprint('api', __name__)

@api_bp.route('/', methods=['GET'])
def get():
    return {'message': 'OK'}, 200

@api_bp.route('/exportDefault', methods=['POST'])
def post_json():
    try:
        response = request.get_json()
    except:
        return {'message': 'Falha na requisição JSON'}, 400

    return export_excel(response)

@api_bp.route('/exportTabs', methods=['POST'])
def post_tabs():
    try:
        response = request.get_json()
    except:
        return {'message': 'Falha na requisição JSON'}, 400

    return export_tabs(response)

@api_bp.route('/exportCompositions', methods=['POST'])
def post_compos():
    try:
        response = request.get_json()
    except:
        return {'message': 'Falha na requisição JSON'}, 400
    
    return export_compositions(response)