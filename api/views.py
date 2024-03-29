from flask import Blueprint, request
import json
from api.routes.export_excel import export_excel
from api.routes.export_compositions import export_compositions
from api.routes.export_suppliers import export_suppliers
from api.routes.export_tabs import export_tabs
from api.routes.export_with_children import export_with_children

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

@api_bp.route('/exportWithChildren', methods=['POST'])
def post_with_children(): 
    try:
        response = request.get_json()
    except:
        return {'message': 'Falha na requisição JSON'}, 420
    
    return export_with_children(response)

@api_bp.route('/exportSuppliers', methods=['POST'])
def post_suppliers():
    try:
        response = request.get_json()
    except:
        return { 'message': 'Falha na requisição JSON '}, 400
    
    return export_suppliers(response)