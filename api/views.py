from flask import Blueprint, request
import json
from api.controllers.fileController import export_excel

api_bp = Blueprint('api', __name__)

@api_bp.route('/exportDefault', methods=['POST'])
def post_json():
    response = request.get_json()

    return export_excel(response)
