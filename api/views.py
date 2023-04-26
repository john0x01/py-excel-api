from flask import Blueprint, jsonify, request

api_bp = Blueprint('api', __name__)

@api_bp.route('/export', methods=['POST'])
def post_json():
    data = request.get_json()
    return jsonify(data) 