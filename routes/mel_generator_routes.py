from flask import Blueprint, request, render_template, current_app
from services.file_processor import process_file
from utils.file_utils import save_uploaded_file, remove_file

mel_generator_bp = Blueprint('mel-generator', __name__, urel_prefix='/mel-generator')

@mel_generator_bp.route('/', methods=['GET'])