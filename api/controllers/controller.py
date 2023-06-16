from flask import send_file
import pandas as pd
import xlsxwriter
import io

from export_compositions import export_compositions
from export_excel import export_excel
from export_tabs import export_tabs

class Controller():
    @staticmethod
    def export_excel(request_data):
        return export_excel(request_data, pd, io, xlsxwriter, send_file)
    
    @staticmethod
    def export_tabs(request_data):
        return export_tabs(request_data, pd, io, xlsxwriter, send_file)

    @staticmethod
    def export_compositions(request_data):
        return export_compositions(request_data, pd, io, xlsxwriter, send_file)

