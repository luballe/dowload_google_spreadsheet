import os
import gspread
import pandas as pd
import os
from http.server import BaseHTTPRequestHandler, HTTPServer

class SimpleHTTPRequestHandler(BaseHTTPRequestHandler):

    gc = gspread.service_account(filename='service_account.json')
    sheet = gc.open_by_key('1pPewqrwerwerIEW4IUGSaQt0o')

    def save_file(self,worksheet_name):
        worksheet = self.sheet.worksheet(worksheet_name)
        data = worksheet.get_all_values()
        file_name =worksheet_name+".csv"
        with open(file_name, 'w') as f:
            for row in data:
                f.write(','.join(row) + '\n')

    def prepare_files(self):
        sheets = ['Empresas','Colegios','Empresas_Estrategias_txt','Alianza_Acciones_txt','Colegios_Estrategias_txt']
        nombre_excel = 'sumate_files.xlsx'
        writer = pd.ExcelWriter(nombre_excel) # Arbitrary output name
        for sheet in sheets:
            print(sheet)
            self.save_file(sheet)
            csvfilename = sheet + '.csv'
            df = pd.read_csv(csvfilename)
            if (sheet == 'Empresas'):
                df = df.drop(columns=['🔒 Row ID'])
            elif (sheet == 'Colegios'):
                df = df.drop(columns=['🔒 Row ID','id_localidad'])
            elif (sheet == 'Empresas_Estrategias_txt'):
                df = df.drop(columns=['id_Empresa','id_Estrategia'])
            elif (sheet == 'Alianza_Acciones_txt'):
                df = df.drop(columns=['id_alianza','id_accion','id_empresa','id_colegio'])
            elif (sheet == 'Colegios_Estrategias_txt'):
                df = df.drop(columns=['id_colegio','id_estrategia'])
            df.to_excel(writer,sheet_name=os.path.splitext(csvfilename)[0],index=False)
        writer.close()

    def do_GET(self):
        if self.path == '/download_sumate':
            self.prepare_files()
            filepath = 'sumate_files.xlsx'
            with open(filepath, 'rb') as file:
                file_size = os.path.getsize(filepath)
                self.send_response(200)
                self.send_header('Content-type', 'application/octet-stream')
                self.send_header('Content-Disposition', f'attachment; filename={os.path.basename(filepath)}')
                self.send_header('Content-length', str(file_size))
                self.end_headers()
                # Read and send file in chunks
                chunk_size = 1024 * 8
                while True:
                    chunk = file.read(chunk_size)
                    if not chunk:
                        break
                    self.wfile.write(chunk)
        else:
            self.send_response(404)
            self.end_headers()

if __name__ == '__main__':
    port = 8282
    server_address = ('', port)
    httpd = HTTPServer(server_address, SimpleHTTPRequestHandler)
    print('Server active on port '+ str(port) +'...')
    httpd.serve_forever()
