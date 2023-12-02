import PySimpleGUI as sg
# import PySimpleGUIQt as sg
import sys
import os.path
import OpenSSL.crypto
import os
import requests
import datetime
import configparser
import xlsxwriter
import tempfile
import hashlib

class dotdict(dict):
   """dot.notation access to dictionary attributes"""
   __getattr__ = dict.get
   __setattr__ = dict.__setitem__
   __delattr__ = dict.__delitem__

CACHE_KEY = None
CACHE_ESTADO_LISTA = {"CODIGO":[], "UF":[]} 
CACHE_URF_LISTA = {"CODIGO":[], "DESCRICAO":[]} 

# determine if application is a script file or frozen exe
if getattr(sys, 'frozen', False):
    application_path = os.path.dirname(sys.executable)
elif __file__:
    application_path = os.path.dirname(__file__)

CHAVE_NFE_LEN = 44
CURRENT_DIR = application_path
CERTIFICATE_DIR = os.path.join(CURRENT_DIR, "certs")
CONFIG_FILE = os.path.join(CURRENT_DIR, "config.ini")
    
PORTAL_UNICO_HOST = "https://portalunico.siscomex.gov.br"
PORTAL_UNICO_URL_AUTENTICAR = "/portal/api/autenticar"
PORTAL_UNICO_URL_CONSULTA = "/due/api/ext/due"
# "/cct/api/ext/deposito-carga/estoque-nota-fiscal/"

def getCurrentDir():
   return CURRENT_DIR

def getCertificateDir():
   if not os.path.exists(CERTIFICATE_DIR):
      os.makedirs(CERTIFICATE_DIR)
   return CERTIFICATE_DIR

def getCeritificatePEMFile():
   return os.path.join(getCertificateDir(), "personal.pem")

def getCeritificateKEYFile():
   return os.path.join(getCertificateDir(), "personal.key")

"""
   Salvar configurações do App
"""
def saveConfig(file_pfx, file_pem, file_key):
   # config = configparser.ConfigParser()
   config = loadConfig()
   config['PFX'] = {}
   config['PFX']["file"] = file_pfx
   config['PEM'] = {}
   config['PEM']["file"] = file_pem
   config['PEM']["file_key"] = file_key
   with open(CONFIG_FILE, 'w') as configfile:
      config.write(configfile)

"""
   Carregar configurações do App
"""
def loadConfig():
   config = configparser.ConfigParser()
   config.read(CONFIG_FILE)   
   return config

"""
   Converter arquivo PFX para PEM
"""
def pfx_to_pem(pfx_path, pfx_password, pem_file, key_file, error_message=None):
   def addErr(msg):
      if not error_message is None:
         error_message(msg)

   if not os.path.isfile(pfx_path):
      return False
   if os.path.isfile(pem_file):
      os.remove(pem_file)
   if os.path.isfile(key_file):
      os.remove(key_file)   

   # gravar arquivo .pem
   with open(pem_file, 'wb') as f_pem:
      pfx = None
      with open(pfx_path, 'rb') as f:
         pfx = f.read()         

      try:
         p12 = OpenSSL.crypto.load_pkcs12(pfx, pfx_password)         
      except Exception as e:
         msg_err = str(e)
         if "mac verify failure" in msg_err:
            addErr("senha inválida.")
         else:
            addErr(str(e))
         return False

      if p12.get_certificate().has_expired():
         time_string = p12.get_certificate().get_notAfter().decode("utf-8")
         addErr("Ceritificado expirou em "+str(datetime.datetime.strptime(time_string, "%Y%m%d%H%M%SZ")))
         return False

      # gravar arquivo .key
      with open(key_file, 'wb') as f_key:
         f_key.write(OpenSSL.crypto.dump_privatekey(OpenSSL.crypto.FILETYPE_PEM, p12.get_privatekey()))

      f_pem.write(OpenSSL.crypto.dump_certificate(OpenSSL.crypto.FILETYPE_PEM, p12.get_certificate()))
      ca = p12.get_ca_certificates()
      if ca is not None:
         for cert in ca:
            f_pem.write(OpenSSL.crypto.dump_certificate(OpenSSL.crypto.FILETYPE_PEM, cert))

   return (os.path.isfile(pem_file) and os.path.isfile(key_file))

"""
   Tela para Edição de Configurações
"""
def Config():
   cfg = loadConfig()
   arquivo_pfx = ""
   if "PFX" in cfg and "file" in cfg["PFX"]:
      arquivo_pfx = cfg["PFX"]["file"]

   layout = [  [sg.Text('Certificado Digital:')],
               [sg.Text('Arquivo: ', size=(15, 1)), sg.InputText(key='-file1-', default_text=arquivo_pfx), sg.FileBrowse()],
               [sg.Button('Confirmar'),sg.Button("Sair")]  ]

   window = sg.Window('Configurações', layout, size=(600,100))
   while True:
      event, values = window.read()       # type: (str, dict)
      if event in (sg.WIN_CLOSED, 'Exit'):
         break
      if event == 'Confirmar':
         pfx_path = values["-file1-"]
         if not os.path.isfile(pfx_path):
            sg.popup(f"Arquivo ${pfx_path} não encontrado.")
         else:
            pfx_password = sg.popup_get_text('Informe a senha do arquivo: ', title="Textbox")
            errMsg = []
            errMsgFunc = lambda x: errMsg.append(x)
            if pfx_to_pem(pfx_path, pfx_password, getCeritificatePEMFile(), getCeritificateKEYFile(), errMsgFunc):
               saveConfig(pfx_path, getCeritificatePEMFile(), getCeritificateKEYFile())
               sg.popup("Arquivo configurado.")
            else:
               sg.popup("Erro no certificado informado: "+"\n".join(errMsg))
               continue
         break
      if event == 'Sair':
         break 

   window.close()

"""
   getLicenseKey
"""
def getLicenseKey():
   cfg = loadConfig()
   if "LicenseKey" in cfg and "key" in cfg["LicenseKey"]:
      global CACHE_KEY
      CACHE_KEY = cfg["LicenseKey"]["key"]

"""
   hasLicenseForNFE
"""
def hasLicenseForNFE(chaveNfe):
   if CACHE_KEY is None:
      getLicenseKey()
      if CACHE_KEY is None:
         return False 

   password = chaveNfe[6:14]
   salt = "ofi-consulta-due"
   dataBase_password = password+salt+"."+salt[::-1]+password[::-1]
   hashed = hashlib.md5(dataBase_password.encode())
   if hashed.hexdigest() != CACHE_KEY:
      return False 
   
   return True 

"""
   autenticarPU
"""
def autenticarPU(aHeaderRet={}):
   headers = {
      "Content-Type": "application/json",
      "Role-Type": "IMPEXP",
      "User-Agent": "Consulta-PU",
      "Accept": "*/*",
      "Cache-Control": "no-cache",
      "Accept-Encoding": "gzip, deflate",
      "Connection": "keep-alive"
   }

   url = PORTAL_UNICO_HOST+PORTAL_UNICO_URL_AUTENTICAR
   certs = (getCeritificatePEMFile(),getCeritificateKEYFile())

   result = requests.get(
      url, 
      headers=headers, 
      cert=certs
   )

   # print(result)

   if not result.ok:
      return False

   cSet_Token = result.headers["Set-Token"]
   cCRSF_TOKEN = result.headers["X-CSRF-Token"]

   aHeaderRet["Content-Type"] = "application/json"
   aHeaderRet["Role-Type"] =  "IMPEXP"
   aHeaderRet["X-CSRF-Token"] =cCRSF_TOKEN
   aHeaderRet["Authorization"] = cSet_Token
   aHeaderRet["User-Agent"] = "Consulta-PU"
   aHeaderRet["Accept"] = "*/*"
   aHeaderRet["Cache-Control"] = "no-cache"
   aHeaderRet["Accept-Encoding"] = "gzip, deflate"
   aHeaderRet["Connection"] = "keep-alive"
   aHeaderRet["cache-control"] = "no-cache"
   
   return True

def getEstado(codigo):
   if len(CACHE_ESTADO_LISTA["CODIGO"]) == 0:
      with open(os.path.join(getCurrentDir(), "data","estados.csv"),"r") as f:
         buffer = f.read()
         rows = buffer.split("\n")
         rows = rows[1:]
         for line in rows:
            fields = line.split(";")
            if len(fields) >= 2:
               CACHE_ESTADO_LISTA["CODIGO"].append(fields[0])
               CACHE_ESTADO_LISTA["UF"].append(fields[1])
         
   index = CACHE_ESTADO_LISTA["CODIGO"].index(codigo)
   if index >= 0:
      return CACHE_ESTADO_LISTA["UF"][index]
   
   return codigo
      
def getUrf(codigo):
   if len(CACHE_URF_LISTA["CODIGO"]) == 0:
      with open(os.path.join(getCurrentDir(), "data","urfs.csv"),"r") as f:
         buffer = f.read()
         rows = buffer.split("\n")
         rows = rows[1:]
         for line in rows:
            fields = line.split(";")
            if len(fields) >= 2:
               CACHE_URF_LISTA["CODIGO"].append(fields[0])
               CACHE_URF_LISTA["DESCRICAO"].append(fields[1])
         
   index = CACHE_URF_LISTA["CODIGO"].index(codigo)
   if index >= 0:
      return codigo+" "+CACHE_URF_LISTA["DESCRICAO"][index]
   
   return codigo

"""
   consultaDUE
"""
def consultaDUE(headers, chave_nfe, resultDUE):
   # print(headers, chave_nfe, resultDUE)
   def funcRetornoErr(result):
      resultDUE["statusDUE"] = ""
      resultDUE["numeroDUE"] = ""
      resultDUE["numeroRUC"] = ""
      resultDUE["dataRegistro"] = ""
      resultDUE["dataAverbacao"] = ""
      resultDUE["msgErr"] = f"error code: {result.status_code} ({result.reason})"
      return False

   resultDUE["chave_nfe"] = chave_nfe
    
   # certs = (getCeritificatePEMFile(),getCeritificateKEYFile())
   # requestParams={"numeroNFE": chave_nfe}

   url = PORTAL_UNICO_HOST+PORTAL_UNICO_URL_CONSULTA 
   # url = url
   requestParams={"nota-fiscal": chave_nfe,"Authorization": headers["Authorization"], "X-CSRF-Token": headers["X-CSRF-Token"]}
   # del headers["Authorization"]
   # del headers["X-CSRF-Token"]

   hasLicense = hasLicenseForNFE(chave_nfe)
   result = None
   if hasLicense:
      result = requests.get(
         url, 
         params=requestParams,
         headers=headers, 
      )
   else:
      result = {"ok": False, "status_code": 99999999, "reason": "CNPJ ? não está licenciado para o uso do App.".replace("?",chave_nfe[6:20])}
      result = dotdict(result)
   # cert=certs

   resultDUE["numero"] = chave_nfe
   resultDUE["numeroNF"] = chave_nfe[25:34]
   resultDUE["estadoNF"] = getEstado(chave_nfe[:2])

   if not result.ok or result.status_code == 204:
      return funcRetornoErr(result)

   json = result.json()

   if len(json) == 0 or  not "href" in json[0].keys():
      result = dotdict({"status_code": "404", "reason": "Nota Fiscal não encontrada."})
      return funcRetornoErr(result)
   
   url = json[0]["href"]
   requestParams={}
   result = requests.get(
      url, 
      params=requestParams,
      headers=headers, 
   )
   if not result.ok:
      return funcRetornoErr(result)

   json = result.json()

   resultDUE["statusDUE"] = json["situacao"]
   resultDUE["numeroDUE"] = json["numero"]
   resultDUE["numeroRUC"] = json["ruc"]
   resultDUE["dataRegistro"] = json["dataDeRegistro"]

   dataAverbacao = ""
   eventoAverbacao = [x for x in json["eventosDoHistorico"] if x["evento"] == "Averbação"]
   if len(eventoAverbacao) > 0:
      dataAverbacao = eventoAverbacao[0]["dataEHoraDoEvento"]
   resultDUE["dataAverbacao"] = dataAverbacao
   resultDUE["msgErr"] = f"error code: {result.status_code} ({result.reason})"

   if resultDUE["dataRegistro"] != "":
      resultDUE["dataRegistro"] = datetime.datetime.strptime(resultDUE["dataRegistro"][:10], '%Y-%m-%d').date()

   if resultDUE["dataAverbacao"] != "":
      resultDUE["dataAverbacao"] = datetime.datetime.strptime(resultDUE["dataAverbacao"][:10], '%Y-%m-%d').date()

   return True

"""
   exportToExcelFile
"""
def exportToExcelFile(list):

   refDate = datetime.datetime.now()
   dest_filename = os.path.join(
      tempfile.gettempdir(),
      f"resultado-pu-{refDate.strftime('%Y-%m-%dT%H%M%S')}.xlsx"
   )

   column_header = [
      "numero",
      "numeroNF",
      "estadoNF",
      "statusConsulta",
      "statusDUE",
      "numeroDUE",
      "numeroRUC",
      "dataRegistro",
      "dataAverbacao",
      "msgErr"
   ]
   
   data = []
   for item in list:
      row = [""]*len(column_header)

      for index, column_name in enumerate(column_header):
         if column_name in item:
            row[index] = item[column_name]
      
      data.append(row)

   workbook = xlsxwriter.Workbook(dest_filename)
   format_date = workbook.add_format({'num_format': 'dd/mm/yy'})
   worksheet1 = workbook.add_worksheet()
   # Add a table to the worksheet.
   columns =  [({"header": item} if item != "dataRegistro" and item != "dataAverbacao" else {"header": item, "format": format_date}) for item in column_header]
   worksheet1.add_table(
      0, 0, len(data), len(column_header)-1,
      {
         "data": data,
         "columns": columns
      },
   )
   workbook.close()

   return dest_filename

"""
   processarConsulta
"""
def processarConsulta(listaChaveNFe):
   result = True

   MLINE_KEY = '-MLINE-'
   BTN_ACAO = "-BTN-"

   layout = [  [sg.Button("Iniciar", key=BTN_ACAO)],
               [sg.Text('Consultando chaves de NF-e:')],
               [sg.ProgressBar(len(listaChaveNFe), orientation='h', expand_x=True, size=(20, 20),  key='-PBAR-')],
               [sg.Text('', key='-OUT-', enable_events=True, font=('Arial Bold', 16), justification='center', expand_x=True)],
               [sg.Multiline(size=(140,50), key=MLINE_KEY, reroute_cprint=True, write_only=True, disabled=True)]]

   window = sg.Window('Consulta Portal Único Exportação', layout, size=(800,600))

   signals = {}
   signals["userCancelled"] = False
   signals["btnRunning"] = False
   signals["firstLoop"] = True
   signals["btnSair"] = False

   sg.cprint_set_output_destination(window, MLINE_KEY)

   def handleEvents(event, values, signals):
      if event in (sg.WIN_CLOSED, 'Exit') or  event == 'Cancelar':
         signals["userCancelled"] = True
         return False
      
      if event == BTN_ACAO:
         if signals["btnSair"]:
            return False
         
         if signals["btnRunning"]:
            signals["btnRunning"] = False
            window[BTN_ACAO].update(text="Iniciar")
         else:
            signals["btnRunning"] = True
            window[BTN_ACAO].update(text="Cancelar")

      return True

   while not signals["userCancelled"]:      
      event, values = window.read(timeout=100)       # type: (str, dict)      
      if not handleEvents(event, values, signals):
         break 

      if signals["firstLoop"]:
         signals["firstLoop"] = False 
         window[BTN_ACAO].click()
         continue

      if signals["btnRunning"]:
         window['-PBAR-'].update(current_count=0)
         headers={}
         if not autenticarPU(headers):
            sg.cprint("Erro para obter token de autenticação, certifique-se que os certificados estão válidos.")
         else: 
            results = []
            total = len(listaChaveNFe)
            for i in range(total):
               event, values = window.read(timeout=100)       # type: (str, dict)      
               chave_nfe = listaChaveNFe[i]
               sg.cprint(f"Buscando chave {chave_nfe}")
               resultDUE = {}
               if consultaDUE(headers, chave_nfe, resultDUE):
                  resultDUE["statusConsulta"] = True
               else:
                  resultDUE["statusConsulta"] = False
               results.append(resultDUE)
               if not handleEvents(event, values, signals):
                  break
               if not signals["btnRunning"]:
                  break
               window['-PBAR-'].update(current_count=i+1)
               window['-OUT-'].update(f"{i+1}/{total}")

            if not signals["btnRunning"]:
               sg.cprint("Processamento cancelado pelo usuário.")
            else:
               fileXlsx = exportToExcelFile(results)
               import webbrowser
               if not webbrowser.open('file:///' + fileXlsx):
                  webbrowser.open('file:///' +os.path.dirname(fileXlsx))
               sg.cprint("Resultado gerado em: ", fileXlsx)
               sg.cprint("Fim de processamento.")

         window.read(timeout=100)
         if not signals["btnRunning"]:
            result = False
            sg.popup("Processamento cancelado pelo usuário.")
         else:
            result = True
            # sg.popup("Processamento concluído.")
         signals["btnRunning"] = False
         # window.close()
         signals["btnSair"] = True
         window[BTN_ACAO].update(text="Fechar")


   window.close()

   return result

"""
   Convert Text to List
"""
def convertTextToList(memo_chaves):
   list = []
   lista_chaves = (memo_chaves.strip()+"\n").split("\n")
   for key in lista_chaves:      
      row = key.strip()
      while len(row) >= CHAVE_NFE_LEN:
         chave_nfe = row[0:44]
         if len(row) >= CHAVE_NFE_LEN:
            row = row[44:]
         if len(chave_nfe) == CHAVE_NFE_LEN:
            if not chave_nfe in list:
               list.append(chave_nfe)

   # print(len(list))
   return list

"""
   Valida formato das chaves de NFe para pesquisa.
"""
def validChaves(lista_chaves, error_list=[""]):

   if len(lista_chaves) == 0:
      error_list[0] = "Nenhuma chave de nfe com 44 digitos foi encontrada."
      return False
   
   return True

"""
   Aplicativo para consulta do Portal Unico CCT/Exportação.
"""
def main():
   MLINE_KEY = '-MLINE-'

   sg.theme("gray gray gray")

   layout = [  [sg.Button('Consultar'),sg.Button("Limpar"),sg.Button('Configurações'),sg.Button("Sair")],
               [sg.Text('Relação de chaves de NF-e:')],
               [sg.Multiline(size=(140,50), key=MLINE_KEY, reroute_cprint=True, write_only=True)]]

   window = sg.Window('Consulta Portal Único Exportação', layout, size=(800,600))
   while True:
      event, values = window.read()       # type: (str, dict)

      if event in (sg.WIN_CLOSED, 'Exit'):
         break

      if event == "Limpar":
         window[MLINE_KEY].update(value="")

      if event == 'Consultar':
         content = window[MLINE_KEY].get()
         error = [""]
         list = convertTextToList(content)         
         if validChaves(list, error):
            window.hide()
            processarConsulta(list)
            window.un_hide()
         else:
            sg.popup(error[0])

      if event == 'Configurações':
         window.hide()
         Config()
         window.un_hide()
      if event == 'Sair':
         break 

   window.close()

if __name__ == "__main__":
   main()
