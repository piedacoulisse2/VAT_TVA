import http.client
import xlwings as xw
import zeep
import requests
import numpy as np

@xw.func(async_mode='threading')
def tva(code_pays, code_tva):
    wsdl = 'http://ec.europa.eu/taxation_customs/vies/checkVatService.wsdl'
    client = zeep.Client(wsdl=wsdl)
    r = client.service.checkVat(code_pays, code_tva)
    return r

@xw.func(async_mode='threading')
def main():
    wb = xw.Book.caller()
    sht = wb.sheets[0]
    nbr_elts = sht.range(1, 1).options(numbers=int).value
    for i in range(nbr_elts):
        a = str(sht.range(3 + i, 2).value)
        b = str(sht.range(3 + i, 3).value)
        a = a.replace(" ", "")
        b = b.replace(" ", "")
        re = tva(a, b)
        sht.range(3 + i, 4).value = re['valid']
        sht.range(3 + i, 5).value = re['name']
        sht.range(3 + i, 6).value = re['address']
        sht.range(3 + i, 7).value = re['requestDate']

@xw.func(async_mode='threading')
def tva_valid(code_pays , code_tva ):
    code_pays= code_pays.replace(" ", "")
    code_tva=code_tva.replace(" ", "")
    if code_tva == 'NA':
        code_tva = code_pays[2:len(code_pays)]
        code_pays = code_pays[0:2]
    try :
        re = tva(code_pays, code_tva)
        r = re['valid']
    except:
        r = "Erreur"
    return r

@xw.func(async_mode='threading')
def tva_name(code_pays, code_tva):
    code_pays = code_pays.replace(" ", "")
    code_tva = code_tva.replace(" ", "")
    if code_tva == 'NA':
        code_tva = code_pays[2:len(code_pays)]
        code_pays = code_pays[0:2]
    try:
        re = tva(code_pays, code_tva)
        r = re['name']
    except:
        r = 'Erreur'
    return r

@xw.func(async_mode='threading')
def tva_address(code_pays, code_tva):
    code_pays = code_pays.replace(" ", "")
    code_tva = code_tva.replace(" ", "")
    if code_tva == 'NA':
        code_tva = code_pays[2:len(code_pays)]
        code_pays = code_pays[0:2]
    try :
        re = tva(code_pays, code_tva)
        r = re['address']
    except:
        r = 'Erreur'
    return r

@xw.func(async_mode='threading')
def tva_requestDate(code_pays, code_tva):
    code_pays = code_pays.replace(" ", "")
    code_tva = code_tva.replace(" ", "")
    if code_tva == 'NA':
        code_tva = code_pays[2:len(code_pays)]
        code_pays = code_pays[0:2]
    try:
        re = tva(code_pays, code_tva)
        r = re['requestDate']
    except:
        r = 'Erreur'
    return r

@xw.func()
def NumeroTVA(SIREN):
    SIREN = int(SIREN)
    SIREN_STR = str(SIREN)
    CLE = (12 + 3 * (np.mod(SIREN, 97)))
    CLE = np.mod(CLE, 97)
    CLE_STR = str(CLE)
    return "FR" + CLE_STR + SIREN_STR

@xw.ret(expand='table')
def SirenINSEE():
    wb = xw.Book.caller()
    sht = wb.sheets[3]
    Access_Token =  sht.range("Access_Token").value
    #Access_Token = '9edacf79-703e-380c-a10a-8397dcfe66d2'
    try:
        NBLIGNE = str(sht.range('NBLIGNE').options(numbers=int).value+ 9)
    except:
        NBLIGNE = "9"
    plage = 'A9:E'+NBLIGNE
    sht.range(plage).value = ""

    Noms = sht.range("Recherche_Entreprise").value
    Noms = str(Noms)
    url = "https://api.insee.fr/entreprises/sirene/V3/siren?q=periode(denominationUniteLegale%3A"+Noms+")&champs=siren%2CdenominationUniteLegale&masquerValeursNulles=true&debut=0"

    payload = {}
    headers = {
      'Authorization': 'Bearer ' + Access_Token,
      'Cookie': 'pdapimgateway=1830169354.22560.0000; INSEE=1627925258.20480.0000'
    }

    response = requests.request("GET", url, headers=headers, data = payload)

    status_reponse = response.status_code
    if status_reponse != 200:
        if status_reponse == 401:
            msg = "Erreur d'authentification à l'API. Vérifier la clé de l'API INSEE."
            sht.range('A10').value = msg
        sht.range('A9').value = "Erreur API INSEE"
        exit()
    datastore = response.json()
    sht.range('A9').value = "Recherche en cours"
    iterations = datastore['header']['nombre']

    r = []
    for i in range(iterations):
        SIRENE = datastore['unitesLegales'][i]['siren']
        try:
            NumTVA = NumeroTVA(SIRENE)
            NumTVA = str(NumTVA)

        except(RuntimeError, TypeError, NameError,ValueError,OSError):
            NumTVA = "Erreur lors du calcul"
        try:
            code_tva = NumTVA[2:len(NumTVA)]
            code_pays = NumTVA[0:2]
            re = tva(code_pays, code_tva)
            ValidTVA = re['valid']
            AdresseTVA = re['address']

        except(RuntimeError, TypeError, NameError, ValueError,OSError):
            ValidTVA ="Erreur connexion VIES"
            AdresseTVA = ""

        r.append([datastore['unitesLegales'][i]['periodesUniteLegale'][0]['denominationUniteLegale'],
                  SIRENE,NumTVA,ValidTVA,AdresseTVA])

    sht.range('A9').value = r
    #return r


def getToken():
    conn = http.client.HTTPSConnection("")

    payload = "grant_type=client_credentials&client_id=%24%7Baccount.clientId%7D&client_secret=T86WxnGfYNVu1Xf0VHDEHBvheTsa&audience=IgqHxRNY6Ic7Zk05esSvuB1rW2ga"

    headers = { 'content-type': "application/x-www-form-urlencoded" }

    conn.request("POST", "https://api.insee.fr/token", payload, headers)

    res = conn.getresponse()
    data = res.read()

    print(data.decode("utf-8"))



