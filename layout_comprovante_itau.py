import pdfplumber
import pandas as pd
import os

def extract_text_from_coordinates(page, coordinates):
    cropped = page.crop(coordinates)
    return cropped.extract_text()


resultados = []

origem_comprovantes='C:\\Users\\Gabriel\\Desktop\\chromedriver_win32\\leituraDeComprovantes\\Boletos\\'
destino_relatorio='C:\\Users\\Gabriel\\Desktop\\chromedriver_win32\\leituraDeComprovantes\\resultados.xlsx'

# Percorra todos os arquivos no diretório
for raiz, subdiretorios, arquivos in os.walk(origem_comprovantes):
    for nome_arquivo in arquivos:
        caminhoPDF = os.path.join(raiz, nome_arquivo)

        with pdfplumber.open(caminhoPDF) as pdf:
            for page in pdf.pages:

                coordinates = {
                    "beneficiario": (72.9412538, 191.4860910000001, 239.09281080000002, 201.45319100000006), #
                    "cnpj_beneficiario": (291.5328538, 203.4466910000001, 380.7483659, 213.41379100000006), 
                    "data_vencimento": (410.8473538, 203.4466910000001,  460.7227222, 213.41379100000006),
                    "valor_boleto": (409.9365538, 225.37439100000006, 448.72850700000004, 235.34149100000002),
                    "desconto": (409.9365538,247.30209100000002,429.3325304, 257.269191),
                    "multa": (409.9365538,269.229891, 429.3325304, 279.196991),
                    "valor_pagamento": (410.8473538, 291.1575909999, 449.63930700000003, 301.124691),
                    "data_pagamento": (410.8473538, 313.085291, 460.7227222, 323.052391),
                    "boleto": (277.3414, 156.918191,  557.1976338, 166.885291)
                }

                page_data = {}

                if "SEARA" in  extract_text_from_coordinates(page, coordinates['beneficiario']):
                    
                    coordinates = {
                    "beneficiario": (72.9412538, 181.428391, 106.7297228, 191.395491), #
                    "cnpj_beneficiario": (291.5328538, 193.388891, 380.7483659, 203.355991), 
                    "data_vencimento": (410.8473538, 193.388891,  460.7227222, 203.355991),
                    "valor_boleto": (409.9365538, 215.316691, 448.72850700000004, 225.283791),
                    "desconto": (409.9365538,237.244391,429.3325304, 247.211491),
                    "multa": (409.9365538,259.172091, 429.3325304, 269.139191),
                    "valor_pagamento": (410.8473538, 281.099791, 449.63930700000003, 291.066891),
                    "data_pagamento": (410.8473538, 303.027491, 460.7227222, 312.994591),
                    "boleto": (277.3414, 156.918191,  557.1976338, 166.885291)                    
                    } 

                elif "UNIMED" in  extract_text_from_coordinates(page, coordinates['beneficiario']):
                    
                    coordinates = {
                    "beneficiario": (72.9412538, 181.428391, 112.2514962, 191.395491), #
                    "cnpj_beneficiario": (291.5328538, 193.388891, 380.7483659, 203.355991), 
                    "data_vencimento": (410.8473538, 193.388891,  460.7227222, 203.355991),
                    "valor_boleto": (409.9365538, 215.316691, 448.72850700000004, 225.283791),
                    "desconto": (409.9365538,237.244391,429.3325304, 247.211491),
                    "multa": (409.9365538,259.172091, 429.3325304, 269.139191),
                    "valor_pagamento": (410.8473538, 281.099791, 449.63930700000003, 291.066891),
                    "data_pagamento": (410.8473538, 303.027491, 460.7227222, 312.994591),
                    "boleto": (277.3414, 156.918191,  557.1976338, 166.885291)                    
                    }   

                elif "UNNIIM" in  extract_text_from_coordinates(page, coordinates['beneficiario']):
                    
                    coordinates = {
                    "beneficiario": (72.9412538, 172.367291, 112.2514962,182.334391 ), #
                    "cnpj_beneficiario": (291.5328538, 184.327891, 380.7483659, 194.294991), 
                    "data_vencimento": (410.8473538, 193.388891,  460.7227222,203.355991 ),
                    "valor_boleto": (409.9365538, 206.255591, 448.72850700000004,216.222691 ),
                    "desconto": (409.9365538,228.183391,429.3325304,238.150491 ),
                    "multa": (409.9365538,250.111091, 429.3325304, 260.078191),
                    "valor_pagamento": (410.8473538, 272.038791, 449.63930700000003, 282.005891),
                    "data_pagamento": (410.8473538, 293.966491, 460.7227222,303.933591 ),
                    "boleto": (277.3414, 156.918191,  557.1976338, 166.885291)

                    }  
                elif  "NEESSTTLL" in extract_text_from_coordinates(page, coordinates['beneficiario']):
                    coordinates = {
                    "beneficiario":  (66.0, 180.0, 209.0, 190.0), #
                    "cnpj_beneficiario": (291.5328538, 194.385591, 380.7483659, 204.352691), 
                    "data_vencimento": (410.8473538, 194.385591,  460.7227222,204.352691 ),
                    "valor_boleto": (409.9365538, 216.313391, 448.72850700000004,226.280491 ),
                    "desconto": (409.9365538,238.241091,429.3325304,248.208191 ),
                    "multa": (409.9365538,260.168791, 429.3325304, 270.135891),
                    "valor_pagamento": (410.8473538, 282.096491, 449.63930700000003, 292.063591),
                    "data_pagamento": (407.0, 300.0,468.0 , 315.0),
                    "boleto": (277.3414, 156.918191,  557.1976338, 166.885291)                    
                    }
                elif "MAASSTTE" in extract_text_from_coordinates(page, coordinates['beneficiario']):
                    coordinates = {
                    "beneficiario":  (72.9412538, 182.425091,114.4741595, 192.392191), #
                    "cnpj_beneficiario": (291.5328538, 194.385591, 380.7483659, 204.352691), 
                    "data_vencimento": (410.8473538, 194.385591,  460.7227222,204.352691 ),
                    "valor_boleto": (409.9365538, 216.313391, 448.72850700000004,226.2804911 ),
                    "desconto": (409.9365538,238.241091,429.3325304,248.208191 ),
                    "multa": (409.9365538,260.168791, 429.3325304, 270.135891),
                    "valor_pagamento": (410.8473538, 282.096491, 449.63930700000003, 292.063591),
                    "data_pagamento": (410.8473538, 304.024291 , 460.7227222,313.991391),
                    "boleto": (277.3414, 156.918191, 557.1976338, 166.885291)                    
                    }      

                for key, coord in coordinates.items():
                    text = extract_text_from_coordinates(page, coord)
                    #print(text)
                    page_data[key] = text



                resultados.append(page_data)


df = pd.DataFrame(resultados)
df.to_excel(destino_relatorio, index=False, engine='openpyxl')

