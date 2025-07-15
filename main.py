import os
import requests
from datetime import datetime
from base64 import b64decode
import pandas as pd
from zeep import Client
from zeep.transports import Transports
from zeep.plugins import HistoryPlugin

# Varáveis fixas
WSDL = "https://www1.nfe.fazenda.gov.br/NFeDistribuicaoDFe/NFeDistribuicaoDFe.asmx?wsdl"
CAMINHO_CERTS = "certificados_convertidos"
DIRETORIO_XML = "XML_RECEITA"

def formatar_nome(nome):
    return nome.replace(" ", "_").replace("-", "").replace("/", "_")

def download_nfe(cnpj, cert_path, key_path, nome_empresa, apelido):
    mes_input = input("Digite o número do mês (MM) ou deixe em branco para o mês anterior: ")
    if mes_input.strip():
        mes_usado = mes_input.zfill(2)
    else:
        data_referencia = datetime.now().replace(day=1) - timedelta(days=1)
        mes_usado = data_referencia.strftime("%m")
    pasta_destino = os.path.join(DIRETORIO_XML, nome_empresa, mes_atual, apelido)
    os.makedirs(pasta_destino, exist_ok=True)
    
    session = requests.Session()
    session.cert = (cert_path, key_path)
    transport = Transports(session=session)
    history = HistoryPlugin()
    
    client = Client(wsdl=WSDL, transport=transport, plugins=[history])
    
    try:
        print(f"\n Consultando NF-e para: {nome_empresa}")
        request_data = {
            'distDFeInt': {
                'tpAmb': 1,
                'cUFAutor': '41',
                'CNPJ': cnpj,
                'distNSU': {
                    'ultNSU': '000000000000000'
                }
            }    
        }

        response = client.service.nfeDistDFeInteresse(**request_data)
        lote = getattr(response, 'loteDistDFeInt', None)
        
        if not lote or not hasattr(lote, 'docZip'):
            print("Nenhuma NF-e encontrada.")
            return
        
        qtd_baixados = 0
        for doc in lote.docZip:
            xml_zip_base64 = doc._value_1
            chave = doc.NSU
            xml_bytes = b64decode(xml_zip_base64)
            caminho_arquivo = os.path.join(pasta_destino, f"nfe_{chave}.xml")
            with open(caminho_arquivo, "wb") as f:
                f.write(xml_bytes)
            print(f"NF-e salva: {caminho_arquivo}")
            qtd_baixados += 1
        return qtd_baixados
            
    except Exception as e:
        print(f"Erro para {nome_empresa}: {e}")
        
def executar_para_todas():
    df = pd.read_excel("config/empresas.xlsx")
    logs = []
    
    for _, row in df.iterrows():
        nome_empresa = row["NOME_EMPRESA"]
        apelido = row["APELIDO"]
        cnpj = str(row["CNPJ"]).zfill(14)
        
        nome_formatado = formatar_nome(nome_empresa)
        cert_path = os.path.join(CAMINHO_CERTS, f"cert-{nome_formatado}.pem")
        key_path = os.path.join(CAMINHO_CERTS, f"key-{nome_formatado}.pem")
        
        if not (os.path.exists(cert_path) and os.path.exists(key_path)):
            mensagem = "Certificado não encontrado"
            print(f"{mensagem} para a empresa: {nome_empresa}")
            logs.append({
                "NOME_EMPRESA": nome_empresa,
                "CNPJ": cnpj,
                "STATUS": "ERRO",
                "MENSAGEM": mensagem
            })
            continue
        
        try:
            quantidade = download_nfe(cnpj, cert_path, key_path, nome_empresa, apelido)
            logs.append({
                 "NOME_EMPRESA": nome_empresa,
                "CNPJ": cnpj,
                "STATUS": "SUCESSO",
                "MENSAGEM": f"{quantidade} notas baixadas"                
            })
        except Exception as e:
            logs.append({
                 "NOME_EMPRESA": nome_empresa,
                "CNPJ": cnpj,
                "STATUS": "ERRO",
                "MENSAGEM": str(e)
            })
    
    # Salva o log no final
    df.log = pd.DataFrame(logs)
    df.log.to_excel("status.execucao.xlsx", index=False)
    print("\n Log de execução salvo em: status_execucao.xlsx")
       
        

if __name__ == "__main__":
    executar_para_todas()
        
