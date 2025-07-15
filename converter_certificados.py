import os
import pandas as pd
import subprocess

CAMINHO_CERTIFICADOS = 'certificados'
CAMINHO_CONVERTIDO = 'certificados_convertidos'
os.makedirs(CAMINHO_CONVERTIDO, exist_ok=True)

empresas = pd.read_excel('config/empresas.xlsx')

for _, row in empresas.iterrows():
    nome_empresa = row['NOME_EMPRESA']
    nome_arquivo = None
    
    # Buscar certificado correspondente
    for arquivo in os.listdir(CAMINHO_CERTIFICADOS):
        if arquivo.startswith(nome_empresa) and arquivo.endswith(".pfx"):
            nome_arquivo = arquivo
            break
        
    if not nome_arquivo:
        print(f"Certificado não encontrado para a empresa: {nome_empresa}")
        continue

    # Extrair a senha do certificado    
    try:
        senha = nome_arquivo.split("_")[1].replace(".pfx", "")
    except IndexError:
        print(f"Nome do certificado mal formatado ou não encontrado: {nome_arquivo}")
        continue
    
    caminho_pfx = os.path.join(CAMINHO_CERTIFICADOS, nome_arquivo)
    nome_limpo = nome_empresa.replace(" ", "_").replace("-","").replace("/","_")
        
    cert_pem = os.path.join(CAMINHO_CONVERTIDO, f"cert-{nome_limpo}.pem")
    key_pem = os.path.join(CAMINHO_CONVERTIDO, f"key-{nome_limpo}.pem")
    
    print(f"\n🔧 Convertendo certificado para empresa: {nome_empresa}")
    print(f"📄 Arquivo PFX: {caminho_pfx}")
    print(f"🔐 Senha extraída: {senha}")
    
    subprocess.run([
        "openssl", "pkcs12", "-in", caminho_pfx,
        "-clcerts", "-nokeys", "-out", cert_pem,
        "-passin", f"pass:{senha}"
    ], shell=True)
    
    subprocess.run([
        "openssl", "pkcs12", "-in", caminho_pfx,
        "-nocerts", "-nodes", "-out", key_pem,
        "-passin", f"pass:{senha}"
    ], shell=True)