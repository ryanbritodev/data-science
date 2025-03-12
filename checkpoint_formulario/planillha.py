import os
import ssl
import urllib3
import openpyxl
from shareplum import Site
from shareplum import Office365
from shareplum.site import Version
from dotenv import load_dotenv

# Carregando variáveis de ambiente (credenciais)
load_dotenv()

# Configurações para certificação SSL (Firewall da FIAP)
urllib3.disable_warnings()  # Desabilita avisos sobre certificação SSL
ssl._create_default_https_context = ssl._create_unverified_context  # Contexto da certificação SSL para não verificar certificados

# URL para Planilha do Excel com a planilha atualizada pelo Microsoft Forms
sharepoint_url = "https://fiapcom-my.sharepoint.com"
site_url = "https://fiapcom-my.sharepoint.com/personal/rm554497_fiap_com_br"

# Caminho do Arquivo
caminho_arquivo = "Documents/Impacto do Trabalho Remoto na Eficiência do Trabalhador.xlsx"
arquivo_destino = "planilha.xlsx"

# Credenciais do Microsoft 365
usuario = os.getenv("USUARIO")
senha = os.getenv("SENHA")

caminho_planilha = f"{os.path.dirname(__file__)}\\planilha.xlsx"

# Realizando a autenticação e realizando o download do arquivo
try:
    # Autenticação no Microsoft 365
    authcookie = Office365(sharepoint_url, username=usuario, password=senha).GetCookies()
    site = Site(site_url, version=Version.v365, authcookie=authcookie)

    # Acessando a pasta do Sharepoint
    folder = site.Folder("Documents")

    # Baixar arquivo
    file_content = folder.get_file(caminho_arquivo.split('/')[-1])

    # Salvar planilha com os dados do formulário
    with open(arquivo_destino, "wb") as f:
        f.write(file_content)

    print(f"Arquivo baixado com sucesso! Caminho: {caminho_planilha}")
# Caso ocorra algum erro no download
except Exception as e:
    print(f"Erro ao baixar o arquivo: {str(e)}")
    import traceback

    traceback.print_exc()

# Abrindo planilha com os dados
objeto_workbook = openpyxl.load_workbook(caminho_planilha)
objeto_planilha = objeto_workbook.active

# Acessando célula para testes
celula_teste = objeto_planilha.cell(row=2, column=7)
print(celula_teste.value)

