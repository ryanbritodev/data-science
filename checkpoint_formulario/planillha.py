from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File

# URL para Planilha do Excel com a planilha atualizada pelo Microsoft Forms
url_base = "https://fiapcom-my.sharepoint.com"
caminho_planilha = "/:x:/g/personal/rm554497_fiap_com_br/EaY8NqVSNJ9AhJstgu8x8nMB4y-nMoS8b4iV4H_WVCfYdw?e=ujr2NT"

# Credenciais
usuario = ""
senha = ""

# Contexto para Acessar o Sharepoint
contexto_cliente = ClientContext(url_base + caminho_planilha).with_credentials(usuario, senha)


