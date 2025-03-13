import streamlit as st
import pandas as pd
import plotly.express as px
import os
import ssl
import urllib3
import traceback
import openpyxl
from collections import Counter
from shareplum import Site
from shareplum import Office365
from shareplum.site import Version
from dotenv import load_dotenv


def baixar_planilha():
    """
    --> Função para baixar a planilha do SharePoint
    """
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

        st.success(f"Arquivo baixado com sucesso! Caminho {caminho_planilha}")
        return arquivo_destino
    except Exception as e:
        st.error(f"Erro ao baixar o arquivo: {str(e)}")
        traceback.print_exc()
        return None


def ler_dados_planilha(caminho_planilha=None):
    """
    --> Função para ler os dados da planilha já baixada
    :param caminho_planilha: Caminho para planilha que deseja realizar a leitura (valor padrão: None)
    """
    if caminho_planilha is None:
        caminho_planilha = "planilha.xlsx"

    try:
        # Abrindo planilha com os dados
        objeto_workbook = openpyxl.load_workbook(caminho_planilha)
        objeto_planilha = objeto_workbook.active

        # Iterando sobre todos os valores da coluna das idades (coluna G) e armazenando em uma lista
        valores_idade = []
        for celula in objeto_planilha["G"]:
            if celula.value is not None:  # Verifica se a célula não está vazia
                valores_idade.append(celula.value)

        # Remove o cabeçalho
        if valores_idade and isinstance(valores_idade[0], str) and not any(
                d in valores_idade[0].lower() for d in ['ano', '18-', '25-', '35-']):
            valores_idade.pop(0)

        # Fechando o workbook da Planilha
        objeto_workbook.close()

        return valores_idade
    except Exception as e:
        st.error(f"Erro ao ler a planilha: {str(e)}")
        return []


def main():
    """
    --> Função com a execução da interface principal usando o Streamlit
    """
    # Configuração da página
    st.set_page_config(
        page_title="Dashboard Trabalho Remoto",
        page_icon="📊",
        layout="wide"
    )

    # Título do dashboard
    st.title("📊 Dashboard Trabalho Remoto")

    # Sidebar para opções
    st.sidebar.header("⚙️ Opções")

    # Botão para atualizar os dados
    if st.sidebar.button("📥 Baixar dados atualizados"):
        arquivo = baixar_planilha()
        if arquivo:
            st.sidebar.success("Dados atualizados com sucesso!")

    valores_idade = ler_dados_planilha()

    if not valores_idade:
        st.warning(
            "Não foi possível ler os dados da planilha. Certifique-se de que o arquivo existe ou use os dados de exemplo.")
        st.stop()

    # Contagem de faixas etárias
    contagem_idades = Counter(valores_idade)

    # Criar um DataFrame para facilitar a visualização
    df_idades = pd.DataFrame({
        'Faixa Etária': list(contagem_idades.keys()),
        'Quantidade': list(contagem_idades.values())
    })

    # Ordenar o DataFrame pelas faixas etárias de forma lógica
    ordem_faixas = ['Menos de 18 anos', '18-24 anos', '25-34 anos', '35-44 anos',
                    '45-54 anos', '55-64 anos', '65 anos ou mais']

    # Filtra apenas as faixas que existem nos dados
    ordem_faixas_filtrada = [faixa for faixa in ordem_faixas if faixa in df_idades['Faixa Etária'].values]

    # Ordena o DataFrame se houver faixas etárias válidas
    if ordem_faixas_filtrada:
        df_idades['Faixa Etária'] = pd.Categorical(df_idades['Faixa Etária'],
                                                   categories=ordem_faixas_filtrada,
                                                   ordered=True)
        df_idades = df_idades.sort_values('Faixa Etária')

    # Criar layout do dashboard em colunas
    col1, col2 = st.columns(2)

    # Coluna 1: Tabela de Dados
    with col1:
        st.header("⏳ Idade")
        st.image(f"{os.path.dirname(__file__)}\\assets\\idades.png", caption="Idade dos Participantes da Pesquisa")
        st.subheader("Tabela de Distribuição")

        # Adiciona coluna de percentual
        total_respostas = df_idades['Quantidade'].sum()
        df_idades['Percentual'] = df_idades['Quantidade'].apply(lambda x: f"{(x / total_respostas * 100):.1f}%")

        # Exibe a tabela formatada
        st.dataframe(
            df_idades,
            column_config={
                "Quantidade": st.column_config.NumberColumn(format="%d"),
            },
            use_container_width=True,
            hide_index=True
        )

        # Adiciona métricas
        st.subheader("Métricas")
        metricas_col1, metricas_col2, metricas_col3 = st.columns(3)

        with metricas_col1:
            st.metric("Total de Respostas", total_respostas)

        with metricas_col2:
            faixa_mais_comum = df_idades.loc[df_idades['Quantidade'].idxmax(), 'Faixa Etária']
            st.metric("Faixa Etária Mais Comum", faixa_mais_comum)

        with metricas_col3:
            maior_quantidade = df_idades['Quantidade'].max()
            percentual_maior = (maior_quantidade / total_respostas) * 100
            st.metric("Representação da Faixa Mais Comum", f"{percentual_maior:.1f}%")

    # Coluna 2: Visualizações Gráficas
    with col2:
        # Tipo de gráfico (com radio buttons)
        tipo_grafico = st.radio(
            "Selecione o tipo de gráfico:",
            ["Gráfico de Barras", "Gráfico de Pizza", "Treemap", "Funil"],
            horizontal=True
        )

        if tipo_grafico == "Gráfico de Barras":
            fig = px.bar(
                df_idades,
                x='Faixa Etária',
                y='Quantidade',
                text='Quantidade',
                color='Faixa Etária',
                title="Distribuição por Faixa Etária",
                height=400
            )
            fig.update_layout(xaxis_title="Faixa Etária", yaxis_title="Quantidade")
            st.plotly_chart(fig, use_container_width=True)

        elif tipo_grafico == "Gráfico de Pizza":
            fig = px.pie(
                df_idades,
                values='Quantidade',
                names='Faixa Etária',
                title="Distribuição por Faixa Etária (%)",
                height=400
            )
            fig.update_traces(textposition='inside', textinfo='percent+label')
            st.plotly_chart(fig, use_container_width=True)

        elif tipo_grafico == "Treemap":
            fig = px.treemap(
                df_idades,
                path=['Faixa Etária'],
                values='Quantidade',
                title="Distribuição por Faixa Etária",
                height=400
            )
            fig.update_traces(textinfo='label+percent entry')
            st.plotly_chart(fig, use_container_width=True)

        else:
            fig = px.funnel(
                df_idades,
                x='Quantidade',
                y='Faixa Etária',
                title="Distribuição por Faixa Etária",
                height=400
            )
            st.plotly_chart(fig, use_container_width=True)

        # Distribuições adicionais e análises
        st.subheader("Análise Expandida")

        # Mostra os dados brutos se solicitado
        if st.checkbox("Mostrar dados brutos"):
            st.write("Lista de todas as idades coletadas:")
            st.write(valores_idade)

    # Sobre o dashboard
    st.sidebar.markdown("---")
    st.sidebar.subheader("Sobre")
    st.sidebar.caption(
        """
        Este dashboard apresenta a distribuição de idades dos participantes
        da pesquisa sobre "Impacto do Trabalho Remoto na Eficiência do Trabalhador".

        Os dados são carregados diretamente do SharePoint.
        """
    )


# Executa a aplicação
if __name__ == "__main__":
    main()
