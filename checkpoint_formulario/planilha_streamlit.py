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

caminho_planilha = f"{os.path.dirname(__file__)}\\planilha.xlsx"


def baixar_planilha():
    """
    --> Função para baixar a planilha do Microsoft 365 no SharePoint
    """
    # Carregando variáveis de ambiente (credenciais de usuário)
    load_dotenv()

    # Configurações do certificado SSL (Firewall da FIAP)
    urllib3.disable_warnings()  # Desabilita os avisos do certificado SSL
    ssl._create_default_https_context = ssl._create_unverified_context  # Criando um contexto que não verifica a certificação SSL

    # URL da Planilha do Excel no Sharepoint atualizada pelo Microsoft Forms
    sharepoint_url = "https://fiapcom-my.sharepoint.com"
    site_url = "https://fiapcom-my.sharepoint.com/personal/rm554497_fiap_com_br"

    # Caminho do Arquivo dentro do Sharepoint
    caminho_arquivo = "Documents/Impacto do Trabalho Remoto na Eficiência do Trabalhador.xlsx"
    arquivo_destino = "planilha.xlsx"

    # Credenciais do Microsoft 365
    usuario = os.getenv("USUARIO")
    senha = os.getenv("SENHA")

    try:
        # Autenticação no Microsoft 365
        authcookie = Office365(sharepoint_url, username=usuario, password=senha).GetCookies()
        site = Site(site_url, version=Version.v365, authcookie=authcookie)

        # Acessando a pasta do Sharepoint
        folder = site.Folder("Documents")

        # Baixando arquivo (get)
        file_content = folder.get_file(caminho_arquivo.split('/')[-1])

        # Salvar planilha com os dados do formulário
        with open(arquivo_destino, "wb") as f:
            f.write(file_content)

        st.success(f'Arquivo baixado com sucesso! Caminho do arquivo: "{caminho_planilha}"')
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

        # Iterando sobre todos os valores das colunas e armazenando em listas
        valores_idade = []
        for celula in objeto_planilha["G"]:
            if celula.value is not None:  # Verifica se a célula não está vazia
                valores_idade.append(celula.value)

        valores_frequencia = []
        for celula in objeto_planilha["H"]:
            if celula.value is not None:  # Verifica se a célula não está vazia
                valores_frequencia.append(celula.value)

        valores_produtividade = []
        for celula in objeto_planilha["I"]:
            if celula.value is not None:  # Verifica se a célula não está vazia
                # Converte para int se for número
                if isinstance(celula.value, (int, float)):
                    valores_produtividade.append(int(celula.value))
                else:
                    try:
                        valores_produtividade.append(int(celula.value))
                    except:
                        pass  # Ignora valores que não podem ser convertidos

        # Remove o cabeçalho
        if valores_idade and isinstance(valores_idade[0], str) and not any(
                d in valores_idade[0].lower() for d in ['ano', '18-', '25-', '35-']):
            valores_idade.pop(0)

        # Remove o cabeçalho das frequências
        if valores_frequencia and isinstance(valores_frequencia[0], str):
            valores_frequencia.pop(0)

        # Fechando o workbook da Planilha
        objeto_workbook.close()

        return valores_idade, valores_frequencia, valores_produtividade
    except Exception as e:
        st.error(f"Erro ao ler a planilha: {str(e)}")
        return [], []


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

    # Verificação da existência do caminho da planilha
    if os.path.exists(caminho_planilha):
        pass
    else:
        baixar_planilha()

    # Título do dashboard
    st.title("📊 Dashboard Trabalho Remoto")

    # Sidebar para opções
    st.sidebar.header("⚙️ Opções")

    # Botão para atualizar os dados (executa função para baixar a planilha novamente
    if st.sidebar.button("🔄 Atualizar Dados"):
        arquivo = baixar_planilha()
        if arquivo:
            st.sidebar.success("Dados atualizados com sucesso!")

    valores_idade, valores_frequencia, valores_produtividade = ler_dados_planilha()

    # Lista com as frequências esperadas (exceto "Outra")
    frequencias_validas = [
        'Trabalho 100% remoto',
        'Trabalho em regime híbrido (parte presencial, parte remoto)',
        'Trabalho principalmente presencial, mas ocasionalmente remoto',
        'Trabalho exclusivamente presencial'
    ]

    def classificar_frequencia(texto):
        if texto in frequencias_validas:
            return texto
        else:
            return 'Outra'

    # Aplica a classificação na lista de frequências coletadas
    valores_frequencia_classificados = [classificar_frequencia(texto) for texto in valores_frequencia]

    if not valores_idade or not valores_frequencia:
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
    ordem_faixas = ['18-24 anos', '25-34 anos', '35-44 anos',
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

    total_respostas = df_idades['Quantidade'].sum()

    # Coluna 1: Tabela de Dados
    with col1:
        st.title(f"Total de Respostas: {total_respostas}")
        st.header("⏳ Idade")
        st.image(f"{os.path.dirname(__file__)}\\assets\\idades.png", caption="Idade dos Participantes da Pesquisa")
        st.subheader("Tabela de Distribuição")

        # Adiciona coluna de percentual
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
            faixa_mais_comum = df_idades.loc[df_idades['Quantidade'].idxmax(), 'Faixa Etária']
            st.metric("Faixa Etária Mais Comum (Moda)", faixa_mais_comum)

        with metricas_col2:
            maior_quantidade = df_idades['Quantidade'].max()
            percentual_maior = (maior_quantidade / total_respostas) * 100
            st.metric("Representação da Faixa Mais Comum", f"{percentual_maior:.1f}%")

        with metricas_col3:
            # Dicionário para mapear cada faixa etária para um valor aproximado
            valores_medios = {
                '18-24 anos': 21,
                '25-34 anos': .5,
                '35-44 anos': 39.5,
                '45-54 anos': 49.5,
                '55-64 anos': 59.5,
                '65 anos ou mais': 70
            }

            # Calcular a média ponderada
            soma_ponderada = 0
            for i, row in df_idades.iterrows():
                faixa = row['Faixa Etária']
                if faixa in valores_medios:
                    soma_ponderada += valores_medios[faixa] * row['Quantidade']

            media_ponderada = soma_ponderada / total_respostas

            st.metric("Média de Idade (estimada)", f"{media_ponderada:.1f} anos")

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

    # Separador
    st.markdown("---")

    # Contagem das frequências de trabalho remoto
    contagem_frequencia = Counter(valores_frequencia_classificados)

    # Criar um DataFrame para facilitar a visualização
    df_frequencia = pd.DataFrame({
        'Frequência': list(contagem_frequencia.keys()),
        'Quantidade': list(contagem_frequencia.values())
    })

    # Ordenar o DataFrame por frequência de trabalho remoto de forma lógica
    ordem_frequencia = ['Trabalho 100% remoto', 'Trabalho em regime híbrido (parte presencial, parte remoto)', 'Trabalho principalmente presencial, mas ocasionalmente remoto',
                        'Trabalho exclusivamente presencial', 'Outra']

    # Filtra apenas as frequências que existem nos dados
    ordem_frequencia_filtrada = [freq for freq in ordem_frequencia if freq in df_frequencia['Frequência'].values]

    # Ordena o DataFrame se houver frequências válidas
    if ordem_frequencia_filtrada:
        df_frequencia['Frequência'] = pd.Categorical(df_frequencia['Frequência'],
                                                     categories=ordem_frequencia_filtrada,
                                                     ordered=True)
        df_frequencia = df_frequencia.sort_values('Frequência')

    # Criar layout da segunda seção em colunas
    freq_col1, freq_col2 = st.columns(2)

    # Coluna 1: Tabela de Dados de Frequência
    with freq_col1:

        # Título da seção
        st.header("🏠 Frequência de Trabalho Remoto")

        # Imagem
        st.image(f"{os.path.dirname(__file__)}\\assets\\remoto.png", caption="Frequência de Trabalho Remoto")

        st.subheader("Tabela de Distribuição")

        # Adiciona coluna de percentual
        total_respostas_freq = df_frequencia['Quantidade'].sum()
        df_frequencia['Percentual'] = df_frequencia['Quantidade'].apply(
            lambda x: f"{(x / total_respostas_freq * 100):.1f}%")

        # Exibe a tabela formatada
        st.dataframe(
            df_frequencia,
            column_config={
                "Quantidade": st.column_config.NumberColumn(format="%d"),
            },
            use_container_width=True,
            hide_index=True
        )

        # Adiciona métricas
        st.subheader("Métricas")
        freq_metricas_col1, freq_metricas_col2 = st.columns(2)

        with freq_metricas_col1:
            freq_mais_comum = df_frequencia.loc[df_frequencia['Quantidade'].idxmax(), 'Frequência']
            st.metric("Frequência Mais Comum", freq_mais_comum)

        with freq_metricas_col2:
            maior_quant_freq = df_frequencia['Quantidade'].max()
            percentual_maior_freq = (maior_quant_freq / total_respostas_freq) * 100
            st.metric("Representação", f"{percentual_maior_freq:.1f}%")


    # Coluna 2: Visualizações Gráficas de Frequência
    with freq_col2:
        # Tipo de gráfico (com radio buttons)
        tipo_grafico_freq = st.radio(
            "Selecione o tipo de gráfico:",
            ["Gráfico de Barras", "Gráfico de Pizza", "Treemap", "Funil"],
            horizontal=True,
            key="grafico_frequencia"  # Chave única para este componente
        )

        if tipo_grafico_freq == "Gráfico de Barras":
            fig = px.bar(
                df_frequencia,
                x='Frequência',
                y='Quantidade',
                text='Quantidade',
                color='Frequência',
                title="Distribuição por Frequência de Trabalho Remoto",
                height=400
            )
            fig.update_layout(xaxis_title="Frequência de Trabalho Remoto", yaxis_title="Quantidade")
            st.plotly_chart(fig, use_container_width=True)

        elif tipo_grafico_freq == "Gráfico de Pizza":
            fig = px.pie(
                df_frequencia,
                values='Quantidade',
                names='Frequência',
                title="Distribuição por Frequência de Trabalho Remoto (%)",
                height=400
            )
            fig.update_traces(textposition='inside', textinfo='percent+label')
            st.plotly_chart(fig, use_container_width=True)

        elif tipo_grafico_freq == "Treemap":
            fig = px.treemap(
                df_frequencia,
                path=['Frequência'],
                values='Quantidade',
                title="Distribuição por Frequência de Trabalho Remoto",
                height=400
            )
            fig.update_traces(textinfo='label+percent entry')
            st.plotly_chart(fig, use_container_width=True)

        else:  # Funil
            fig = px.funnel(
                df_frequencia,
                x='Quantidade',
                y='Frequência',
                title="Distribuição por Frequência de Trabalho Remoto",
                height=400
            )
            st.plotly_chart(fig, use_container_width=True)

        # Análise adicional
        st.subheader("Análise Expandida")

        # Mostra os dados brutos se solicitado
        if st.checkbox("Mostrar dados brutos", key="mostrar_dados_frequencia"):
            st.write("Lista de todas as frequências coletadas:")
            st.write(valores_frequencia)

    # --- Seção de Produtividade ---
    st.markdown("---")
    st.header("📈 Produtividade")

    # Cria DataFrame com contagem das respostas (valores de 1 a 5)
    contagem_produtividade = Counter(valores_produtividade)
    df_produtividade = pd.DataFrame({
        'Produtividade': list(contagem_produtividade.keys()),
        'Quantidade': list(contagem_produtividade.values())
    })

    # Ordena as classificações de 1 a 5
    ordem_produtividade = [1, 2, 3, 4, 5]
    df_produtividade['Produtividade'] = pd.Categorical(
        df_produtividade['Produtividade'],
        categories=ordem_produtividade,
        ordered=True
    )
    df_produtividade = df_produtividade.sort_values('Quantidade')

    # Calcula a média geral de produtividade
    media_produtividade = sum(valores_produtividade) / len(valores_produtividade)

    # Cria layout em duas colunas para a seção de produtividade
    prod_col1, prod_col2 = st.columns(2)

    with prod_col1:
        # Imagem
        st.image(f"{os.path.dirname(__file__)}\\assets\\produtividade.png", caption="Produtividade no Trabalho Remoto")
        st.subheader("Tabela de Distribuição")
        total_respostas_prod = df_produtividade['Quantidade'].sum()
        df_produtividade['Percentual'] = df_produtividade['Quantidade'].apply(
            lambda x: f"{(x / total_respostas_prod * 100):.1f}%"
        )
        st.dataframe(
            df_produtividade,
            column_config={"Quantidade": st.column_config.NumberColumn(format="%d")},
            use_container_width=True,
            hide_index=True
        )
        st.subheader("Métricas")
        st.metric("Média de Produtividade", f"{media_produtividade:.1f}")

    with prod_col2:
        st.subheader("Visualização Gráfica")
        tipo_grafico_prod = st.radio(
            "Selecione o tipo de gráfico:",
            ["Gráfico de Barras", "Gráfico de Pizza", "Treemap", "Funil"],
            horizontal=True,
            key="grafico_produtividade"
        )
        if tipo_grafico_prod == "Gráfico de Barras":
            fig = px.bar(
                df_produtividade,
                x='Produtividade',
                y='Quantidade',
                text='Quantidade',
                color='Produtividade',
                title="Distribuição de Produtividade",
                height=400
            )
            fig.update_layout(xaxis_title="Classificação", yaxis_title="Quantidade")
            st.plotly_chart(fig, use_container_width=True)
        elif tipo_grafico_prod == "Gráfico de Pizza":
            fig = px.pie(
                df_produtividade,
                values='Quantidade',
                names='Produtividade',
                title="Distribuição de Produtividade (%)",
                height=400
            )
            fig.update_traces(textposition='inside', textinfo='percent+label')
            st.plotly_chart(fig, use_container_width=True)
        elif tipo_grafico_prod == "Treemap":
            fig = px.treemap(
                df_produtividade,
                path=['Produtividade'],
                values='Quantidade',
                title="Distribuição de Produtividade",
                height=400
            )
            fig.update_traces(textinfo='label+percent entry')
            st.plotly_chart(fig, use_container_width=True)
        else:  # Gráfico Funil
            fig = px.funnel(
                df_produtividade,
                x='Quantidade',
                y='Produtividade',
                title="Distribuição de Produtividade",
                height=400
            )
            st.plotly_chart(fig, use_container_width=True)

        # Opção para mostrar os dados brutos
        if st.checkbox("Mostrar dados brutos de Produtividade", key="mostrar_produtividade"):
            st.write("Lista de classificações de produtividade:")
            st.write(valores_produtividade)

    # Sobre o dashboard
    st.sidebar.markdown("---")
    st.sidebar.subheader("Sobre")
    st.sidebar.markdown(
        """
        Este dashboard apresenta a distribuição de idades dos participantes
        da pesquisa sobre ["Impacto do Trabalho Remoto na Eficiência do Trabalhador"](https://forms.office.com/Pages/ResponsePage.aspx?id=4r_bEbiJSUW-EM7DZOWVUQQHfmHLeF1GuQtv6hEPk_xUQk1ZVzJaR1FMUjlXRzZEMDBXTFdFME5LUi4u).

        Os dados são carregados diretamente do SharePoint.
        """
    )
    st.sidebar.markdown("---")
    st.sidebar.subheader("Participantes")
    st.sidebar.markdown(
        """
        - Prof. Dr. Marcos Crivelaro - PF0076
        - Arthur Cotrick Pagani - RM554510
        - Diogo Leles Franciulli - RM558487
        - Felipe Sousa de Oliveira - RM559085
        - Ryan Brito Pereira Ramos - RM554497
        - Vitor Chaves - RM557067
        """
    )


# Executa a aplicação somente se for executado diretamente
if __name__ == "__main__":
    main()
