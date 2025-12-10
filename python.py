import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from datetime import datetime
import numpy as np
from io import BytesIO
import os
import re
import plotly.io as pio
from jinja2 import Template

# Configuração da página
st.set_page_config(
    page_title="Relatório de Contatos - CARGAS NITERÓI",
    page_icon="📊",
    layout="wide"
)

# Cabeçalho
st.title("📊 Relatório de Contatos - CARGAS NITERÓI")
st.markdown("**Análise feita por Kaynan Monteiro e David Florencio**")
st.markdown("---")

# Função para normalizar texto (remover acentos e maiúsculas)
def normalize_text(text):
    if pd.isna(text):
        return text
    text = str(text)
    # Remover espaços extras
    text = text.strip()
    # Converter para minúsculas
    text = text.lower()
    # Remover acentos
    text = re.sub(r'[áàãâä]', 'a', text)
    text = re.sub(r'[éèêë]', 'e', text)
    text = re.sub(r'[íìîï]', 'i', text)
    text = re.sub(r'[óòõôö]', 'o', text)
    text = re.sub(r'[úùûü]', 'u', text)
    text = re.sub(r'[ç]', 'c', text)
    return text

# Função para normalizar situação
def normalize_situacao(situacao):
    if pd.isna(situacao):
        return "Não informado"
    
    situacao = str(situacao).strip()
    situacao_normalizada = normalize_text(situacao)
    
    # Mapeamento de situações equivalentes
    if situacao_normalizada in ['não atende', 'nao atende', 'não atend', 'n atend']:
        return "Não atende"
    elif situacao_normalizada in ['não acatou', 'nao acatou', 'n acatou']:
        return "Não acatou"
    elif situacao_normalizada in ['número incorreto', 'numero incorreto', 'tel errado', 'telefone incorreto']:
        return "Número incorreto"
    elif situacao_normalizada in ['baixada', 'empresa baixada']:
        return "Baixada"
    elif 'retornar' in situacao_normalizada:
        return "Retornar em horário"
    
    return situacao

# Função para parsear data/hora
def parse_datetime(value):
    if pd.isna(value):
        return None
    
    if isinstance(value, datetime):
        return value
    
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()
    
    value_str = str(value).strip()
    
    # Padrões comuns encontrados nos dados
    patterns = [
        '%Y-%m-%d %H:%M:%S',  # 2025-09-03 09:31:55
        '%d/%m/%Y %H:%M',     # 03/09/2025 09:31
        '%d/%m - %H:%M',      # 14/10 - 00:00
        '%d/%m/%Y - %H:%M',   # 02/10/2025 - 15:33
        '%d/%m - %H:%M',      # 07/10 - 15:00
        '%d/%m/%y %H:%M',     # 03/09/25 09:31
        '%Y-%m-%d',           # 2025-09-03
        '%d/%m/%Y',           # 03/09/2025
    ]
    
    for pattern in patterns:
        try:
            # Para padrão com ano incompleto, ajustar
            if pattern == '%d/%m/%y %H:%M' and len(value_str.split()[0].split('/')[2]) == 2:
                # Adicionar século 20 se ano for menor que 50
                parts = value_str.split()
                date_part = parts[0]
                time_part = parts[1] if len(parts) > 1 else '00:00'
                day, month, year = date_part.split('/')
                year_full = f"20{year}" if int(year) < 50 else f"19{year}"
                value_str = f"{day}/{month}/{year_full} {time_part}"
            
            return datetime.strptime(value_str, pattern)
        except:
            continue
    
    # Tentar extrair hora de formato "dd/mm - hh:mm" sem ano
    if '- ' in value_str and ':' in value_str:
        try:
            # Adicionar ano atual
            parts = value_str.split(' - ')
            if len(parts) == 2:
                time_part = parts[1].strip()
                if ':' in time_part:
                    hour = int(time_part.split(':')[0])
                    # Criar datetime com data fictícia (usaremos só a hora)
                    return datetime(2025, 1, 1, hour, 0)
        except:
            pass
    
    return None

# Função para carregar os dados do Excel
@st.cache_data
def load_excel_data(file_path):
    """
    Carrega os dados do arquivo Excel com múltiplas planilhas
    """
    try:
        # Ler todas as planilhas
        excel_file = pd.ExcelFile(file_path)
        
        # Carregar cada planilha
        dfs = {}
        
        for sheet_name in excel_file.sheet_names:
            # Tentar ler com diferentes cabeçalhos
            try:
                # Tentar ler com header=1 (segunda linha)
                df = pd.read_excel(excel_file, sheet_name=sheet_name, header=1)
            except:
                # Se falhar, tentar com header=0
                df = pd.read_excel(excel_file, sheet_name=sheet_name, header=0)
            
            # Limpar nomes das colunas
            df.columns = [str(col).strip() for col in df.columns]
            
            # Remover linhas completamente vazias
            df = df.dropna(how='all')
            
            # Armazenar com nome da planilha
            dfs[sheet_name] = df
            
            st.sidebar.success(f"✅ {sheet_name}: {len(df)} registros")
        
        return dfs
        
    except Exception as e:
        st.error(f"Erro ao carregar arquivo Excel: {e}")
        st.error(f"Detalhes: {str(e)}")
        return {}

# Função para processar e limpar os dados
def clean_data(df, sheet_name):
    """
    Limpa e padroniza os dados
    """
    # Fazer uma cópia
    df_clean = df.copy()
    
    # Normalizar nomes das colunas
    column_mapping = {
        'CNPJ': ['CNPJ'],
        'RAZÃO SOCIAL': ['RAZÃO SOCIAL', 'RAZÃO SOCIAL'],
        'TEL 1': ['TEL 1', 'TEL1', 'TEL 1'],
        'TEL 2': ['TEL 2', 'TEL2', 'TEL 2'],
        'E-MAIL': ['E-MAIL', 'E-MAIL', 'EMAIL'],
        'SITUAÇÃO': ['SITUAÇÃO', 'SITUAÇÃO', 'SITUACAO'],
        'OBSERVAÇÃO': ['OBSERVAÇÃO', 'OBSERVAÇÃO', 'OBSERVACAO']
    }
    
    # Para colunas de data/hora (apenas na primeira planilha)
    date_columns = []
    if sheet_name == 'CARGAS_NITEROI':
        date_columns = ['Data / Hora 1', 'Data / Hora 2', 'Data / Hora 3']
        # Verificar se as colunas existem com nomes diferentes
        for i in range(1, 4):
            possible_names = [f'Data / Hora {i}', f'Data_Hora_{i}', f'Data Hora {i}', f'Data_Hora {i}']
            for name in possible_names:
                if name in df_clean.columns:
                    date_columns.append(name)
    
    # Processar colunas de data/hora
    for col in date_columns:
        if col in df_clean.columns:
            # Converter para datetime
            df_clean[col] = df_clean[col].apply(parse_datetime)
    
    # Normalizar situação
    if 'SITUAÇÃO' in df_clean.columns:
        df_clean['SITUAÇÃO_NORMALIZADA'] = df_clean['SITUAÇÃO'].apply(normalize_situacao)
    else:
        # Procurar por coluna de situação com nome diferente
        for col in df_clean.columns:
            if 'situação' in normalize_text(col) or 'situacao' in normalize_text(col):
                df_clean['SITUAÇÃO_NORMALIZADA'] = df_clean[col].apply(normalize_situacao)
                break
    
    # Limpar valores de telefone
    for tel_col in ['TEL 1', 'TEL 2']:
        if tel_col in df_clean.columns:
            df_clean[tel_col] = df_clean[tel_col].astype(str).str.strip()
            # Converter valores numéricos para string
            df_clean[tel_col] = df_clean[tel_col].apply(
                lambda x: str(int(float(x))) if isinstance(x, (int, float)) and not pd.isna(x) else x
            )
            df_clean[tel_col] = df_clean[tel_col].replace(['nan', 'None', 'NaN', 'NaT', 'nat', ''], None)
    
    # Limpar email
    if 'E-MAIL' in df_clean.columns:
        df_clean['E-MAIL'] = df_clean['E-MAIL'].astype(str).str.strip()
        df_clean['E-MAIL'] = df_clean['E-MAIL'].replace(['nan', 'None', 'NaN', 'NaT', 'nat', ''], None)
    
    return df_clean

# Função para gerar gráfico de pizza
def create_pie_chart(df, title):
    """
    Cria gráfico de pizza para distribuição de situações
    """
    if 'SITUAÇÃO_NORMALIZADA' not in df.columns:
        return None
    
    situacao_counts = df['SITUAÇÃO_NORMALIZADA'].value_counts().reset_index()
    situacao_counts.columns = ['SITUAÇÃO', 'QUANTIDADE']
    
    # Ordenar por quantidade (decrescente)
    situacao_counts = situacao_counts.sort_values('QUANTIDADE', ascending=False)
    
    fig = px.pie(
        situacao_counts, 
        values='QUANTIDADE', 
        names='SITUAÇÃO',
        title=f"<b>{title}</b>",
        color_discrete_sequence=px.colors.qualitative.Set3,
        hover_data=['QUANTIDADE']
    )
    fig.update_traces(
        textposition='inside', 
        textinfo='percent+label',
        hovertemplate='<b>%{label}</b><br>Quantidade: %{value}<br>Percentual: %{percent:.1%}',
        textfont=dict(size=12)
    )
    fig.update_layout(
        height=500,
        showlegend=True,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=-0.2,  # Ajustado para não sobrepor
            xanchor="center",
            x=0.5
        ),
        title=dict(
            x=0.5,
            xanchor='center',
            font=dict(size=16)
        ),
        margin=dict(t=80, b=100, l=20, r=20)  # Margem ajustada
    )
    return fig

# Função para gerar gráfico de colunas de horários
def create_calls_chart(df, title):
    """
    Cria gráfico de colunas para horários de ligações
    """
    horas_todas = []
    
    # Verificar todas as colunas que podem ser de data/hora
    date_cols = []
    for col in df.columns:
        col_lower = str(col).lower()
        if any(term in col_lower for term in ['data', 'hora', 'data_hora', 'data / hora']):
            date_cols.append(col)
    
    # Se não encontrou automaticamente, procurar colunas específicas
    if not date_cols:
        for i in range(1, 4):
            for pattern in [f'Data / Hora {i}', f'Data_Hora_{i}', f'Data Hora {i}']:
                if pattern in df.columns:
                    date_cols.append(pattern)
    
    for col in date_cols:
        if col in df.columns:
            col_data = df[col].dropna()
            
            for value in col_data:
                dt = parse_datetime(value)
                if dt:
                    horas_todas.append(dt.hour)
    
    if horas_todas:
        horas_df = pd.DataFrame({'HORA': horas_todas})
        horas_counts = horas_df['HORA'].value_counts().sort_index().reset_index()
        horas_counts.columns = ['HORA', 'QUANTIDADE']
        
        fig = px.bar(
            horas_counts,
            x='HORA',
            y='QUANTIDADE',
            title=f"<b>{title}</b>",
            labels={'HORA': 'Hora do Dia', 'QUANTIDADE': 'Número de Ligações'},
            color='QUANTIDADE',
            color_continuous_scale='Blues',
            text='QUANTIDADE'
        )
        fig.update_traces(
            textposition='outside',
            hovertemplate='<b>Hora: %{x}:00</b><br>Ligações: %{y}'
        )
        fig.update_layout(
            height=500,
            xaxis=dict(
                tickmode='linear',
                dtick=1,
                title='Hora do Dia'
            ),
            yaxis=dict(title='Quantidade de Ligações'),
            title=dict(
                x=0.5,
                xanchor='center',
                font=dict(size=16)
            )
        )
        return fig
    else:
        return None

# Função para exibir observações importantes
def show_important_observations(df, title):
    """
    Exibe observações importantes (excluindo "Não atende")
    """
    if 'SITUAÇÃO_NORMALIZADA' not in df.columns:
        # Procurar coluna de observação
        obs_col = None
        for col in df.columns:
            if 'observação' in normalize_text(col) or 'observacao' in normalize_text(col):
                obs_col = col
                break
        
        if not obs_col:
            st.warning(f"Dados incompletos em {title}")
            return
    
    # Usar situação normalizada
    situacao_col = 'SITUAÇÃO_NORMALIZADA'
    obs_col = 'OBSERVAÇÃO' if 'OBSERVAÇÃO' in df.columns else obs_col
    
    # Filtrar observações onde a situação NÃO é "Não atende"
    mask = df[situacao_col] != "Não atende"
    
    df_filtrado = df[mask].copy()
    
    if len(df_filtrado) > 0:
        st.subheader(f"📝 {title}")
        
        # Métricas
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.metric("Observações Importantes", len(df_filtrado))
        
        with col2:
            st.metric("Porcentagem do Total", f"{(len(df_filtrado)/len(df)*100):.1f}%")
        
        with col3:
            situacoes_unicas = df_filtrado[situacao_col].nunique()
            st.metric("Situações Diferentes", situacoes_unicas)
        
        # Resumo das situações
        st.markdown("**Situações encontradas:**")
        situacoes_counts = df_filtrado[situacao_col].value_counts()
        for situacao, count in situacoes_counts.items():
            st.markdown(f"- **{situacao}**: {count} ocorrências")
        
        # Tabela expandível
        with st.expander("🔍 Ver detalhes das observações"):
            for idx, row in df_filtrado.iterrows():
                # Obter nome da empresa
                empresa = 'Não informado'
                for col in ['RAZÃO SOCIAL', 'RAZÃO SOCIAL']:
                    if col in row and not pd.isna(row[col]):
                        empresa = row[col]
                        break
                
                # Obter CNPJ
                cnpj = 'Não informado'
                for col in ['CNPJ']:
                    if col in row and not pd.isna(row[col]):
                        cnpj = row[col]
                        break
                
                # Obter observação
                observacao = 'Sem observação'
                if obs_col in row and not pd.isna(row[obs_col]):
                    observacao = row[obs_col]
                
                st.markdown(f"### {empresa}")
                st.markdown(f"**CNPJ:** {cnpj}")
                st.markdown(f"**Situação:** `{row[situacao_col]}`")
                st.markdown(f"**Observação:** {observacao}")
                st.markdown("---")
    else:
        st.info(f"Não há observações importantes em {title} (todas são 'Não atende')")

# Função para calcular métricas
def calculate_metrics(df, sheet_name):
    """
    Calcula métricas para uma planilha
    """
    metrics = {
        'Planilha': sheet_name,
        'Total Empresas': len(df),
    }
    
    # Telefones
    tel1_count = 0
    tel2_count = 0
    if 'TEL 1' in df.columns:
        tel1_count = df['TEL 1'].notna().sum()
        metrics['Com Telefone 1'] = tel1_count
    
    if 'TEL 2' in df.columns:
        tel2_count = df['TEL 2'].notna().sum()
        metrics['Com Telefone 2'] = tel2_count
    
    metrics['Total Telefones'] = tel1_count + tel2_count
    
    # Emails
    if 'E-MAIL' in df.columns:
        email_count = df['E-MAIL'].notna().sum()
        metrics['Com Email'] = email_count
    
    # Situações
    if 'SITUAÇÃO_NORMALIZADA' in df.columns:
        metrics['Situações Únicas'] = df['SITUAÇÃO_NORMALIZADA'].nunique()
        # Adicionar contagem das principais situações
        situacao_counts = df['SITUAÇÃO_NORMALIZADA'].value_counts()
        for situacao, count in situacao_counts.head(3).items():
            metrics[f"{situacao[:15]}..."] = count
    
    return metrics

# Função para download do Excel
def get_excel_download_link(df_dict, filename):
    """
    Cria link para download do Excel
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, df in df_dict.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    output.seek(0)
    
    return output

# Função para gerar relatório HTML
def generate_html_report(dfs_clean, filename="relatorio_cargas_niteroi.html"):
    """
    Gera um relatório HTML interativo para compartilhar
    """
    # Criar DataFrame consolidado
    dfs_consolidado = []
    for sheet_name, df in dfs_clean.items():
        df_copy = df.copy()
        # Manter apenas colunas comuns
        common_cols = []
        for col in ['CNPJ', 'RAZÃO SOCIAL', 'TEL 1', 'TEL 2', 'E-MAIL', 'SITUAÇÃO_NORMALIZADA', 'OBSERVAÇÃO']:
            if col in df_copy.columns:
                common_cols.append(col)
        
        df_copy = df_copy[common_cols]
        dfs_consolidado.append(df_copy)
    
    if not dfs_consolidado:
        return None
    
    df_consolidado = pd.concat(dfs_consolidado, ignore_index=True)
    
    # Gerar gráficos
    fig_pie = create_pie_chart(df_consolidado, "Distribuição de Situações - Consolidado")
    fig_calls = create_calls_chart(dfs_clean.get('CARGAS_NITEROI', pd.DataFrame()), 
                                   "Horários de Ligações")
    
    # Converter gráficos para HTML
    pie_html = pio.to_html(fig_pie, full_html=False) if fig_pie else ""
    calls_html = pio.to_html(fig_calls, full_html=False) if fig_calls else ""
    
    # Calcular métricas
    metrics = calculate_metrics(df_consolidado, "Consolidado")
    
    # Contar observações importantes
    if 'SITUAÇÃO_NORMALIZADA' in df_consolidado.columns:
        obs_importantes = len(df_consolidado[df_consolidado['SITUAÇÃO_NORMALIZADA'] != "Não atende"])
        percentual_obs = round((obs_importantes / len(df_consolidado) * 100), 1)
        
        # Contagem de situações
        situacoes_contagem = df_consolidado[df_consolidado['SITUAÇÃO_NORMALIZADA'] != "Não atende"]['SITUAÇÃO_NORMALIZADA'].value_counts().to_dict()
    else:
        obs_importantes = 0
        percentual_obs = 0
        situacoes_contagem = {}
    
    # Template HTML
    html_template = """
    <!DOCTYPE html>
    <html lang="pt-BR">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Relatório CARGAS NITERÓI</title>
        <script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
        <style>
            body {
                font-family: Arial, sans-serif;
                margin: 20px;
                background-color: #f5f5f5;
            }
            .header {
                background-color: #2c3e50;
                color: white;
                padding: 20px;
                border-radius: 10px;
                margin-bottom: 20px;
            }
            .metrics {
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
                gap: 15px;
                margin-bottom: 30px;
            }
            .metric-card {
                background: white;
                padding: 20px;
                border-radius: 8px;
                box-shadow: 0 2px 4px rgba(0,0,0,0.1);
                text-align: center;
            }
            .metric-value {
                font-size: 2em;
                font-weight: bold;
                color: #2c3e50;
            }
            .metric-label {
                color: #7f8c8d;
                margin-top: 5px;
            }
            .chart-container {
                background: white;
                padding: 20px;
                border-radius: 8px;
                box-shadow: 0 2px 4px rgba(0,0,0,0.1);
                margin-bottom: 20px;
            }
            .observations {
                background: white;
                padding: 20px;
                border-radius: 8px;
                box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            }
            .footer {
                text-align: center;
                margin-top: 30px;
                color: #7f8c8d;
                font-size: 0.9em;
            }
            h1, h2, h3 {
                color: #2c3e50;
            }
        </style>
    </head>
    <body>
        <div class="header">
            <h1>📊 Relatório de Contatos - CARGAS NITERÓI</h1>
            <p>Análise feita por Kaynan Monteiro e David Florencio</p>
            <p>Gerado em: {{data_geracao}}</p>
        </div>
        
        <div class="metrics">
            <div class="metric-card">
                <div class="metric-value">{{total_empresas}}</div>
                <div class="metric-label">Total de Empresas</div>
            </div>
            <div class="metric-card">
                <div class="metric-value">{{total_telefones}}</div>
                <div class="metric-label">Total de Telefones</div>
            </div>
            <div class="metric-card">
                <div class="metric-value">{{total_emails}}</div>
                <div class="metric-label">Total de Emails</div>
            </div>
            <div class="metric-card">
                <div class="metric-value">{{situacoes_unicas}}</div>
                <div class="metric-label">Situações Únicas</div>
            </div>
        </div>
        
        {% if pie_html %}
        <div class="chart-container">
            <h2>Distribuição de Situações</h2>
            {{pie_html}}
        </div>
        {% endif %}
        
        {% if calls_html %}
        <div class="chart-container">
            <h2>Horários de Ligações</h2>
            {{calls_html}}
        </div>
        {% endif %}
        
        <div class="observations">
            <h2>📝 Observações Importantes</h2>
            <p>Total de observações importantes: <strong>{{obs_importantes}}</strong></p>
            <p>Percentual do total: <strong>{{percentual_obs}}%</strong></p>
            
            <h3>Resumo por Situação:</h3>
            <ul>
                {% for situacao, quantidade in situacoes_contagem.items() %}
                <li><strong>{{situacao}}:</strong> {{quantidade}} ocorrências</li>
                {% endfor %}
            </ul>
        </div>
        
        <div class="footer">
            <p>Relatório gerado automaticamente - Sistema de Análise de CARGAS NITERÓI</p>
            <p>Para atualizar os dados, execute o sistema Python com o arquivo Excel atualizado</p>
        </div>
    </body>
    </html>
    """
    
    # Renderizar template
    template = Template(html_template)
    html_content = template.render(
        data_geracao=datetime.now().strftime("%d/%m/%Y %H:%M"),
        total_empresas=len(df_consolidado),
        total_telefones=metrics.get('Total Telefones', 0),
        total_emails=metrics.get('Com Email', 0),
        situacoes_unicas=metrics.get('Situações Únicas', 0),
        pie_html=pie_html,
        calls_html=calls_html,
        obs_importantes=obs_importantes,
        percentual_obs=percentual_obs,
        situacoes_contagem=situacoes_contagem
    )
    
    # Salvar arquivo HTML
    with open(filename, 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    return filename

# Interface principal
def main():
    # Upload do arquivo
    st.sidebar.title("📂 Upload de Arquivo")
    
    uploaded_file = st.sidebar.file_uploader(
        "Carregue o arquivo Excel (NITEROI_BIRA.xlsx)",
        type=['xlsx', 'xls']
    )
    
    if uploaded_file is not None:
        # Salvar arquivo temporariamente
        with open("temp_uploaded.xlsx", "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        # Carregar dados
        with st.spinner("Carregando e processando dados..."):
            dfs = load_excel_data("temp_uploaded.xlsx")
        
        if dfs:
            # Limpar dados
            dfs_clean = {}
            for sheet_name, df in dfs.items():
                dfs_clean[sheet_name] = clean_data(df, sheet_name)
            
            # Sidebar navigation
            st.sidebar.title("Navegação")
            sheet_names = list(dfs_clean.keys())
            selected_sheet = st.sidebar.selectbox(
                "Selecione a planilha:",
                ["VISÃO GERAL"] + sheet_names
            )
            
            # Botão de download Excel
            st.sidebar.markdown("---")
            st.sidebar.subheader("📤 Exportar Dados")
            
            download_data = get_excel_download_link(dfs_clean, "dados_processados.xlsx")
            st.sidebar.download_button(
                label="📥 Baixar dados em Excel",
                data=download_data,
                file_name="dados_processados.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            # Botão para gerar relatório HTML
            st.sidebar.markdown("---")
            st.sidebar.subheader("📄 Relatório para Compartilhar")
            
            if st.sidebar.button("🔄 Gerar Relatório HTML"):
                with st.spinner("Gerando relatório HTML..."):
                    try:
                        html_file = generate_html_report(dfs_clean)
                        with open(html_file, "rb") as f:
                            st.sidebar.download_button(
                                label="⬇️ Baixar Relatório HTML",
                                data=f,
                                file_name="relatorio_cargas_niteroi.html",
                                mime="text/html"
                            )
                        st.sidebar.success("✅ Relatório HTML gerado com sucesso!")
                    except Exception as e:
                        st.sidebar.error(f"Erro ao gerar HTML: {e}")
            
            # Página: Visão Geral
            if selected_sheet == "VISÃO GERAL":
                st.header("📈 Visão Geral Consolidada")
                
                # Criar DataFrame consolidado
                dfs_consolidado = []
                for sheet_name, df in dfs_clean.items():
                    df_copy = df.copy()
                    # Manter apenas colunas comuns
                    common_cols = []
                    for col in ['CNPJ', 'RAZÃO SOCIAL', 'TEL 1', 'TEL 2', 'E-MAIL', 'SITUAÇÃO_NORMALIZADA', 'OBSERVAÇÃO']:
                        if col in df_copy.columns:
                            common_cols.append(col)
                    
                    df_copy = df_copy[common_cols]
                    df_copy['ORIGEM'] = sheet_name
                    dfs_consolidado.append(df_copy)
                
                if dfs_consolidado:
                    df_consolidado = pd.concat(dfs_consolidado, ignore_index=True)
                    
                    # Métricas gerais
                    col1, col2, col3, col4 = st.columns(4)
                    
                    with col1:
                        st.metric("Total de Empresas", len(df_consolidado))
                    
                    with col2:
                        tel1_count = df_consolidado['TEL 1'].notna().sum() if 'TEL 1' in df_consolidado.columns else 0
                        tel2_count = df_consolidado['TEL 2'].notna().sum() if 'TEL 2' in df_consolidado.columns else 0
                        st.metric("Total de Telefones", tel1_count + tel2_count)
                    
                    with col3:
                        email_count = df_consolidado['E-MAIL'].notna().sum() if 'E-MAIL' in df_consolidado.columns else 0
                        st.metric("Total de Emails", email_count)
                    
                    with col4:
                        st.metric("Planilhas", len(dfs_clean))
                    
                    st.markdown("---")
                    
                    # Gráficos
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        fig_pie = create_pie_chart(df_consolidado, "Distribuição de Situações - Consolidado")
                        if fig_pie:
                            st.plotly_chart(fig_pie, use_container_width=True)
                        else:
                            st.warning("Não há dados de situação para gráfico")
                    
                    with col2:
                        # Usar primeira planilha para horários (se tiver dados de data/hora)
                        first_sheet_name = sheet_names[0]
                        first_df = dfs_clean[first_sheet_name]
                        fig_calls = create_calls_chart(first_df, f"Horários de Ligações - {first_sheet_name}")
                        if fig_calls:
                            st.plotly_chart(fig_calls, use_container_width=True)
                        else:
                            st.info("Não foram encontrados dados de horários nas colunas de data")
                    
                    # Observações importantes
                    show_important_observations(df_consolidado, "Observações Importantes - Consolidado")
                    
                    # Tabela resumo por planilha
                    st.subheader("📋 Resumo por Planilha")
                    
                    metrics_data = []
                    for sheet_name, df in dfs_clean.items():
                        metrics = calculate_metrics(df, sheet_name)
                        metrics_data.append(metrics)
                    
                    if metrics_data:
                        resumo_df = pd.DataFrame(metrics_data)
                        st.dataframe(resumo_df, use_container_width=True, height=300)
                        
                        # Gráfico de barras comparativo
                        fig_comparativo = px.bar(
                            resumo_df,
                            x='Planilha',
                            y=['Total Empresas', 'Total Telefones'],
                            title='<b>Comparativo entre Planilhas</b>',
                            barmode='group',
                            color_discrete_sequence=px.colors.qualitative.Pastel,
                            labels={'value': 'Quantidade', 'variable': 'Métrica'}
                        )
                        fig_comparativo.update_layout(
                            height=400,
                            title=dict(x=0.5, xanchor='center')
                        )
                        st.plotly_chart(fig_comparativo, use_container_width=True)
                else:
                    st.warning("Não foi possível consolidar os dados")
            
            # Páginas individuais das planilhas
            else:
                df = dfs_clean[selected_sheet]
                
                st.header(f"📋 {selected_sheet}")
                st.caption(f"Total de registros: {len(df)}")
                
                # Exibir primeiras linhas para verificação
                with st.expander("🔍 Ver primeiras linhas da planilha"):
                    st.dataframe(df.head(), use_container_width=True)
                
                # Métricas da planilha
                col1, col2, col3, col4 = st.columns(4)
                
                metrics = calculate_metrics(df, selected_sheet)
                
                with col1:
                    st.metric("Empresas", metrics['Total Empresas'])
                
                with col2:
                    st.metric("Telefones", metrics.get('Total Telefones', 0))
                
                with col3:
                    st.metric("Emails", metrics.get('Com Email', 0))
                
                with col4:
                    if 'Situações Únicas' in metrics:
                        st.metric("Situações", metrics['Situações Únicas'])
                
                st.markdown("---")
                
                # Gráficos específicos da planilha
                col1, col2 = st.columns(2)
                
                with col1:
                                    fig_pie = create_pie_chart(df, f"Distribuição de Situações - {selected_sheet}")
                if fig_pie:
                        st.plotly_chart(fig_pie, use_container_width=True)
                
                with col2:
                    if selected_sheet == 'CARGAS_NITEROI':
                        # Verificar colunas de data disponíveis
                        date_cols_info = []
                        for col in df.columns:
                            if any(term in str(col).lower() for term in ['data', 'hora']):
                                non_null = df[col].notna().sum()
                                date_cols_info.append(f"{col}: {non_null} valores")
                        
                        if date_cols_info:
                            st.sidebar.info("Colunas de data encontradas:")
                            for info in date_cols_info:
                                st.sidebar.write(f"  • {info}")
                        
                        fig_calls = create_calls_chart(df, f"Horários de Ligações - {selected_sheet}")
                        if fig_calls:
                            st.plotly_chart(fig_calls, use_container_width=True)
                        else:
                            # Mostrar distribuição de outra forma
                            if 'TEL 1' in df.columns:
                                tel_counts = pd.DataFrame({
                                    'Status': ['Com Telefone 1', 'Sem Telefone 1'],
                                    'Quantidade': [
                                        df['TEL 1'].notna().sum(),
                                        df['TEL 1'].isna().sum()
                                    ]
                                })
                                
                                fig_tel = px.pie(
                                    tel_counts,
                                    values='Quantidade',
                                    names='Status',
                                    title=f'<b>Distribuição de Telefones 1 - {selected_sheet}</b>',
                                    color_discrete_sequence=['#2E86AB', '#A23B72']
                                )
                                fig_tel.update_traces(
                                    textposition='inside', 
                                    textinfo='percent+label'
                                )
                                fig_tel.update_layout(
                                    title=dict(x=0.5, xanchor='center')
                                )
                                st.plotly_chart(fig_tel, use_container_width=True)
                    else:
                        # Para outras planilhas, mostrar distribuição de telefones
                        if 'TEL 1' in df.columns and 'TEL 2' in df.columns:
                            tel_data = pd.DataFrame({
                                'Tipo': ['Com TEL 1', 'Com TEL 2', 'Com ambos', 'Sem telefone'],
                                'Quantidade': [
                                    (df['TEL 1'].notna() & df['TEL 2'].isna()).sum(),
                                    (df['TEL 2'].notna() & df['TEL 1'].isna()).sum(),
                                    (df['TEL 1'].notna() & df['TEL 2'].notna()).sum(),
                                    (df['TEL 1'].isna() & df['TEL 2'].isna()).sum()
                                ]
                            })
                            
                            fig_tel = px.bar(
                                tel_data,
                                x='Tipo',
                                y='Quantidade',
                                title=f'<b>Distribuição de Telefones - {selected_sheet}</b>',
                                color='Tipo',
                                color_discrete_sequence=px.colors.qualitative.Set2,
                                text='Quantidade'
                            )
                            fig_tel.update_traces(textposition='outside')
                            fig_tel.update_layout(
                                height=500,
                                title=dict(x=0.5, xanchor='center'),
                                showlegend=False
                            )
                            st.plotly_chart(fig_tel, use_container_width=True)
                        elif 'TEL 1' in df.columns:
                            # Apenas TEL 1 disponível
                            tel_counts = pd.DataFrame({
                                'Status': ['Com Telefone', 'Sem Telefone'],
                                'Quantidade': [
                                    df['TEL 1'].notna().sum(),
                                    df['TEL 1'].isna().sum()
                                ]
                            })
                            
                            fig_tel = px.pie(
                                tel_counts,
                                values='Quantidade',
                                names='Status',
                                title=f'<b>Distribuição de Telefones - {selected_sheet}</b>',
                                color_discrete_sequence=['#2E86AB', '#A23B72']
                            )
                            fig_tel.update_traces(
                                textposition='inside', 
                                textinfo='percent+label'
                            )
                            fig_tel.update_layout(
                                title=dict(x=0.5, xanchor='center')
                            )
                            st.plotly_chart(fig_tel, use_container_width=True)
                
                # Observações importantes
                show_important_observations(df, f"Observações Importantes - {selected_sheet}")
                
                # Tabela com dados brutos (opcional)
                with st.expander("📄 Ver dados completos da planilha"):
                    st.dataframe(df, use_container_width=True)
        
        else:
            st.error("Não foi possível carregar os dados do arquivo.")
        
        # Limpar arquivo temporário
        try:
            os.remove("temp_uploaded.xlsx")
        except:
            pass
    
    else:
        # Tela inicial sem arquivo
        st.info("👈 Por favor, carregue o arquivo Excel na barra lateral")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("📋 Estrutura Esperada do Arquivo")
            st.markdown("""
            O arquivo deve conter:
            1. **Planilha CARGAS_NITEROI:**
               - Colunas: CNPJ, RAZÃO SOCIAL, TEL 1, TEL 2, E-MAIL
               - Data/Hora 1, Data/Hora 2, Data/Hora 3
               - SITUAÇÃO, OBSERVAÇÃO
            
            2. **Outras planilhas:**
               - Colunas similares, sem datas
            """)
        
        with col2:
            st.subheader("🎯 Principais Funcionalidades")
            st.markdown("""
            ✅ **Normalização automática** das situações
            ✅ **Análise de horários** incluindo Data/Hora 3
            ✅ **Layout ajustado** sem sobreposição
            ✅ **Observações filtradas** (exceto "Não atende")
            ✅ **Métricas detalhadas** por planilha
            ✅ **Download dos dados** processados
            ✅ **Relatório HTML** para compartilhar
            """)
        
        st.markdown("---")
        
        # Instruções para compartilhar
        st.subheader("📤 Como Compartilhar com Seu Diretor")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.markdown("""
            ### 📄 **Relatório HTML**
            1. Carregue o arquivo Excel
            2. Clique em "Gerar Relatório HTML"
            3. Baixe o arquivo .html
            4. Envie por email
            """)
        
        with col2:
            st.markdown("""
            ### 📊 **Executável**
            1. Instale o PyInstaller:
               ```bash
               pip install pyinstaller
               ```
            2. Crie o executável:
               ```bash
               pyinstaller --onefile relatorio_cargas_niteroi.py
               ```
            """)
        
        with col3:
            st.markdown("""
            ### 🌐 **Online**
            1. Crie conta no Streamlit Cloud
            2. Suba o código para GitHub
            3. Conecte e compartilhe o link
            4. Acesse de qualquer lugar
            """)
        
        st.markdown("---")
        st.subheader("👥 Desenvolvido por:")
        st.markdown("**Kaynan Monteiro** e **David Florencio**")
        
        # Adicionar instruções rápidas
        with st.expander("⚡ Instruções Rápidas"):
            st.markdown("""
            1. **Para usar:** Carregue o arquivo Excel na barra lateral
            2. **Para análise:** Navegue entre as abas na barra lateral
            3. **Para exportar:** Use os botões na barra lateral para:
               - 📥 Dados processados em Excel
               - 📄 Relatório HTML para compartilhar
            4. **Para compartilhar:** Gere o HTML e envie por email
            """)

# Executar aplicação
if __name__ == "__main__":
    main()