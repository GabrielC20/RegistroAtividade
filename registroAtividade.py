import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import tkinter as tk
from tkinter import ttk

# Função para adicionar valores nas barras do gráfico
def adicionar_valores_barras(ax):
    for p in ax.patches:
        total_minutos = p.get_height()
        horas, minutos = divmod(total_minutos, 60)  
        ax.annotate(f'{int(horas)}h {int(minutos)}min',
                    (p.get_x() + p.get_width() / 2., p.get_height()),
                    ha='center', va='center', xytext=(0, 10), textcoords='offset points')
        
# Função para exibir os gráficos
def exibir_graficos():
    mes = int(entry_mes.get())
    ano = int(entry_ano.get())
    df_filtrado = df[(df['Data'].dt.month == mes) & (df['Data'].dt.year == ano)]

    # Contagem de atividades realizadas no mês
    gabriel_count = df_filtrado[df_filtrado['Nome'].str.upper() == 'GABRIEL'].shape[0]
    uilliam_count = df_filtrado[df_filtrado['Nome'].str.upper() == 'UILLIAM'].shape[0]

    # Filtrar e somar atividades de Gabriel e Uilliam
    gabriel_atividades = filtrar_por_nome(df_filtrado, 'GABRIEL').groupby('Data')['Horas_Gastas'].sum().reset_index()
    uilliam_atividades = filtrar_por_nome(df_filtrado, 'UILLIAM').groupby('Data')['Horas_Gastas'].sum().reset_index()

    # Gráficos
    plt.figure(figsize=(16, 12))

    # Gráfico para Gabriel
    plt.subplot(2, 2, 1)
    ax1 = sns.barplot(x='Data', y='Horas_Gastas', data=gabriel_atividades, color='blue')
    plt.title('Horas de atividade Gabriel')
    plt.xlabel('Data')
    plt.ylabel('Total de Tempo Gasto')
    plt.xticks(rotation=45)
    plt.legend(['Gabriel'], loc='upper right')
    adicionar_valores_barras(ax1)

    # Gráfico para Uilliam
    plt.subplot(2, 2, 2)
    ax2 = sns.barplot(x='Data', y='Horas_Gastas', data=uilliam_atividades, color='orange')
    plt.title('Horas de atividade Uilliam')
    plt.xlabel('Data')
    plt.ylabel('Total de Tempo Gasto')
    plt.xticks(rotation=45)
    plt.legend(['Uilliam'], loc='upper right')
    adicionar_valores_barras(ax2)

    # Gráfico em pizza por mês (à esquerda)
    plt.subplot(2, 2, 3)
    labels = ['Gabriel', 'Uilliam']
    sizes = [gabriel_count, uilliam_count]
    plt.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=140, colors=['blue', 'orange'])
    plt.title(f'Quantidade de Atividades por Mês - {mes}/{ano}')
    plt.axis('equal') 

    # Gráfico de barras para a quantidade de atividades (à direita)
    plt.subplot(2, 2, 4)
    atividades = pd.DataFrame({
        'Nome': ['Gabriel', 'Uilliam'],
        'Quantidade': [gabriel_count, uilliam_count]
    })
    ax4 = sns.barplot(x='Nome', y='Quantidade', data=atividades, palette=['blue', 'orange'])
    plt.title(f'Quantidade de Atividades Realizadas - {mes}/{ano}')
    plt.xlabel('Nome')
    plt.ylabel('Quantidade de Atividades')
    for p in ax4.patches:
        ax4.annotate(f'{int(p.get_height())}',
                     (p.get_x() + p.get_width() / 2., p.get_height()),
                     ha='center', va='center', xytext=(0, 10), textcoords='offset points')

    # Ajustar layout
    plt.tight_layout(h_pad=3, w_pad=3)

    # Maximizar janela
    mng = plt.get_current_fig_manager()
    mng.window.state('zoomed')

    plt.show()

# Função para filtrar atividades pelo nome
def filtrar_por_nome(df, nome):
    return df[df['Nome'].str.upper() == nome.upper()]

# Interface gráfica
root = tk.Tk()
root.title('Dashboard de Atividades')

# Label e campo de entrada para o mês
label_mes = ttk.Label(root, text="Digite o número do mês (1-12):")
label_mes.pack(pady=10)

entry_mes = ttk.Entry(root)
entry_mes.pack(pady=5)

# Label e campo de entrada para o ano
label_ano = ttk.Label(root, text="Digite o ano (ex: 2024):")
label_ano.pack(pady=10)

entry_ano = ttk.Entry(root)
entry_ano.pack(pady=5)

# Botão para gerar gráficos
btn_exibir = ttk.Button(root, text="Exibir Gráficos", command=exibir_graficos)
btn_exibir.pack(pady=20)

# Carregar os dados
df = pd.read_excel('../registro-de-atividades.xlsx', engine='openpyxl', header=1)
df.columns = df.columns.str.strip()
df['Data'] = pd.to_datetime(df['Data'], format='%d/%m/%Y', errors='coerce')
df['Inicio'] = pd.to_datetime(df['Data'].dt.strftime('%Y-%m-%d') + ' ' + df['Inicio'].astype(str), format='%Y-%m-%d %H:%M:%S', errors='coerce')
df['Fim'] = pd.to_datetime(df['Data'].dt.strftime('%Y-%m-%d') + ' ' + df['Fim'].astype(str), format='%Y-%m-%d %H:%M:%S', errors='coerce')
df['Horas_Gastas'] = (df['Fim'] - df['Inicio']).dt.total_seconds() / 60  

# Iniciar a interface
root.mainloop()
