import pandas as pd
import tkinter as tk
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from tkinter import messagebox
from matplotlib.ticker import PercentFormatter
from tkinter import ttk, filedialog
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg


def limpar_dados():
    global df, current_graph, fig_pareto, fig_abc, canvas_widget_pareto, canvas_widget_abc, text_box, selected_item_var, selected_item_dropdown

    df = None

    if 'text_box' in globals():
        text_box.config(state=tk.NORMAL)
        text_box.delete(1.0, tk.END)
        text_box.config(state=tk.DISABLED)

    selected_item_var.set('')
    if 'selected_item_dropdown' in globals():
        selected_item_dropdown['values'] = ()

    if canvas_widget_pareto is not None:
        canvas_widget_pareto.grid_forget()
    fig_pareto = None

    if canvas_widget_abc is not None:
        canvas_widget_abc.grid_forget()
    fig_abc = None

    current_graph = None
    coluna_numeros_combobox['values'] = ()
    coluna_nomes_combobox['values'] = ()
    coluna_numeros_var.set('')
    coluna_nomes_var.set('')
    colunas_var.set('')
    colunas_listbox.config(height=0)
    selected_item_var.set('')


df = None
current_graph = None
fig_pareto = None
fig_abc = None


def is_dataframe_valid():
    global df
    return df is not None


def carregar_arquivo():
    global df
    if df is not None:
        return
    file_path = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")])

    if file_path:
        try:
            df = pd.read_excel(file_path, engine='openpyxl')
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao carregar o arquivo Excel: {e}")
            return

        if not df.empty:
            colunas_var.set(tuple(df.columns))
            colunas_listbox.config(height=min(20, len(df.columns)))
            coluna_numeros_combobox['values'] = tuple(eval(colunas_var.get()))
            coluna_nomes_combobox['values'] = tuple(eval(colunas_var.get()))
        else:
            messagebox.showerror("Erro", "O arquivo Excel selecionado está vazio. Por favor, selecione outro arquivo.")


def highlight_selected_item():
    global current_graph
    selected_item = selected_item_var.get()

    if selected_item:
        if current_graph == "pareto":
            gerar_grafico_pareto(highlight_item=selected_item)
        elif current_graph == "abc":
            gerar_grafico_abc(highlight_item=selected_item)


def gerar_grafico_pareto(highlight_item=None):
    if highlight_item is None:
        highlight_item = selected_item_var.get()
    global df, canvas_widget_pareto, fig_pareto, current_graph
    if df is None or coluna_numeros_var.get() == "" or coluna_nomes_var.get() == "":
        return

    current_graph = "pareto"

    if canvas_widget_pareto is not None:
        canvas_widget_pareto.grid_forget()

    coluna_numeros = coluna_numeros_var.get()
    coluna_nomes = coluna_nomes_var.get()

    dados = df[[coluna_numeros, coluna_nomes]]
    dados_ordenados = dados.sort_values(by=coluna_numeros, ascending=False)
    dados_ordenados['PorcentagemAcumulada'] = dados_ordenados[coluna_numeros].cumsum() / dados_ordenados[coluna_numeros].sum()

    dados_ordenados['Índice:'] = range(1, len(dados_ordenados) + 1)
    dados_ordenados.set_index('Índice:', inplace=True)

    fig, ax1 = plt.subplots(figsize=(12, 6))

    fig_pareto = fig

    bars = ax1.bar(dados_ordenados.index, dados_ordenados[coluna_numeros])

    if highlight_item:
        item_index = dados_ordenados[dados_ordenados[coluna_nomes] == highlight_item].index[0]
        bars[item_index - 1].set_color('g')
        legend_item = mpatches.Patch(color='g', label=f"{highlight_item}: Índice {item_index}")
        ax1.legend(handles=[legend_item], loc='upper left')

    ax1.set_xlabel('Índice')
    ax1.set_ylabel(coluna_numeros)

    ax2 = ax1.twinx()
    ax2.plot(dados_ordenados.index, dados_ordenados['PorcentagemAcumulada'], color='r', marker='D', ms=2)
    ax2.yaxis.set_major_formatter(PercentFormatter(1))
    ax2.yaxis.set_major_locator(plt.MaxNLocator(10))
    ax2.set_ylabel('Porcentagem acumulada')

    plt.xticks(rotation=45, ha='right')
    plt.title('Diagrama de Pareto')
    plt.tight_layout()

    canvas = FigureCanvasTkAgg(fig, master=app)
    canvas.draw_idle()
    canvas_widget_pareto = canvas.get_tk_widget()
    canvas_widget_pareto.grid(row=0, column=1, rowspan=row_offset, sticky=tk.W + tk.E + tk.N + tk.S)
    app.columnconfigure(1, weight=1)
    app.rowconfigure(0, weight=1)

    update_text_and_dropdown(dados_ordenados, coluna_nomes)


def remover_grafico_pareto():
    global canvas_widget_pareto
    if canvas_widget_pareto is not None:
        canvas_widget_pareto.grid_forget()


def gerar_grafico_abc(highlight_item=None):
    if highlight_item is None:
        highlight_item = selected_item_var.get()
    global df, canvas_widget_abc, fig_abc, current_graph
    if df is None or coluna_numeros_var.get() == "" or coluna_nomes_var.get() == "":
        return

    current_graph = "abc"

    if canvas_widget_abc is not None:
        canvas_widget_abc.grid_forget()

    coluna_numeros = coluna_numeros_var.get()
    coluna_nomes = coluna_nomes_var.get()

    dados = df[[coluna_numeros, coluna_nomes]]
    dados_ordenados = dados.sort_values(by=coluna_numeros, ascending=False)
    dados_ordenados['Porcentagem'] = dados_ordenados[coluna_numeros] / dados_ordenados[coluna_numeros].sum()
    dados_ordenados['PorcentagemAcumulada'] = dados_ordenados['Porcentagem'].cumsum()

    dados_ordenados['Índice:'] = range(1, len(dados_ordenados) + 1)
    dados_ordenados.set_index('Índice:', inplace=True)

    fig, ax = plt.subplots(figsize=(12, 6))

    fig_abc = fig

    bars = ax.bar(dados_ordenados.index, dados_ordenados['Porcentagem'])

    # Determine the limits for A, B, and C regions
    a_limit = 0.8
    b_limit = 0.95

    # Get the indices for A, B, and C regions
    a_region_indices = dados_ordenados[dados_ordenados['PorcentagemAcumulada'] <= a_limit].index
    b_region_indices = dados_ordenados[(dados_ordenados['PorcentagemAcumulada'] > a_limit) &
                                       (dados_ordenados['PorcentagemAcumulada'] <= b_limit)].index
    c_region_indices = dados_ordenados[dados_ordenados['PorcentagemAcumulada'] > b_limit].index

    # Color the bars based on their region
    for index in a_region_indices:
        bars[index - 1].set_color('C0')
    for index in b_region_indices:
        bars[index - 1].set_color('C1')
    for index in c_region_indices:
        bars[index - 1].set_color('C2')

    # Create legend for A, B, and C regions
    a_patch = mpatches.Patch(color='C0', label='Região A')
    b_patch = mpatches.Patch(color='C1', label='Região B')
    c_patch = mpatches.Patch(color='C2', label='Região C')
    ax.legend(handles=[a_patch, b_patch, c_patch], loc='upper left')

    if highlight_item:
        item_index = dados_ordenados[dados_ordenados[coluna_nomes] == highlight_item].index[0]
        bars[item_index - 1].set_color('g')
        legend_item = mpatches.Patch(color='g', label=f"{highlight_item}: Índice {item_index}")
        ax.legend(handles=[legend_item], loc='upper left')

    ax.set_xlabel('Índice')
    ax.set_ylabel(coluna_numeros)

    ax.plot(dados_ordenados.index, dados_ordenados['PorcentagemAcumulada'], color='r', marker='D', ms=2)
    ax.yaxis.set_major_formatter(PercentFormatter(1))
    ax.yaxis.set_major_locator(plt.MaxNLocator(10))
    ax.set_ylabel('Porcentagem Acumulada')

    plt.xticks(rotation=0, ha='right')
    plt.title('Gráfico da Curva ABC')
    plt.tight_layout()

    canvas = FigureCanvasTkAgg(fig, master=app)
    canvas.draw()
    canvas_widget_abc = canvas.get_tk_widget()
    canvas_widget_abc.grid(row=0, column=2, rowspan=row_offset, sticky=tk.W + tk.E + tk.N + tk.S)
    app.columnconfigure(2, weight=1)
    app.rowconfigure(0, weight=1)

    update_text_and_dropdown(dados_ordenados, coluna_nomes)


def remover_grafico_abc():
    global canvas_widget_abc
    if canvas_widget_abc is not None:
        canvas_widget_abc.grid_forget()


def update_text_and_dropdown(dados_ordenados, coluna_nomes):
    global text_box, selected_item_var, selected_item_dropdown
    ttk.Label(frame, text='Nome do Rótulo com base no Índice do Gráfico:').grid(row=9, column=0, columnspan=2, sticky=tk.W)
    text_box = tk.Text(frame, wrap=tk.WORD, height=10, width=50, font=("Arial", 8))
    text_box.grid(row=10, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W)

    index_name_mapping = dados_ordenados[coluna_nomes].to_string()
    text_box.delete(1.0, tk.END)
    text_box.insert(tk.END, index_name_mapping)
    text_box.config(state=tk.DISABLED)

    selected_item_var.set('')
    ttk.Label(frame, text='Escolha o Rótulo para referenciar no Gráfico:').grid(row=11, column=0, columnspan=2, sticky=tk.W)
    selected_item_dropdown = ttk.Combobox(frame, textvariable=selected_item_var,
                                          values=tuple(dados_ordenados[coluna_nomes]), state='readonly')
    selected_item_dropdown.grid(row=12, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W + tk.E + tk.N)
    selected_item_dropdown.bind('<<ComboboxSelected>>', lambda event: highlight_selected_item())


def salvar_grafico_pareto():
    global fig_pareto
    if fig_pareto is None:
        messagebox.showwarning("Aviso", "Nenhum gráfico de Pareto gerado.")
        return
    export_path_pareto = filedialog.asksaveasfilename(defaultextension='.png',
                                                      filetypes=[("Arquivos de imagem", "*.png")],
                                                      title="Salvar Diagrama de Pareto")
    if export_path_pareto:
        fig_pareto.savefig(export_path_pareto, dpi=1200)


def salvar_grafico_abc():
    global fig_abc
    if fig_abc is None:
        messagebox.showwarning("Aviso", "Nenhum gráfico da Curva ABC gerado.")
        return
    export_path_abc = filedialog.asksaveasfilename(defaultextension='.png',
                                                   filetypes=[("Arquivos de imagem", "*.png")],
                                                   title="Salvar Gráfico da Curva ABC")
    if export_path_abc:
        fig_abc.savefig(export_path_abc, dpi=1200)


canvas = None


def on_resize(event):
    global canvas, fig_pareto, fig_abc

    if canvas is not None:
        w, h = event.width, event.height
        if fig_pareto is not None:
            fig_pareto.set_size_inches(w / fig_pareto.dpi, h / fig_pareto.dpi)
        else:
            fig_pareto = plt.figure()
        if fig_abc is not None:
            fig_abc.set_size_inches(w / fig_abc.dpi, h / fig_abc.dpi)
        else:
            fig_abc = plt.figure()

        if current_graph == "pareto":
            w, h = event.width, event.height
            fig_pareto.set_size_inches(w / fig_pareto.dpi, h / fig_pareto.dpi)
            fig_pareto.tight_layout()
            canvas.draw_idle()

        if current_graph == "abc":
            w, h = event.width, event.height
            fig_abc.set_size_inches(w / fig_abc.dpi, h / fig_abc.dpi)
            fig_abc.tight_layout()
            canvas.draw_idle()


app = tk.Tk()

app.title('Diagrama de Pareto e Curva ABC')

selected_item_var = tk.StringVar()

frame = ttk.Frame(app, padding='10')
frame.grid(row=0, column=0, sticky=tk.W + tk.E + tk.N + tk.S)

frame.grid_columnconfigure(0, weight=1)
frame.grid_rowconfigure(12, weight=1)

app.columnconfigure(1, weight=1)
app.rowconfigure(1, weight=1)

canvas_widget_pareto = None
canvas_widget_abc = None

row_offset = 0

ttk.Button(frame, text='Carregar arquivo',
           command=carregar_arquivo).grid(row=row_offset, column=0, sticky=tk.W + tk.N)

row_offset = 1

colunas_var = tk.StringVar(value="")
colunas_listbox = tk.Listbox(frame, listvariable=colunas_var,
                             exportselection=False, width=0, height=0, font=("Arial", 8))
colunas_listbox.grid(row=row_offset, column=0, columnspan=2, sticky=tk.W + tk.E + tk.N)

row_offset += 1

ttk.Label(frame, text='Coluna referente aos valores numéricos:').grid(row=row_offset, column=0, sticky=tk.W + tk.N)

row_offset += 1

coluna_numeros_var = tk.StringVar(value=None)
coluna_nomes_var = tk.StringVar(value=None)
coluna_numeros_combobox = ttk.Combobox(frame, textvariable=coluna_numeros_var, values=tuple(colunas_var.get()),
                                       state='readonly')
coluna_numeros_combobox.grid(row=row_offset, column=0, columnspan=2, sticky=tk.W + tk.E + tk.N)

row_offset += 1

ttk.Label(frame, text='Coluna referente a identificação dos ítens:').grid(row=row_offset, column=0, sticky=tk.W + tk.N)

row_offset += 1

coluna_nomes_combobox = ttk.Combobox(frame, textvariable=coluna_nomes_var, values=tuple(colunas_var.get()),
                                     state='readonly')
coluna_nomes_combobox.grid(row=row_offset, column=0, columnspan=2, sticky=tk.E + tk.W + tk.N)

row_offset += 1

ttk.Button(frame, text='Gerar gráfico Pareto', command=lambda: [remover_grafico_abc(),
                                                                gerar_grafico_pareto()]).grid(row=row_offset,
                                                                                              column=0, pady='10',
                                                                                              sticky=tk.W + tk.N)
ttk.Button(frame, text='Salvar gráfico Pareto',
           command=salvar_grafico_pareto).grid(row=row_offset, column=1, padx=(10, 0), pady='10', sticky=tk.E + tk.N)

row_offset += 1

ttk.Button(frame, text='Gerar gráfico ABC', command=lambda: [remover_grafico_pareto(),
                                                             gerar_grafico_abc()]).grid(row=row_offset,
                                                                                        column=0, pady='10',
                                                                                        sticky=tk.W + tk.N)
ttk.Button(frame, text='Salvar gráfico ABC',
           command=salvar_grafico_abc).grid(row=row_offset, column=1, padx=(20, 0), pady='10', sticky=tk.E + tk.N)

row_offset += 1

ttk.Button(frame, text='Limpar dados e gráficos', command=limpar_dados).grid(row=row_offset, column=0, pady='10',
                                                                             sticky=tk.W + tk.N)

for child in frame.winfo_children():
    child.grid_configure(padx=5, pady=5)

app.bind('<Configure>', on_resize)

app.update_idletasks()

app.mainloop()

"""
O programa é uma aplicação em Python que utiliza o módulo ‘tkinter’ para criar uma interface gráfica do usuário (GUI) que permite carregar um arquivo Excel, selecionar colunas e gerar gráficos de Pareto e Curva ABC. A biblioteca ‘pandas’ é utilizada para manipular os dados e a biblioteca ‘matplotlib’ para criar e exibir os gráficos.

Vamos analisar cada função e suas respectivas operações:

1. limpar_dados(): Essa função limpa os dados e os elementos da GUI relacionados aos gráficos e ao DataFrame, removendo os gráficos e redefinindo variáveis globais.
2. is_dataframe_valid(): Retorna True se o DataFrame df não for None, indicando que os dados foram carregados.
3. carregar_arquivo(): Permite ao usuário selecionar um arquivo Excel e carregá-lo como um DataFrame do pandas. Se o arquivo for carregado com sucesso, atualiza a lista de colunas na interface.
4. highlight_selected_item(): Destaca o item selecionado no gráfico atual (Pareto ou ABC).
5. gerar_grafico_pareto(highlight_item=None): Gera o gráfico de Pareto com base nas colunas selecionadas e destaca o item especificado.
6. remover_grafico_pareto(): Remove o gráfico de Pareto da GUI.
7. gerar_grafico_abc(highlight_item=None): Gera o gráfico da Curva ABC com base nas colunas selecionadas e destaca o item especificado.
8. remover_grafico_abc(): Remove o gráfico da Curva ABC da GUI.
9. update_text_and_dropdown(dados_ordenados, coluna_nomes): Atualiza a caixa de texto e a lista suspensa com base nos dados ordenados e na coluna de nomes.
10. salvar_grafico_pareto(): Permite ao usuário salvar o gráfico de Pareto como uma imagem.
11. salvar_grafico_abc(): Permite ao usuário salvar o gráfico da Curva ABC como uma imagem.
12. on_resize(event): Redimensiona os gráficos quando a janela é redimensionada.

O programa também define variáveis globais e configura a interface gráfica do usuário (GUI) com botões, caixas de texto e listas suspensas para interagir com o usuário.
Os laços (loops) e condicionais no código são usados principalmente para manipular os dados e atualizar a GUI. Por exemplo, o laço for em gerar_grafico_abc() é usado para colorir as barras com base na região (A, B ou C) a que pertencem. Os condicionais, como "if df is None", são usados para verificar se os dados foram carregados antes de executar determinadas ações.
Em resumo, o programa é uma aplicação que permite ao usuário carregar um arquivo Excel, selecionar colunas e gerar gráficos de Pareto e Curva ABC, com a opção de destacar itens específicos nos gráficos e salvá-los como imagens.
"""
"""
Neste projeto, foram utilizadas as seguintes bibliotecas:

- `pandas`: Uma biblioteca de software criada para manipulação e análise de dados. Em particular, oferece estruturas de dados e operações para manipular tabelas numéricas e séries temporais. Agradecemos à comunidade `pandas` por desenvolver e manter esta ferramenta essencial.

- `tkinter`: É a biblioteca padrão de interface gráfica do usuário (GUI) para Python. É uma interface de alto nível para Tk. Agradecemos aos mantenedores do `tkinter` por fornecerem uma maneira simples e eficaz de criar interfaces de usuário.

- `matplotlib`: Uma biblioteca de plotagem 2D em Python que produz figuras de qualidade em uma variedade de formatos impressos e ambientes interativos. Agradecemos à comunidade `matplotlib` por fornecer uma ferramenta tão versátil e poderosa para visualização de dados.

- `matplotlib.patches`: Um módulo para trabalhar com patches. Um Patch é uma forma 2D com uma face colorida e uma borda. Agradecemos aos desenvolvedores do `matplotlib.patches` por fornecerem esta ferramenta útil para personalizar nossas visualizações.

- `openpyxl`: Uma biblioteca Python para ler/escrever arquivos Excel 2010 xlsx/xlsm/xltx/xltm. Agradecemos aos desenvolvedores do `openpyxl` por permitir a interação eficiente com os arquivos do Excel.

Por fim, gostaríamos de expressar nossa gratidão a todas as comunidades de desenvolvedores que contribuíram para essas bibliotecas. Seu trabalho duro e dedicação tornam projetos como este possíveis.
"""
