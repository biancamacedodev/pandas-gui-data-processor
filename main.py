import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from tkinter import Label, TOP
from tkinter import Tk
from tkinter import Toplevel
from tkinter import Button, Entry
from tkinter import Menu
from tkinter import ttk

import numpy as np
import pandas as pd
from pandastable import Table  


class ExcelEditor:
    # cria uma classe ExcelEditor que herda da classe Tkinter (usa a janela principal como master)
    def __init__(self, master: tk.Tk) -> None:
        # inicializa a classe ExcelEditor com a janela principal
        self.master = master

        self.resultado_label = Label(
            self.master, text="Total: ", font=("Arial", 16), bg="#F5F5F5"
        )
        self.resultado_label.pack(side=TOP, padx=10, pady=10)

        # cria um dataframe vazio
        self.df = pd.DataFrame()

        # inicializa as variaveis de controle
        self.tree: ttk.Treeview | None = None
        self.table: Table | None = None
        self.filename: str = ""

        # cria os widgets da interface
        self.create_widgets()

        # programa o evento de duplo clique para passar os dados da treeview para tela de edição
        self.tree.bind("<Double-1>", self.editarItens)

    def create_widgets(self) -> None:
        # cria a janela de menu
        menu_bar = Menu(self.master)

        # Cria o menu "Arquivo"
        menu_arquivos = Menu(menu_bar, tearoff=0)

        # adiciona os itens de menu de arquivos
        menu_arquivos.add_command(label="Abrir", command=self.carregar_excel)
        menu_arquivos.add_separator()
        menu_arquivos.add_command(label="Salvar Como", command=self.salvar_como)
        menu_arquivos.add_separator()
        menu_arquivos.add_command(label="Sair", command=self.master.destroy)

        # adiciona o menu Arquivo à barra de menus
        menu_bar.add_cascade(label="Arquivo", menu=menu_arquivos)

        # --------------------------------------------------------

        # Cria o menu "Editar"
        # O tearoff=0 é uma configuração de menu que, quando definida como 0, desativa a função de arrastar
        menu_edicao = Menu(menu_bar, tearoff=0)

        menu_edicao.add_command(label="Renomear Coluna", command=self.renomear_coluna)
        menu_edicao.add_command(label="Remover Coluna", command=self.remover_coluna)
        menu_edicao.add_command(label="Filtrar", command=self.filtrar)
        menu_edicao.add_command(label="Pivot", command=self.master.destroy)
        menu_edicao.add_command(label="Group", command=self.group)
        menu_edicao.add_command(
            label="Remover linhas em branco", command=self.remover_linhas_em_branco
        )
        menu_edicao.add_command(
            label="Remover linhas alternadas", command=self.remove_linhas_alternadas
        )
        menu_edicao.add_command(
            label="Remover Duplicados", command=self.remover_duplicados
        )

        # Adiciona o menu Edição à barra de menus
        menu_bar.add_cascade(label="Editar", menu=menu_edicao)

        # --------------------------------------------------------

        # Cria o menu "Merge"
        merge_menu = Menu(menu_bar, tearoff=0)

        merge_menu.add_command(label="Inner Join", command=self.merge_inner_join)
        merge_menu.add_command(label="Join Full", command=self.merge_join_full)
        merge_menu.add_command(label="Left Join", command=self.merge_left_join)
        merge_menu.add_command(label="Merge Outer", command=self.merge_outer)

        # adiciona o menu Merge à barra de menus
        menu_bar.add_cascade(label="Merge", menu=merge_menu)

        # --------------------------------------------------------

        # Cria o menu "Relatórios"
        relatorio_menu = Menu(menu_bar, tearoff=0)

        relatorio_menu.add_command(
            label="Consolidar", command=self.consolidar_arquivos
        )
        relatorio_menu.add_command(label="Quebra", command=self.quebrar_arquivos)

        # adiciona o menu Relatório à barra de menus
        menu_bar.add_cascade(label="Relatório", menu=relatorio_menu)

        # define a barra de menu como a barra de menu principal
        self.master.config(menu=menu_bar)

        # criando a Treeview
        self.tree = ttk.Treeview(self.master)

        # coloca o widget de árvore na janela principal
        self.tree.pack(expand=False, fill="both")

    def soma_colunas_com_valor(self) -> None:
        resultados: list[str] = []

        for coluna in self.df.columns:
            if pd.api.types.is_numeric_dtype(self.df[coluna]):
                valores_numericos = self.df[coluna][0:]

                valores_numericos = pd.to_numeric(
                    valores_numericos, errors="coerce"
                )

                valores_numericos = valores_numericos[~np.isnan(valores_numericos)]

                soma = valores_numericos.sum()

                resultado = f"A soma da coluna {coluna} é {soma}"

                resultados.append(resultado)

        self.resultado_label.config(text="\n".join(resultados))

    def carregar_excel(self) -> None:
        # define os tipos de arquivos que podem ser abertos
        tipo_de_arquivo = (("Excel files", "*.xlsx;*.xls"), ("All files", "*.*"))

        # abre a janela para selecionar o arquivo que armazena o caminho na variável
        self.nome_do_arquivo = filedialog.askopenfilename(
            title="Selecione o arquivo", filetypes=tipo_de_arquivo
        )

        if not self.nome_do_arquivo:
            return

        try:
            self.df = pd.read_excel(self.nome_do_arquivo)
            self.atualiza_treeview()
        except Exception as e:  
            messagebox.showerror("Erro", f"Não foi possível abrir o arquivo: {e}")
            return

        # calcula soma das colunas com valores
        self.soma_colunas_com_valor()

    def atualiza_treeview(self) -> None:
        if self.tree is None:
            return

        # Apaga todo o conteúdo da treeview
        self.tree.delete(*self.tree.get_children())

        # Define as colunas da treeview com base nas colunas do df (Dataframe)
        self.tree["columns"] = list(self.df.columns)
        self.tree["show"] = "headings"

        for column in self.df.columns:
            # Define o texto do cabeçalho de cada coluna
            self.tree.heading(column, text=column)
            self.tree.column(column, anchor="center")

        # O método iterrows é usado para percorrer cada linha do DataFrame
        for _, row in self.df.iterrows():
            # Converte a linha do dataframe em uma lista e adiciona à treeview
            values = list(row)

            # Converte valores de tipo numpy para python nativo
            for j, value in enumerate(values):
                if isinstance(value, np.generic):
                    values[j] = value.item()

            self.tree.insert("", tk.END, values=values)

    def renomear_coluna(self) -> None:
        janela_renomear_coluna = Toplevel(self.master)
        janela_renomear_coluna.title("Renomear Coluna")

        # define a altura e largura da janela
        largura_janela = 400
        altura_janela = 250

        # obtém a largura e altura da tela do computador
        largura_tela = janela_renomear_coluna.winfo_screenwidth()
        altura_tela = janela_renomear_coluna.winfo_screenheight()

        # calcula a posição da janela para centralizar na tela
        posicao_x = (largura_tela // 2) - (largura_janela // 2)
        posicao_y = (altura_tela // 2) - (altura_janela // 2)

        # define a posição da janela
        janela_renomear_coluna.geometry(
            f"{largura_janela}x{altura_janela}+{posicao_x}+{posicao_y}"
        )

        janela_renomear_coluna.configure(bg="#FFFFFF")

        label_coluna = tk.Label(
            janela_renomear_coluna,
            text="Digite o nome da coluna que vai renomear:",
            font=("Arial", 12),
            bg="#FFFFFF",
        )
        label_coluna.pack(pady=10)
        entry_coluna = Entry(janela_renomear_coluna, font=("Arial", 12))
        entry_coluna.pack()

        label_novo_nome = tk.Label(
            janela_renomear_coluna,
            text="Digite o novo nome da coluna:",
            font=("Arial", 12),
            bg="#FFFFFF",
        )
        label_novo_nome.pack(pady=10)
        entry_novo_nome = Entry(janela_renomear_coluna, font=("Arial", 12))
        entry_novo_nome.pack()

        botao_renomear = Button(
            janela_renomear_coluna,
            text="Renomear",
            font=("Arial", 12),
            command=lambda: self.funcao_renomear_coluna(
                entry_coluna.get(), entry_novo_nome.get(), janela_renomear_coluna
            ),
        )
        botao_renomear.pack(pady=20)

    def funcao_renomear_coluna(
        self, column: str, novo_nome: str, janela_renomear_coluna: Toplevel
    ) -> None:
        if novo_nome and column in self.df.columns:
            self.df = self.df.rename(columns={column: novo_nome})
            self.atualiza_treeview()

        janela_renomear_coluna.destroy()

    def remover_linhas_em_branco(self) -> None:
        # Solicita ao usuário se ele quer mesmo remover as linhas em branco
        resposta = messagebox.askyesno(
            "Remover linhas em branco",
            "Tem certeza que deseja remover as linhas em branco?",
        )

        # Verifica se a resposta é "sim"
        if resposta:
            # Deleta as linhas com valores em branco
            self.df = self.df.dropna(axis=0)

            # Atualiza a árvore (treeview) com o conteúdo do arquivo
            self.atualiza_treeview()

            # Calcula a soma das colunas com valores
            self.soma_colunas_com_valor()

    def remove_linhas_alternadas(self) -> None:
        janela_remove_linhas_alternadas = Toplevel(self.master)
        janela_remove_linhas_alternadas.title("Remover Linhas Alternadas")

        # define a altura e largura da janela
        largura_janela = 400
        altura_janela = 250

        # obtém a largura e altura da tela do computador
        largura_tela = janela_remove_linhas_alternadas.winfo_screenwidth()
        altura_tela = janela_remove_linhas_alternadas.winfo_screenheight()

        # calcula a posição da janela para centralizar na tela
        posicao_x = (largura_tela // 2) - (largura_janela // 2)
        posicao_y = (altura_tela // 2) - (altura_janela // 2)

        # define a posição da janela
        janela_remove_linhas_alternadas.geometry(
            f"{largura_janela}x{altura_janela}+{posicao_x}+{posicao_y}"
        )

        janela_remove_linhas_alternadas.configure(bg="#FFFFFF")

        label_coluna = tk.Label(
            janela_remove_linhas_alternadas,
            text="Digite o número da primeira linha a ser removida:",
            font=("Arial", 12),
            bg="#FFFFFF",
        )
        label_coluna.pack(pady=10)
        entry_linha_inicio = Entry(janela_remove_linhas_alternadas, font=("Arial", 12))
        entry_linha_inicio.pack()

        label_linha_fim = tk.Label(
            janela_remove_linhas_alternadas,
            text="Digite o número da última linha a ser removida:",
            font=("Arial", 12),
            bg="#FFFFFF",
        )
        label_linha_fim.pack(pady=10)
        entry_linha_fim = Entry(janela_remove_linhas_alternadas, font=("Arial", 12))
        entry_linha_fim.pack()

        botao_remover = Button(
            janela_remove_linhas_alternadas,
            text="Remover",
            font=("Arial", 12),
            command=lambda: self.funcao_remove_linhas_alternadas(
                entry_linha_inicio.get(),
                entry_linha_fim.get(),
                janela_remove_linhas_alternadas,
            ),
        )
        botao_remover.pack(pady=20)

    def funcao_remove_linhas_alternadas(
        self, linha_inicio: str, linha_fim: str, janela_remove_linhas_alternadas: Toplevel
    ) -> None:
        try:
            primeira_linha = int(linha_inicio)
            ultima_linha = int(linha_fim)
        except ValueError:
            messagebox.showerror("Erro", "Informe números válidos para as linhas.")
            return

        # Deleta as linhas em um intervalo (índices baseados em 1 na interface)
        self.df = self.df.drop(self.df.index[primeira_linha - 1 : ultima_linha])

        # Atualiza a árvore (treeview) com o conteúdo do arquivo
        self.atualiza_treeview()

        # Calcula a soma das colunas com valores
        self.soma_colunas_com_valor()

        # Fecha a janela secundária
        janela_remove_linhas_alternadas.destroy()

    def remover_duplicados(self) -> None:
        janela_remover_duplicados = Toplevel(self.master)
        janela_remover_duplicados.title("Remover Duplicados")

        # define a altura e largura da janela
        largura_janela = 600
        altura_janela = 150

        # obtém a largura e altura da tela do computador
        largura_tela = janela_remover_duplicados.winfo_screenwidth()
        altura_tela = janela_remover_duplicados.winfo_screenheight()

        # calcula a posição da janela para centralizar na tela
        posicao_x = (largura_tela // 2) - (largura_janela // 2)
        posicao_y = (altura_tela // 2) - (altura_janela // 2)

        # define a posição da janela
        janela_remover_duplicados.geometry(
            f"{largura_janela}x{altura_janela}+{posicao_x}+{posicao_y}"
        )

        janela_remover_duplicados.configure(bg="#FFFFFF")

        label_coluna = tk.Label(
            janela_remover_duplicados,
            text="Digite o nome da coluna que vai remover os itens duplicados:",
            font=("Arial", 12),
            bg="#FFFFFF",
        )
        label_coluna.pack(pady=10)
        entry_coluna = Entry(janela_remover_duplicados, font=("Arial", 12))
        entry_coluna.pack()

        botao_remover = Button(
            janela_remover_duplicados,
            text="Remover",
            font=("Arial", 12),
            command=lambda: self.funcao_remover_duplicados(
                entry_coluna.get(), janela_remover_duplicados
            ),
        )
        botao_remover.pack(pady=20)

    def funcao_remover_duplicados(
        self, coluna: str, janela_remover_duplicados: Toplevel
    ) -> None:
        if coluna and coluna in self.df.columns:
            # deleta os itens duplicados, mantendo a primeira ocorrência
            self.df = self.df.drop_duplicates(subset=coluna, keep="first")

            # Atualiza a árvore (treeview) com o conteúdo do arquivo
            self.atualiza_treeview()

            # Calcula a soma das colunas com valores
            self.soma_colunas_com_valor()

        # Fecha a janela secundária
        janela_remover_duplicados.destroy()

    def remover_coluna(self) -> None:
        janela_remover_coluna = Toplevel(self.master)
        janela_remover_coluna.title("Remover Coluna")

        # define a altura e largura da janela
        largura_janela = 600
        altura_janela = 150

        # obtém a largura e altura da tela do computador
        largura_tela = janela_remover_coluna.winfo_screenwidth()
        altura_tela = janela_remover_coluna.winfo_screenheight()

        # calcula a posição da janela para centralizar na tela
        posicao_x = (largura_tela // 2) - (largura_janela // 2)
        posicao_y = (altura_tela // 2) - (altura_janela // 2)

        # define a posição da janela
        janela_remover_coluna.geometry(
            f"{largura_janela}x{altura_janela}+{posicao_x}+{posicao_y}"
        )

        janela_remover_coluna.configure(bg="#FFFFFF")

        label_coluna = tk.Label(
            janela_remover_coluna,
            text="Digite o nome da coluna que quer remover:",
            font=("Arial", 12),
            bg="#FFFFFF",
        )
        label_coluna.pack(pady=10)
        entry_coluna = Entry(janela_remover_coluna, font=("Arial", 12))
        entry_coluna.pack()

        botao_remover = Button(
            janela_remover_coluna,
            text="Remover",
            font=("Arial", 12),
            command=lambda: self.funcao_remover_coluna(
                entry_coluna.get(), janela_remover_coluna
            ),
        )
        botao_remover.pack(pady=20)

    def funcao_remover_coluna(
        self, coluna: str, janela_remover_coluna: Toplevel
    ) -> None:
        if coluna and coluna in self.df.columns:
            self.df = self.df.drop(columns=coluna)

            # Atualiza a árvore (treeview) com o conteúdo do arquivo
            self.atualiza_treeview()

            # Calcula a soma das colunas com valores
            self.soma_colunas_com_valor()

        # Fecha a janela secundária
        janela_remover_coluna.destroy()

    def filtrar(self) -> None:
        janela_filtrar = Toplevel(self.master)
        janela_filtrar.title("Filtrar")

        # define a altura e largura da janela
        largura_janela = 600
        altura_janela = 200

        # obtém a largura e altura da tela do computador
        largura_tela = janela_filtrar.winfo_screenwidth()
        altura_tela = janela_filtrar.winfo_screenheight()

        # calcula a posição da janela para centralizar na tela
        posicao_x = (largura_tela // 2) - (largura_janela // 2)
        posicao_y = (altura_tela // 2) - (altura_janela // 2)

        # define a posição da janela
        janela_filtrar.geometry(
            f"{largura_janela}x{altura_janela}+{posicao_x}+{posicao_y}"
        )

        janela_filtrar.configure(bg="#FFFFFF")

        label_coluna = tk.Label(
            janela_filtrar,
            text="Digite o nome da coluna que quer filtrar:",
            font=("Arial", 12),
            bg="#FFFFFF",
        )
        label_coluna.pack(pady=10)
        entry_coluna = Entry(janela_filtrar, font=("Arial", 12))
        entry_coluna.pack()

        label_valor = tk.Label(
            janela_filtrar,
            text="Digite o valor a ser filtrado:",
            font=("Arial", 12),
            bg="#FFFFFF",
        )
        label_valor.pack(pady=10)
        entry_valor = Entry(janela_filtrar, font=("Arial", 12))
        entry_valor.pack()

        botao_filtrar = Button(
            janela_filtrar,
            text="Filtrar",
            font=("Arial", 12),
            command=lambda: self.funcao_filtrar(
                entry_coluna.get(), entry_valor.get(), janela_filtrar
            ),
        )
        botao_filtrar.pack(pady=20)

    def funcao_filtrar(
        self, coluna: str, valor: str, janela_filtrar: Toplevel
    ) -> None:
        if coluna and valor and coluna in self.df.columns:
            # filtra o dataframe com base na coluna e valor
            self.df = self.df[self.df[coluna] == valor]

            # Atualiza a árvore (treeview) com o conteúdo do arquivo
            self.atualiza_treeview()

            # Calcula a soma das colunas com valores
            self.soma_colunas_com_valor()

        # Fecha a janela secundária
        janela_filtrar.destroy()

    def group(self) -> None:
        janela_group = Toplevel(self.master)
        janela_group.title("Agrupar")

        # define a altura e largura da janela
        largura_janela = 600
        altura_janela = 200

        # obtém a largura e altura da tela do computador
        largura_tela = janela_group.winfo_screenwidth()
        altura_tela = janela_group.winfo_screenheight()

        # calcula a posição da janela para centralizar na tela
        posicao_x = (largura_tela // 2) - (largura_janela // 2)
        posicao_y = (altura_tela // 2) - (altura_janela // 2)

        # define a posição da janela
        janela_group.geometry(
            f"{largura_janela}x{altura_janela}+{posicao_x}+{posicao_y}"
        )

        janela_group.configure(bg="#FFFFFF")

        label_coluna = tk.Label(
            janela_group,
            text="Digite o nome da coluna que quer agrupar:",
            font=("Arial", 12),
            bg="#FFFFFF",
        )
        label_coluna.pack(pady=10)
        entry_coluna = Entry(janela_group, font=("Arial", 12))
        entry_coluna.pack()

        botao_group = Button(
            janela_group,
            text="Agrupar",
            font=("Arial", 12),
            command=lambda: self.funcao_agrupar(entry_coluna.get(), janela_group),
        )
        botao_group.pack(pady=20)

    def funcao_agrupar(self, coluna: str, janela_group: Toplevel) -> None:
        if not coluna or coluna not in self.df.columns:
            janela_group.destroy()
            return

        if self.tree is None:
            janela_group.destroy()
            return

        # limpa os dados da treeview
        self.tree.delete(*self.tree.get_children())

        group_dados = self.df.groupby(coluna).sum(numeric_only=True)

        for i, linha in group_dados.iterrows():
            values = list(linha)
            for j, value in enumerate(values):
                if isinstance(value, np.generic):
                    values[j] = value.item()

            # inserindo nova linha na treeview
            self.tree.insert("", tk.END, values=[i] + values)

        # fecha a janela secundaria
        janela_group.destroy()

    def merge_inner_join(self) -> None:
        # define os tipos de arquivos que podem ser abertos
        tipo_de_arquivo = (("Excel files", "*.xlsx;*.xls"), ("All files", "*.*"))

        # abre a janela para selecionar o arquivo que armazena o caminho na variável
        nome_do_primeiro_arquivo = filedialog.askopenfilename(
            title="Selecione o primeiro arquivo", filetypes=tipo_de_arquivo
        )
        if not nome_do_primeiro_arquivo:
            return

        nome_do_segundo_arquivo = filedialog.askopenfilename(
            title="Selecione o segundo arquivo", filetypes=tipo_de_arquivo
        )
        if not nome_do_segundo_arquivo:
            return

        # ler os arquivos em formato excel
        primeiro_arquivo = pd.read_excel(nome_do_primeiro_arquivo)
        segundo_arquivo = pd.read_excel(nome_do_segundo_arquivo)

        # pergunta ao usuario qual coluna deve ser utilizada no merge
        coluna_join = simpledialog.askstring(
            "Coluna do Inner Join",
            "Digite o nome da coluna que será utilizada para o Inner Join: ",
        )
        if not coluna_join:
            return

        # realiza o merge mantendo apenas os vendedores presentes em ambos
        self.df = pd.merge(
            primeiro_arquivo,
            segundo_arquivo[["Vendedor", "Total Vendas"]],
            on=coluna_join,
            how="inner",
            suffixes=(" Loja 1", " Loja 2"),
        )

        # Resume os dados relevantes
        colunas_desejadas = [
            "Vendedor",
            "Total Vendas Loja 1",
            "Total Vendas Loja 2",
        ]
        self.df = self.df[colunas_desejadas]

        # Atualiza a treeview com o resultado do merge
        self.atualiza_treeview()

        # Calcula a soma das colunas com valores
        self.soma_colunas_com_valor()

    def merge_join_full(self) -> None:
        # define os tipos de arquivos que podem ser abertos
        tipo_de_arquivo = (("Excel files", "*.xlsx;*.xls"), ("All files", "*.*"))

        nome_do_primeiro_arquivo = filedialog.askopenfilename(
            title="Selecione o primeiro arquivo", filetypes=tipo_de_arquivo
        )
        if not nome_do_primeiro_arquivo:
            return

        nome_do_segundo_arquivo = filedialog.askopenfilename(
            title="Selecione o segundo arquivo", filetypes=tipo_de_arquivo
        )
        if not nome_do_segundo_arquivo:
            return

        primeiro_arquivo = pd.read_excel(nome_do_primeiro_arquivo)
        segundo_arquivo = pd.read_excel(nome_do_segundo_arquivo)

        # realiza o merge utilizando o tipo "full", unindo todos os registros dos dois arquivos
        self.df = pd.concat([primeiro_arquivo, segundo_arquivo], ignore_index=True)

        # Remove as linhas duplicadas com base no Id Vendedor, se existir
        if "Id Vendedor" in self.df.columns:
            self.df = self.df.drop_duplicates(subset="Id Vendedor")

        # Atualiza a treeview com o resultado do merge
        self.atualiza_treeview()

        # Calcula a soma das colunas com valores
        self.soma_colunas_com_valor()

    def merge_left_join(self) -> None:
        # define os tipos de arquivos que podem ser abertos
        tipo_de_arquivo = (("Excel files", "*.xlsx;*.xls"), ("All files", "*.*"))

        nome_do_primeiro_arquivo = filedialog.askopenfilename(
            title="Selecione o primeiro arquivo", filetypes=tipo_de_arquivo
        )
        if not nome_do_primeiro_arquivo:
            return

        nome_do_segundo_arquivo = filedialog.askopenfilename(
            title="Selecione o segundo arquivo", filetypes=tipo_de_arquivo
        )
        if not nome_do_segundo_arquivo:
            return

        primeiro_arquivo = pd.read_excel(nome_do_primeiro_arquivo)
        segundo_arquivo = pd.read_excel(nome_do_segundo_arquivo)

        # realiza o merge utilizando o tipo "left"
        self.df = pd.merge(
            primeiro_arquivo,
            segundo_arquivo,
            on=["Id Vendedor"],
            how="left",
            suffixes=(" Vendas", " Checagem"),
        )

        # remove as linhas em branco
        self.df = self.df.dropna()

        # Atualiza a treeview com o resultado do merge
        self.atualiza_treeview()

        # Calcula a soma das colunas com valores
        self.soma_colunas_com_valor()

    def merge_outer(self) -> None:
        # define os tipos de arquivos que podem ser abertos
        tipo_de_arquivo = (("Excel files", "*.xlsx;*.xls"), ("All files", "*.*"))

        nome_do_primeiro_arquivo = filedialog.askopenfilename(
            title="Selecione o primeiro arquivo", filetypes=tipo_de_arquivo
        )
        if not nome_do_primeiro_arquivo:
            return

        nome_do_segundo_arquivo = filedialog.askopenfilename(
            title="Selecione o segundo arquivo", filetypes=tipo_de_arquivo
        )
        if not nome_do_segundo_arquivo:
            return

        primeiro_arquivo = pd.read_excel(nome_do_primeiro_arquivo)
        segundo_arquivo = pd.read_excel(nome_do_segundo_arquivo)

        # realiza o merge utilizando o tipo "outer"
        self.df = pd.merge(
            primeiro_arquivo,
            segundo_arquivo,
            on=["Id Vendedor"],
            how="outer",
            suffixes=(" Loja 1", " Loja 2"),
        )

        # remove as linhas em branco
        self.df = self.df.dropna()

        # remove coluna redundante, se existir
        if "Vendedor Loja 2" in self.df.columns:
            del self.df["Vendedor Loja 2"]

        # Atualiza a treeview com o resultado do merge
        self.atualiza_treeview()

        # Calcula a soma das colunas com valores
        self.soma_colunas_com_valor()

    def salvar_como(self) -> None:
        # define os tipos de arquivos que podem ser salvos
        tipo_de_arquivo = (("Excel files", "*.xlsx;*.xls"), ("All files", "*.*"))

        # abre a janela para selecionar o local de salvamento e nome do arquivo
        nome_arquivo = filedialog.asksaveasfilename(
            title="Salvar como",
            filetypes=tipo_de_arquivo,
            defaultextension=".xlsx",
        )

        if nome_arquivo:
            try:
                self.df.to_excel(nome_arquivo, index=False)
            except Exception as e:  # noqa: BLE001
                messagebox.showerror(
                    "Erro", f"Não foi possível salvar o arquivo: {e}"
                )

    def editarItens(self, event) -> None:  # noqa: N802, D401
        """Callback de duplo clique para editar um item da treeview."""
        if self.tree is None:
            return

        selecionados = self.tree.selection()
        if not selecionados:
            return

        item_id = selecionados[0]

        # obtém os valores do item selecionado
        values = self.tree.item(item_id, "values")

        # cria a janela de edicao
        janela_edicao = Toplevel(self.master)
        janela_edicao.title("Editar linha")

        # adiciona campos de edicao para cada coluna da tabela
        for linha, nome_coluna in enumerate(self.df.columns):
            label = tk.Label(janela_edicao, text=nome_coluna, font=("Arial", 12))
            label.grid(row=linha, column=0, padx=5, pady=5, sticky="w")

            entry = Entry(janela_edicao, font=("Arial", 12))
            entry.insert(0, values[linha])
            entry.grid(row=linha, column=1, padx=5, pady=5, sticky="we")

        janela_edicao.columnconfigure(1, weight=1)

        # botão de salvar mudanças
        salvar = Button(
            janela_edicao,
            text="Salvar",
            font=("Arial", 12),
            command=lambda: self.salvar_alteracoes(item_id, janela_edicao),
        )

        salvar.grid(row=len(self.df.columns), column=0, columnspan=2, pady=10)

    def salvar_alteracoes(self, item_id: str, janela_edicao: Toplevel) -> None:
        if self.tree is None:
            janela_edicao.destroy()
            return

        # obtém novos valores inseridos na janela de edicao
        novos_valores: list[str] = [
            child.get()
            for child in janela_edicao.winfo_children()
            if isinstance(child, Entry)
        ]

        # encontra o índice da linha no DataFrame correspondente ao item da treeview
        row_index = self.tree.index(item_id)

        # atualiza df com novos valores
        if 0 <= row_index < len(self.df):
            self.df.iloc[row_index, :] = novos_valores

        # atualiza a árvore com os novos valores
        self.tree.item(item_id, values=novos_valores)

        janela_edicao.destroy()

    def consolidar_arquivos(self) -> None:
        try:
            # seleciona a pasta onde tem os arquivos
            caminho_arquivos = filedialog.askdirectory(title="Selecione a pasta")
            if not caminho_arquivos:
                return

            # lista todos os arquivos da pasta
            lista_arquivos = os.listdir(caminho_arquivos)

            # pega o caminho + nome e verifica os 4 últimos dígitos de cada arquivo
            lista_caminho_e_arquivos_excel = [
                os.path.join(caminho_arquivos, arquivo)
                for arquivo in lista_arquivos
                if arquivo.lower().endswith("xlsx")
            ]

            # cria o df
            dados_arquivo = pd.DataFrame()

            # copia todos os dados dos arquivos para variável
            for arquivo in lista_caminho_e_arquivos_excel:
                dados = pd.read_excel(arquivo)
                dados_arquivo = pd.concat([dados_arquivo, dados], ignore_index=True)

            # passa os dados para df
            self.df = dados_arquivo

        except Exception as e:  # noqa: BLE001
            messagebox.showerror("Erro", f"Não foi possível abrir os arquivos: {e}")
            return

        # Atualiza a treeview com o resultado
        self.atualiza_treeview()

        # Calcula a soma das colunas com valores
        self.soma_colunas_com_valor()

    def quebrar_arquivos(self) -> None:
        # seleciona a coluna usada para quebrar
        coluna = simpledialog.askstring(
            "Separar Arquivos", "Informe a coluna que deseja quebrar o arquivo: "
        )

        # Verifica se uma coluna foi selecionada
        if not coluna or coluna not in self.df.columns:
            return

        # agrupa o dataframe por valores únicos na coluna selecionada
        grupos = self.df.groupby(coluna)

        # pede ao usuário que selecione uma pasta para salvar os arquivos quebrados
        pasta_destino = filedialog.askdirectory(
            title="Selecione a pasta que deseja salvar os arquivos"
        )

        # verifica se o usuário selecionou a pasta
        if not pasta_destino:
            return

        # itera pelos grupos e salva cada grupo em um arquivo separado
        for valor, grupo in grupos:
            # remove caracteres inválidos do nome do arquivo (regex)
            nome_arquivo = re.sub(
                r"[^\w\-_.]", "_", f"{coluna}_{valor}.xlsx"
            )

            # cria o caminho completo para o arquivo
            caminho_arquivo = os.path.join(pasta_destino, nome_arquivo)

            # salva o grupo em um arquivo excel
            grupo.to_excel(caminho_arquivo, index=False)

        # exibe mensagem de sucesso
        messagebox.showinfo("Concluído", "Relatórios criados com sucesso!")


def main() -> None:
    janela = Tk()
    janela.title("Editor de Excel com Pandas")
    janela.attributes("-fullscreen", False)

    ExcelEditor(janela)
    janela.mainloop()


if __name__ == "__main__":
    main()

