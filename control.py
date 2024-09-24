import tkinter as tk
from tkinter import messagebox, simpledialog, filedialog
import pandas as pd
import os


class EstoqueApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Controle de Estoque")
        self.estoque = {}  # Dicionário que armazenará {produto: {quantidade, categoria, marca, valor_fabrica, valor_loja}}

        # Configuração da grid responsiva
        root.grid_rowconfigure(7, weight=1)
        root.grid_columnconfigure(1, weight=1)

        # Labels e Entradas
        self.label_categoria = tk.Label(root, text="Categoria")
        self.label_categoria.grid(row=0, column=0, padx=10, pady=10, sticky="w")

        self.entry_categoria = tk.Entry(root)
        self.entry_categoria.grid(row=0, column=1, padx=10, pady=10, sticky="ew")

        self.label_marca = tk.Label(root, text="Marca do Produto")
        self.label_marca.grid(row=1, column=0, padx=10, pady=10, sticky="w")

        self.entry_marca = tk.Entry(root)
        self.entry_marca.grid(row=1, column=1, padx=10, pady=10, sticky="ew")

        self.label_nome = tk.Label(root, text="Nome do Produto")
        self.label_nome.grid(row=2, column=0, padx=10, pady=10, sticky="w")

        self.entry_nome = tk.Entry(root)
        self.entry_nome.grid(row=2, column=1, padx=10, pady=10, sticky="ew")

        self.label_quantidade = tk.Label(root, text="Quantidade")
        self.label_quantidade.grid(row=3, column=0, padx=10, pady=10, sticky="w")

        self.entry_quantidade = tk.Entry(root)
        self.entry_quantidade.grid(row=3, column=1, padx=10, pady=10, sticky="ew")

        self.label_valor_fabrica = tk.Label(root, text="Valor de Fábrica")
        self.label_valor_fabrica.grid(row=4, column=0, padx=10, pady=10, sticky="w")

        self.entry_valor_fabrica = tk.Entry(root)
        self.entry_valor_fabrica.grid(row=4, column=1, padx=10, pady=10, sticky="ew")

        self.label_valor_loja = tk.Label(root, text="Valor Loja")
        self.label_valor_loja.grid(row=5, column=0, padx=10, pady=10, sticky="w")

        self.entry_valor_loja = tk.Entry(root)
        self.entry_valor_loja.grid(row=5, column=1, padx=10, pady=10, sticky="ew")

        # Label acima do Listbox para indicar o título da lista de produtos
        self.label_lista_produtos = tk.Label(root, text="Produtos no Estoque")
        self.label_lista_produtos.grid(row=6, column=0, columnspan=2, padx=10, pady=10, sticky="w")

        # Frame para o Listbox e Scrollbar
        self.frame_listbox = tk.Frame(root)
        self.frame_listbox.grid(row=7, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")

        # Barra de rolagem
        self.scrollbar = tk.Scrollbar(self.frame_listbox, orient="vertical")

        # Listbox para exibir produtos disponíveis
        self.listbox_produtos = tk.Listbox(self.frame_listbox, height=10, yscrollcommand=self.scrollbar.set)
        self.listbox_produtos.pack(side="left", fill="both", expand=True)

        # Configuração da barra de rolagem
        self.scrollbar.config(command=self.listbox_produtos.yview)
        self.scrollbar.pack(side="right", fill="y")

        # Configuração da expansão do frame ao redimensionar
        self.frame_listbox.grid_rowconfigure(0, weight=1)
        self.frame_listbox.grid_columnconfigure(0, weight=1)

        # Botões (organizados em 3x3)
        self.button_add = tk.Button(root, text="Adicionar ao Estoque", command=self.add_produto)
        self.button_add.grid(row=8, column=0, padx=10, pady=10, sticky="ew")

        self.button_remover = tk.Button(root, text="Remover do Estoque", command=self.remover_produto)
        self.button_remover.grid(row=8, column=1, padx=10, pady=10, sticky="ew")

        self.button_atualizar = tk.Button(root, text="Adicionar Unidades", command=self.adicionar_unidades)
        self.button_atualizar.grid(row=9, column=0, padx=10, pady=10, sticky="ew")

        self.button_remover_unidades = tk.Button(root, text="Remover Unidades", command=self.remover_unidades_popup)
        self.button_remover_unidades.grid(row=9, column=1, padx=10, pady=10, sticky="ew")

        self.button_atualizar_preco_fabrica = tk.Button(root, text="Atualizar Preço de Fábrica",
                                                        command=self.atualizar_preco_fabrica)
        self.button_atualizar_preco_fabrica.grid(row=10, column=0, padx=10, pady=10, sticky="ew")

        self.button_atualizar_preco_loja = tk.Button(root, text="Atualizar Preço Loja",
                                                     command=self.atualizar_preco_loja)
        self.button_atualizar_preco_loja.grid(row=10, column=1, padx=10, pady=10, sticky="ew")

        self.button_exportar = tk.Button(root, text="Exportar Estoque para Excel", command=self.exportar_estoque_excel)
        self.button_exportar.grid(row=11, column=0, columnspan=2, padx=10, pady=10, sticky="ew")

        """ Adiciona o botão de importar uma planilha Excel existente
        self.button_importar = tk.Button(root, text="Importar Estoque de Planilha", command=self.importar_estoque_excel)
        self.button_importar.grid(row=12, column=0, columnspan=2, padx=10, pady=10, sticky="ew") """

        # Fecha a aplicação ao clicar no "X"
        self.root.protocol("WM_DELETE_WINDOW", self.root.quit)

        # Carregar estoque do Excel, se existir
        self.carregar_estoque_do_excel()
        self.atualizar_listbox()

    def carregar_estoque_do_excel(self):
        """
        Carrega o estoque existente do arquivo Excel, se ele já existir.
        """
        if os.path.exists('estoque.xlsx'):
            try:
                df = pd.read_excel('estoque.xlsx')
                for _, row in df.iterrows():
                    self.estoque[row['Produto'].lower()] = {
                        'quantidade': row['Quantidade'],
                        'categoria': row['Categoria'].lower(),
                        'marca': row['Marca'].lower(),
                        'valor_fabrica': row['Valor de Fábrica'],
                        'valor_loja': row['Valor Loja']
                    }
                print("Estoque carregado do arquivo 'estoque.xlsx'.")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao carregar estoque: {e}")

    """ def importar_estoque_excel(self):
        
        Importa uma planilha Excel existente e carrega os produtos no sistema.
        - Foca nas colunas 'Marca', 'Quantidade', 'Valor de Fábrica' e 'Categoria', independentemente dos nomes exatos das colunas na planilha.
        - Se a coluna de "Categoria" estiver ausente, ela será detectada automaticamente com base no nome do produto.
        - Deixa o campo "Valor Loja" zerado.
        
        # Abre um diálogo para selecionar o arquivo Excel
        caminho_arquivo = filedialog.askopenfilename(
            title="Selecione a planilha de estoque",
            filetypes=(("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*"))
        )

        if caminho_arquivo:
            try:
                # Lê o arquivo Excel selecionado
                df = pd.read_excel(caminho_arquivo)

                # Normaliza os nomes das colunas para evitar problemas de capitalização e espaços
                df.columns = df.columns.str.strip().str.lower()

                # Palavras-chave para identificar as colunas (mais algumas sugestões)
                palavras_chave_colunas = {
                    'produto': ['produto', 'nome', 'descrição', 'item', 'artigo'],
                    'marca': ['marca', 'fabricante', 'fornecedor', 'empresa'],
                    'quantidade': ['quantidade', 'estoque', 'unidades', 'qtd', 'qtde'],
                    'valor de fábrica': ['valor', 'preço', 'fábrica', 'custo', 'valor fábrica', 'preço de fábrica'],
                    'categoria': ['categoria', 'tipo', 'grupo']  # Categoria será opcional
                }

                # Dicionário para mapear colunas por palavras-chave associadas
                colunas_esperadas = {
                    'produto': None,
                    'marca': None,
                    'quantidade': None,
                    'valor de fábrica': None,
                    'categoria': None  # Se não existir, será detectada automaticamente
                }

                # Verifica as colunas existentes e as mapeia com base nas palavras-chave
                for coluna in df.columns:
                    coluna_str = str(coluna).strip().lower()
                    for campo, palavras_chave in palavras_chave_colunas.items():
                        if any(palavra in coluna_str for palavra in palavras_chave):
                            colunas_esperadas[campo] = coluna
                            break

                # Verifica se colunas essenciais foram identificadas (exceto categoria, que é opcional)
                if None in [colunas_esperadas['produto'], colunas_esperadas['marca'], colunas_esperadas['quantidade'],
                            colunas_esperadas['valor de fábrica']]:
                    raise ValueError(
                        "A planilha está faltando colunas essenciais (Produto, Quantidade, Marca, Valor de Fábrica).")

                # Limpa o estoque atual antes de importar
                self.estoque.clear()

                # Função para detectar a categoria automaticamente com base no nome do produto
                categorias_automaticas = {
                    'teclado': 'Categoria',
                    'mouse': 'Categoria',
                    'monitor': 'Categoria',
                    'cabo': 'Categoria',
                    'fone': 'Categoria',
                    'caixa de som': 'Categoria',
                    'notebook': 'Categoria',
                    'cpu': 'Categoria',
                    'impressora': 'Categoria',
                    'tablet': 'Categoria',
                    'Redragon': 'Marca',
                    'Logitech': 'Marca',
                    'Razer': 'Marca',
                    # Adicione mais categorias conforme necessário
                } 

                def detectar_categoria(produto_nome):
                    
                    Detecta a categoria de um produto com base em seu nome.
                    Se não encontrar correspondência, atribui 'Outros'.
                    
                    produto_nome = produto_nome.lower()
                    for palavra_chave, categoria in categorias_automaticas.items():
                        if palavra_chave in produto_nome:
                            return categoria
                    return 'Outros'  # Caso nenhuma palavra-chave seja encontrada

                # Itera sobre as linhas da planilha e importa os dados para o dicionário de estoque
                for _, row in df.iterrows():
                    nome_produto = str(row[colunas_esperadas['produto']]).strip().lower()
                    marca = str(row[colunas_esperadas['marca']]).strip().lower()
                    quantidade = int(row[colunas_esperadas['quantidade']])
                    valor_fabrica = float(row[colunas_esperadas['valor de fábrica']])

                    # Se a planilha não tiver uma coluna 'Categoria', detecta a categoria automaticamente
                    if colunas_esperadas['categoria'] and pd.notna(row[colunas_esperadas['categoria']]):
                        categoria = str(row[colunas_esperadas['categoria']]).strip().lower()
                    else:
                        categoria = detectar_categoria(nome_produto)

                    # Insere no dicionário de estoque, com o valor da loja sempre zero
                    self.estoque[nome_produto] = {
                        'quantidade': quantidade,
                        'categoria': categoria,  # Categoria automática ou da planilha
                        'marca': marca,
                        'valor_fabrica': valor_fabrica,
                        'valor_loja': 0.0  # Valor da loja sempre zero
                    }

                # Atualiza o Listbox e salva os dados importados no arquivo padrão
                self.atualizar_listbox()
                self.atualizar_planilha_excel()

                messagebox.showinfo("Sucesso", "Estoque importado com sucesso.")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao importar o arquivo: {e}") """

    def validar_entrada(self):
        nome = self.entry_nome.get().strip().lower()
        categoria = self.entry_categoria.get().strip().lower()
        marca = self.entry_marca.get().strip().lower()
        try:
            quantidade = int(self.entry_quantidade.get())
            valor_fabrica = float(self.entry_valor_fabrica.get())
            valor_loja = float(self.entry_valor_loja.get())
            if quantidade < 0 or valor_fabrica < 0 or valor_loja < 0:
                raise ValueError
            return nome, quantidade, categoria, marca, valor_fabrica, valor_loja
        except ValueError:
            messagebox.showwarning("Erro", "Por favor, insira valores válidos.")
            return None, None, None, None, None, None

    def verificar_planilha_existente(self):
        """
        Verifica se a planilha Excel já existe e retorna um DataFrame se existir.
        """
        if os.path.exists('estoque.xlsx'):
            try:
                return pd.read_excel('estoque.xlsx')
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao verificar planilha existente: {e}")
        return None

    def atualizar_planilha_excel(self):
        """
        Função responsável por manter a planilha Excel alinhada com o estado atual do estoque.
        """
        try:
            # Cria um DataFrame com os produtos atuais do estoque
            df_novo = pd.DataFrame([(info['categoria'], info['marca'], produto, info['quantidade'],
                                     info['valor_fabrica'], info['valor_loja'],
                                     info['quantidade'] * info['valor_fabrica'])
                                    for produto, info in
                                    sorted(self.estoque.items(), key=lambda x: (x[1]['categoria'], x[0]))],
                                   columns=['Categoria', 'Marca', 'Produto', 'Quantidade', 'Valor de Fábrica',
                                            'Valor Loja', 'Investimento necessário'])

            # Filtra apenas os produtos que ainda estão no estoque
            df_novo = df_novo[df_novo['Quantidade'] > 0]

            with pd.ExcelWriter('estoque.xlsx', engine='xlsxwriter') as writer:
                df_novo.to_excel(writer, sheet_name='Itens', index=False)

                # Acessando o workbook e a planilha
                workbook = writer.book
                worksheet = writer.sheets['Itens']

                # Ajustando a largura das colunas de acordo com o conteúdo
                for i, column in enumerate(df_novo.columns):
                    max_length = df_novo[column].astype(str).map(len).max()
                    worksheet.set_column(i, i, max_length + 2)

                # Ajustando a largura das colunas adicionais manualmente para garantir a visualização
                worksheet.set_column('A:A', 20)
                worksheet.set_column('B:B', 20)
                worksheet.set_column('C:C', 20)
                worksheet.set_column('D:D', 15)
                worksheet.set_column('E:E', 20)
                worksheet.set_column('F:F', 20)
                worksheet.set_column('G:G', 30)

                # Formata a coluna "Valor" como moeda
                moeda_format = workbook.add_format({'num_format': 'R$ #,##0.00'})

                # Aplicando o formato de moeda para as colunas "Valor de Fábrica", "Valor Loja" e "Investimento necessário"
                worksheet.set_column('E:E', 20, moeda_format)  # Valor de Fábrica
                worksheet.set_column('F:F', 20, moeda_format)  # Valor Loja
                worksheet.set_column('G:G', 30, moeda_format)  # Investimento necessário

            print("Estoque exportado para o arquivo 'estoque.xlsx'.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao exportar estoque: {e}")

    def adicionar_ou_atualizar_produto(self, nome, quantidade, categoria, marca, valor_fabrica, valor_loja):
        """
        Adiciona um novo produto ou atualiza um existente no dicionário de estoque.
        """
        if nome in self.estoque:
            produto = self.estoque[nome]
            produto['quantidade'] += quantidade
            produto['categoria'] = categoria
            produto['marca'] = marca
            produto['valor_fabrica'] = valor_fabrica
            produto['valor_loja'] = valor_loja
        else:
            self.estoque[nome] = {
                'quantidade': quantidade,
                'categoria': categoria,
                'marca': marca,
                'valor_fabrica': valor_fabrica,
                'valor_loja': valor_loja
            }

    def add_produto(self):
        nome, quantidade, categoria, marca, valor_fabrica, valor_loja = self.validar_entrada()
        if nome:
            self.adicionar_ou_atualizar_produto(nome, quantidade, categoria, marca, valor_fabrica, valor_loja)
            self.atualizar_planilha_excel()
            self.atualizar_listbox()
            messagebox.showinfo("Sucesso", f"Produto '{nome.capitalize()}' adicionado/atualizado com sucesso.")

    def remover_produto(self):
        """
        Remove o produto selecionado da interface, do dicionário e da planilha Excel.
        """
        produto_selecionado = self.listbox_produtos.get(tk.ACTIVE)
        if produto_selecionado:
            nome = produto_selecionado.split(" - ")[0].lower()  # Extrair o nome do produto
            if nome in self.estoque:
                del self.estoque[nome]  # Remove o produto do dicionário de estoque

                # Atualiza a planilha Excel para refletir a remoção
                self.atualizar_planilha_excel()

                # Atualiza o Listbox para remover o produto da interface
                self.atualizar_listbox()

                messagebox.showinfo("Sucesso", f"Produto '{produto_selecionado}' removido com sucesso.")
            else:
                messagebox.showwarning("Aviso", "Produto não encontrado no estoque.")
        else:
            messagebox.showwarning("Aviso", "Nenhum produto selecionado.")

    def adicionar_unidades(self):
        """
        Abre um diálogo para adicionar unidades a um produto selecionado.
        """
        produto_selecionado = self.listbox_produtos.get(tk.ACTIVE)
        if not produto_selecionado:
            messagebox.showwarning("Aviso", "Nenhum produto selecionado.")
            return

        nome = produto_selecionado.split(" - ")[0].lower()
        if nome not in self.estoque:
            messagebox.showwarning("Aviso", "Produto não encontrado no estoque.")
            return

        # Solicita a quantidade a ser adicionada
        quantidade = simpledialog.askinteger("Adicionar Unidades",
                                             f"Quantas unidades de '{nome.capitalize()}' deseja adicionar?",
                                             parent=self.root, minvalue=1)
        if quantidade is None:
            # O usuário cancelou a operação
            return

        # Atualiza a quantidade
        self.estoque[nome]['quantidade'] += quantidade
        self.atualizar_planilha_excel()
        self.atualizar_listbox()
        messagebox.showinfo("Sucesso", f"{quantidade} unidades de '{nome.capitalize()}' adicionadas com sucesso.")

    def remover_unidades_popup(self):
        """
        Abre um diálogo para remover unidades de um produto selecionado.
        """
        produto_selecionado = self.listbox_produtos.get(tk.ACTIVE)
        if not produto_selecionado:
            messagebox.showwarning("Aviso", "Nenhum produto selecionado.")
            return

        nome = produto_selecionado.split(" - ")[0].lower()
        if nome not in self.estoque:
            messagebox.showwarning("Aviso", "Produto não encontrado no estoque.")
            return

        # Solicita a quantidade a ser removida
        quantidade = simpledialog.askinteger("Remover Unidades",
                                             f"Quantas unidades de '{nome.capitalize()}' deseja remover?",
                                             parent=self.root, minvalue=1)
        if quantidade is None:
            # O usuário cancelou a operação
            return

        # Verifica se a quantidade para remover é válida
        if self.estoque[nome]['quantidade'] < quantidade:
            messagebox.showwarning("Aviso", "Quantidade para remover maior do que a disponível.")
            return

        # Atualiza a quantidade
        self.estoque[nome]['quantidade'] -= quantidade
        if self.estoque[nome]['quantidade'] == 0:
            del self.estoque[nome]
            messagebox.showinfo("Sucesso",
                                f"Todas as unidades de '{nome.capitalize()}' foram removidas e o produto foi deletado do estoque.")
        else:
            messagebox.showinfo("Sucesso", f"{quantidade} unidades de '{nome.capitalize()}' removidas com sucesso.")

        self.atualizar_planilha_excel()
        self.atualizar_listbox()

    def atualizar_preco_fabrica(self):
        produto_selecionado = self.listbox_produtos.get(tk.ACTIVE)
        if produto_selecionado:
            nome = produto_selecionado.split(" - ")[0].lower()

            # Solicita o novo preço de fábrica
            valor_fabrica = simpledialog.askfloat("Atualizar Preço de Fábrica",
                                                  f"Digite o novo preço de fábrica para '{nome.capitalize()}':",
                                                  parent=self.root)
            if valor_fabrica is None:
                return  # O usuário cancelou a operação

            if nome in self.estoque:
                self.estoque[nome]['valor_fabrica'] = valor_fabrica
                self.atualizar_planilha_excel()
                self.atualizar_listbox()
                messagebox.showinfo("Sucesso",
                                    f"Preço de fábrica do produto '{produto_selecionado}' atualizado com sucesso.")
            else:
                messagebox.showwarning("Aviso", "Produto não encontrado no estoque.")
        else:
            messagebox.showwarning("Aviso", "Nenhum produto selecionado.")

    def atualizar_preco_loja(self):
        produto_selecionado = self.listbox_produtos.get(tk.ACTIVE)
        if produto_selecionado:
            nome = produto_selecionado.split(" - ")[0].lower()

            # Solicita o novo preço de loja
            valor_loja = simpledialog.askfloat("Atualizar Preço Loja",
                                               f"Digite o novo preço de loja para '{nome.capitalize()}':",
                                               parent=self.root)
            if valor_loja is None:
                return  # O usuário cancelou a operação

            if nome in self.estoque:
                self.estoque[nome]['valor_loja'] = valor_loja
                self.atualizar_planilha_excel()
                self.atualizar_listbox()
                messagebox.showinfo("Sucesso",
                                    f"Preço de loja do produto '{produto_selecionado}' atualizado com sucesso.")
            else:
                messagebox.showwarning("Aviso", "Produto não encontrado no estoque.")
        else:
            messagebox.showwarning("Aviso", "Nenhum produto selecionado.")

    def atualizar_listbox(self):
        """
        Atualiza o Listbox com os produtos atuais do estoque.
        """
        self.listbox_produtos.delete(0, tk.END)
        for nome, info in sorted(self.estoque.items(), key=lambda x: (x[1]['categoria'], x[0])):
            self.listbox_produtos.insert(tk.END,
                                         f"{nome.capitalize()} - Categoria: {info['categoria'].capitalize()} - Marca: {info['marca'].capitalize()} - Quantidade: {info['quantidade']} - Valor de Fábrica: R$ {info['valor_fabrica']:.2f} - Valor Loja: R$ {info['valor_loja']:.2f}")

    def exportar_estoque_excel(self):
        """
        Exporta o estoque atual para um arquivo Excel.
        """
        self.atualizar_planilha_excel()


if __name__ == "__main__":
    root = tk.Tk()
    app = EstoqueApp(root)
    root.mainloop()