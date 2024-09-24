import tkinter as tk
from tkinter import messagebox, filedialog
import pandas as pd
import os

class FinanceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Gerenciamento Financeiro")
        self.transacoes = []  # Lista de transações: cada item é um dicionário {tipo, descrição, valor, data}

        # Configuração da grid responsiva
        root.grid_rowconfigure(6, weight=1)
        root.grid_columnconfigure(1, weight=1)

        # Labels e Entradas
        self.label_tipo = tk.Label(root, text="Tipo (Receita/Despesa)")
        self.label_tipo.grid(row=0, column=0, padx=10, pady=10, sticky="w")

        self.entry_tipo = tk.Entry(root)
        self.entry_tipo.grid(row=0, column=1, padx=10, pady=10, sticky="ew")

        self.label_descricao = tk.Label(root, text="Descrição")
        self.label_descricao.grid(row=1, column=0, padx=10, pady=10, sticky="w")

        self.entry_descricao = tk.Entry(root)
        self.entry_descricao.grid(row=1, column=1, padx=10, pady=10, sticky="ew")

        self.label_valor = tk.Label(root, text="Valor")
        self.label_valor.grid(row=2, column=0, padx=10, pady=10, sticky="w")

        self.entry_valor = tk.Entry(root)
        self.entry_valor.grid(row=2, column=1, padx=10, pady=10, sticky="ew")

        # Label acima do Listbox para indicar o título da lista de transações
        self.label_lista_transacoes = tk.Label(root, text="Transações")
        self.label_lista_transacoes.grid(row=3, column=0, columnspan=2, padx=10, pady=10, sticky="w")

        # Frame para o Listbox e Scrollbar
        self.frame_listbox = tk.Frame(root)
        self.frame_listbox.grid(row=4, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")

        # Barra de rolagem
        self.scrollbar = tk.Scrollbar(self.frame_listbox, orient="vertical")

        # Listbox para exibir transações
        self.listbox_transacoes = tk.Listbox(self.frame_listbox, height=10, yscrollcommand=self.scrollbar.set)
        self.listbox_transacoes.pack(side="left", fill="both", expand=True)

        # Configuração da barra de rolagem
        self.scrollbar.config(command=self.listbox_transacoes.yview)
        self.scrollbar.pack(side="right", fill="y")

        # Configuração da expansão do frame ao redimensionar
        self.frame_listbox.grid_rowconfigure(0, weight=1)
        self.frame_listbox.grid_columnconfigure(0, weight=1)

        # Botões
        self.button_add = tk.Button(root, text="Adicionar Transação", command=self.adicionar_transacao)
        self.button_add.grid(row=5, column=0, padx=10, pady=10, sticky="ew")

        self.button_remover = tk.Button(root, text="Remover Transação", command=self.remover_transacao)
        self.button_remover.grid(row=5, column=1, padx=10, pady=10, sticky="ew")

        self.button_exportar = tk.Button(root, text="Exportar para Excel", command=self.exportar_transacoes)
        self.button_exportar.grid(row=6, column=0, columnspan=2, padx=10, pady=10, sticky="ew")

        self.button_saldo = tk.Button(root, text="Ver Saldo", command=self.calcular_saldo)
        self.button_saldo.grid(row=7, column=0, columnspan=2, padx=10, pady=10, sticky="ew")

        self.button_importar = tk.Button(root, text="Importar Planilha", command=self.importar_planilha)
        self.button_importar.grid(row=8, column=0, columnspan=2, padx=10, pady=10, sticky="ew")

        # Carregar transações do Excel, se existir
        self.carregar_transacoes_do_excel()
        self.atualizar_listbox()

    def carregar_transacoes_do_excel(self):
        """Carrega as transações existentes do arquivo Excel, se ele já existir."""
        if os.path.exists('transacoes.xlsx'):
            try:
                df = pd.read_excel('transacoes.xlsx')

                # Limpar espaços e garantir que os nomes das colunas estejam corretos
                df.columns = df.columns.astype(str).str.strip()

                if df.empty:
                    df = pd.DataFrame(columns=['Tipo', 'Descrição', 'Valor', 'Data'])
                else:
                    # Renomear as colunas para o formato correto
                    df.rename(columns=lambda x: x.strip().capitalize(), inplace=True)

                    # Carregar as transações do DataFrame
                    for _, row in df.iterrows():
                        self.transacoes.append({
                            'Tipo': row['Tipo'],
                            'Descrição': row['Descrição'],
                            'Valor': row['Valor'],
                            'Data': row['Data']
                        })

            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao carregar transações: {e}")

    def validar_entrada(self):
        """Valida as entradas do usuário para adicionar uma nova transação."""
        tipo = self.entry_tipo.get().strip().lower()
        descricao = self.entry_descricao.get().strip()
        try:
            valor = float(self.entry_valor.get())
            if tipo not in ['receita', 'despesa'] or valor < 0:
                raise ValueError
            return tipo.capitalize(), descricao, valor
        except ValueError:
            messagebox.showwarning("Erro", "Por favor, insira um tipo válido (Receita/Despesa) e um valor positivo.")
            return None, None, None

    def adicionar_transacao(self):
        """Adiciona uma nova transação (receita ou despesa)."""
        tipo, descricao, valor = self.validar_entrada()
        if tipo:
            self.transacoes.append({
                'Tipo': tipo,
                'Descrição': descricao,
                'Valor': valor,
                'Data': pd.Timestamp.now().strftime("%d/%m/%Y %H:%M")
            })
            self.exportar_transacoes()
            self.atualizar_listbox()
            messagebox.showinfo("Sucesso", f"Transação '{descricao}' adicionada com sucesso.")
            self.limpar_entradas()

    def remover_transacao(self):
        """Remove a transação selecionada."""
        transacao_selecionada = self.listbox_transacoes.get(tk.ACTIVE)
        if transacao_selecionada:
            index = self.listbox_transacoes.index(tk.ACTIVE)
            del self.transacoes[index]
            self.exportar_transacoes()
            self.atualizar_listbox()
            messagebox.showinfo("Sucesso", "Transação removida com sucesso.")
        else:
            messagebox.showwarning("Aviso", "Nenhuma transação selecionada.")

    def atualizar_listbox(self):
        """Atualiza o Listbox com as transações atuais, mostrando receitas primeiro e despesas depois."""
        self.listbox_transacoes.delete(0, tk.END)  # Limpa a listbox

        # Separar receitas e despesas
        receitas = [transacao for transacao in self.transacoes if transacao['Tipo'].lower() == 'receita']
        despesas = [transacao for transacao in self.transacoes if transacao['Tipo'].lower() == 'despesa']

        # Ordenar por data ou outro critério se necessário (opcional)
        receitas.sort(key=lambda x: x['Data'])  # Ordena receitas por data
        despesas.sort(key=lambda x: x['Data'])  # Ordena despesas por data

        # Inserir receitas no Listbox
        for transacao in receitas:
            self.listbox_transacoes.insert(
                tk.END,
                f"{transacao['Data']} - {transacao['Tipo']}: {transacao['Descrição']} - Valor: R$ {transacao['Valor']:.2f}"
            )

        # Inserir despesas no Listbox
        for transacao in despesas:
            self.listbox_transacoes.insert(
                tk.END,
                f"{transacao['Data']} - {transacao['Tipo']}: {transacao['Descrição']} - Valor: R$ {transacao['Valor']:.2f}"
            )

    def calcular_saldo(self):
        """Calcula o saldo atual (receitas - despesas)."""
        saldo = 0
        for transacao in self.transacoes:
            if transacao['Tipo'].lower() == 'receita':
                saldo += transacao['Valor']
            else:
                saldo -= transacao['Valor']
        messagebox.showinfo("Saldo Atual", f"O saldo atual é: R$ {saldo:.2f}")

    def exportar_transacoes(self):
        """Exporta as transações para um arquivo Excel, separando receitas primeiro e despesas depois."""
        try:
            # Separar receitas e despesas
            receitas = [transacao for transacao in self.transacoes if transacao['Tipo'].lower() == 'receita']
            despesas = [transacao for transacao in self.transacoes if transacao['Tipo'].lower() == 'despesa']

            # Combinar receitas e despesas, com receitas primeiro
            transacoes_ordenadas = receitas + despesas

            # Criar DataFrame com as transações organizadas
            df = pd.DataFrame(transacoes_ordenadas)

            # Cria um arquivo Excel com xlsxwriter
            with pd.ExcelWriter('transacoes.xlsx', engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Transações')

                # Acessa o workbook e a worksheet
                workbook = writer.book
                worksheet = writer.sheets['Transações']

                # Ajusta a largura das colunas
                worksheet.set_column('A:A', 20)  # Tipo
                worksheet.set_column('B:B', 20)  # Descrição
                worksheet.set_column('C:C', 20)  # Valor
                worksheet.set_column('D:D', 20)  # Data

                # Formata a coluna "Valor" como moeda
                moeda_format = workbook.add_format({'num_format': 'R$ #,##0.00'})  # Formato de moeda brasileira
                worksheet.set_column('C:C', 20, moeda_format)  # Aplica o formato à coluna "Valor"

            messagebox.showinfo("Sucesso", "Transações exportadas para 'transacoes.xlsx'.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao exportar transações: {e}")

    def importar_planilha(self):
        """Permite ao usuário importar um arquivo Excel externo e identificar colunas automaticamente."""
        try:
            # Abrir um diálogo para o usuário escolher o arquivo Excel
            file_path = filedialog.askopenfilename(
                title="Selecione a planilha Excel",
                filetypes=[("Arquivo Excel", "*.xlsx"), ("Todos os arquivos", "*.*")]
            )

            if file_path:
                # Carregar o arquivo Excel selecionado
                df = pd.read_excel(file_path)

                # Limpar espaços e garantir que os nomes das colunas estejam corretos
                df.columns = df.columns.astype(str).str.strip()

                # Tentativa de identificar as colunas com base em inferências
                colunas = df.columns

                # Identificar a coluna de Tipo (deve conter "Receita" ou "Despesa")
                col_tipo = next(
                    (col for col in colunas if df[col].str.contains('Receita|Despesa', case=False, na=False).any()),
                    None)

                # Identificar a coluna de Descrição (geralmente é texto não numérico)
                col_descricao = next((col for col in colunas if df[col].dtype == 'object' and col != col_tipo), None)

                # Identificar a coluna de Valor (deve ser numérica)
                col_valor = next((col for col in colunas if pd.api.types.is_numeric_dtype(df[col])), None)

                # Identificar a coluna de Data (pode ser inferida se contiver datas ou ser gerada)
                col_data = next((col for col in colunas if pd.to_datetime(df[col], errors='coerce').notna().any()),
                                None)

                if not col_tipo or not col_descricao or not col_valor:
                    raise ValueError("Não foi possível identificar as colunas corretamente. Verifique o arquivo.")

                # Adicionar as transações importadas à lista de transações atual
                for _, row in df.iterrows():
                    tipo = row[col_tipo]
                    descricao = row[col_descricao]
                    valor = row[col_valor]
                    data = row[col_data] if col_data else pd.Timestamp.now().strftime(
                        "%d/%m/%Y %H:%M")  # Adicionar data atual se não houver

                    self.transacoes.append({
                        'Tipo': tipo,
                        'Descrição': descricao,
                        'Valor': valor,
                        'Data': data
                    })

                # Atualizar o Listbox e exportar as transações atualizadas
                self.atualizar_listbox()
                self.exportar_transacoes()

                messagebox.showinfo("Sucesso", "Transações importadas com sucesso!")
            else:
                messagebox.showinfo("Aviso", "Nenhum arquivo foi selecionado.")

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao importar a planilha: {e}")

    def limpar_entradas(self):
        """Limpa os campos de entrada."""
        self.entry_tipo.delete(0, tk.END)
        self.entry_descricao.delete(0, tk.END)
        self.entry_valor.delete(0, tk.END)

if __name__ == "__main__":
    root = tk.Tk()
    app = FinanceApp(root)
    root.mainloop()