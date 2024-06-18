import pandas as pd
from twilio.rest import Client
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import json
import ttkbootstrap as ttk
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import matplotlib.pyplot as plt
from PIL import Image, ImageTk
from plyer import notification



file_path =  None

def enviar_email(destinatario, assunto, corpo_email, remetente, senha):
    # Configurar servidor SMTP
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()

    # Faça o login no servidor SMTP
    server.login(remetente, senha)

    # Criar mensagem de e-mail
    mensagem = MIMEMultipart()
    mensagem['From'] = remetente
    mensagem['To'] = destinatario
    mensagem['Subject'] = assunto

    # Adicionar o body ao e-mail
    mensagem.attach(MIMEText(corpo_email, 'html'))

    # Enviar e-mail
    server.sendmail(remetente, destinatario, mensagem.as_string())

    # Fechar conexão com o servidor SMTP
    server.quit()

def enviar_whatsapp(Nome, Telefone):
    account_sid = 'AC090731d97975a1f9c980ae191e633d8c'
    auth_token = '3f3c8a6d544e02cffc5e204be3c13812'
    client = Client(account_sid, auth_token)

    # URL da mídia que você deseja enviar
    media_url = 'https://i.imgur.com/w4JMJio.png'

    # Crie a mensagem com mídia
    message = client.messages.create(
        content_sid='HX03363df64be6f8ea1694fa7e696f98a8',
        from_='MGdfda1cd7f8f8ce4b6ddc993b7cbd01c1',
        content_variables=json.dumps({'1': Nome}),
        to=f'whatsapp:+55{Telefone}',
        media_url=media_url 
    )

    print(f"Enviando as mensagens para {Nome} ({Telefone}) via Whatsapp...")

def enviar_notificacao(titulo, mensagem):
    notification.notify(
        title=titulo,
        message=mensagem,
        app_name="Prelúdio Produções",  # Nome do seu aplicativo
        timeout=10,  # Tempo em segundos que a notificação será exibida
    )

def enviar_todas_mensagens():
    global dados, data_atual
    enviar_button.state(["disabled"])
    root.update_idletasks()
    # Iterar sobre todas as linhas do DataFrame
    for index, aluno in dados.iterrows():
        # Verificar se a data de vencimento é hoje
        if aluno['Data de Vencimento'].day == data_atual.day:
            # Construir o body do email
            corpo_email = f'''
                <p>Prezado Cliente {aluno['Nome']},</p>
                <p>Vimos por meio deste comunicar que há um débito pendente em sua conta referente à Prelúdio..</p>
                <p>Solicitamos, por gentileza, que providencie o pagamento o mais breve possível por meio do PIX para a seguinte chave: onlinepreludio@gmail.com. Pedimos que encaminhe o comprovante de pagamento para este mesmo endereço de e-mail ou através do número (12) 99682-4870.</p>
                <p>Agradecemos antecipadamente pela sua atenção e colaboração neste assunto.</p>
                <p>Atenciosamente,</p>
                <p>Prelúdio. 🎹🎵</p>
                <p><img src="https://i.imgur.com/w4JMJio.png" alt="Logo"></p>
'''


            # Enviar email
            try:
                enviar_email(aluno['Email'], 'Sua Mensalidade da Empresa Vence Hoje.', corpo_email, 'onlinepreludio@gmail.com', 'rudo ydca fnyn nzju')
                enviar_notificacao('Notificação de Pagamento', f"Notificação enviada para {aluno['Nome']}!")
            except Exception as e:
                print(f"Erro ao enviar e-mail para {aluno['Nome']}: {str(e)}")

            # Enviar mensagem via WhatsApp
            try:
                enviar_whatsapp(aluno['Nome'], aluno['Telefone'])
            except Exception as e:
                print(f"Erro ao enviar mensagem via WhatsApp para {aluno['Nome']}: {str(e)}")

    enviar_button.state(["!disabled"])
    root.update_idletasks()
    print("Enviando Mensagem via E-mail")

def carregar_arquivo():
    global dados
    global mensagem_status
    global file_path

    file_path = filedialog.askopenfilename(title="Selecione o arquivo Excel", filetypes=[("Arquivos Excel", "*.xlsx;*.xls")])

    try:
        # Tenta carregar o DataFrame a partir do arquivo Excel
        dados = pd.read_excel(file_path, dtype={'Pagou Mensalidade': bool})
        mensagem_status.set(f"Arquivo Excel '{file_path}'\n{' ' * ((len(file_path) + 40 - len('carregado com sucesso!')) // 2)}carregado com sucesso!")




    except Exception as e:
        mensagem_status.set(f"Erro ao carregar o arquivo Excel:\n{str(e)}")

def formatar_data_vencimento(data):
    # Formatar a data para "dia-mês-ano"
    return data.strftime("%d/%m/%Y")

# Função para exibir mensagem na interface      
def exibir_alunos_vencendo_hoje():
    global dados
    global data_atual

    alunos_vencendo_hoje = dados[dados['Data de Vencimento'].dt.day == data_atual.day]

    if not alunos_vencendo_hoje.empty:
        # Criar uma nova janela para exibir as informações dos alunos
        janela_alunos = tk.Toplevel(root)
        janela_alunos.title("Alunos Vencendo Hoje")

        # Criar o widget Treeview
        tree = ttk.Treeview(janela_alunos)
        tree['columns'] = tuple(alunos_vencendo_hoje.columns)

        # Definir as colunas
        for col in alunos_vencendo_hoje.columns:
            tree.column(col, anchor='center')
            tree.heading(col, text=col)

        # Preencher a tabela com os dados dos alunos
        for index, row in alunos_vencendo_hoje.iterrows():
            # Traduzir os valores booleanos para as strings
            status_pagamento = "Pagou" if row['Pagou Mensalidade'] else "Não Pagou"
            # Formatar a data de vencimento
            data_vencimento_formatada = formatar_data_vencimento(row['Data de Vencimento'])

            tree.insert('', 'end', values=(row['Nome'], row['Email'], row['Telefone'], data_vencimento_formatada, status_pagamento, row['Mensalidades Atrasadas'], row['Disciplina']))
            tree.tag_configure(f'I{index}', background='#333333', foreground='white')  # Ajuste opcional para a cor de fundo
        # Adicionar o widget Treeview à janela
        tree.pack(padx=10, pady=10)

    else:
        messagebox.showinfo("Nenhum Cliente", "Nenhum cliente com mensalidade vencendo hoje.")

def sort_column(tree, col, reverse=False):
    data = [(tree.set(child, col), child) for child in tree.get_children('')]
    data.sort(reverse=reverse)

    for index, (val, child) in enumerate(data):
        tree.move(child, '', index)

    tree.heading(col, command=lambda: sort_column(tree, col, not reverse))

def formatar_data(data):
    # Verificar se a data não é nula e é do tipo datetime
    if pd.notna(data) and isinstance(data, datetime):
        return datetime.strftime(data, "%d/%m/%Y")
    else:
        return ""

def recarregar_janela_alunos(tree, dados, file_path):
    # Limpar todos os itens na Treeview
    for item_id in tree.get_children():
        tree.delete(item_id)

    # Preencher a tabela com os dados atualizados
    for index, row in dados.iterrows():
        # Traduzir os valores 'Pagou Mensalidade' para 'Não Pagou' ou 'Pagou'
        status_pagamento = "Pagou" if row['Pagou Mensalidade'] else "Não Pagou"
        # Formatar a data de vencimento
        data_vencimento_formatada = "" if pd.isna(row['Data de Vencimento']) else row['Data de Vencimento'].strftime("%d/%m/%Y")
        values_list = [row['Nome'], row['Email'], row['Telefone'], data_vencimento_formatada, status_pagamento, str(row['Mensalidades Atrasadas']), row['Disciplina']]
        item_id = tree.insert('', 'end', values=values_list, tags=(f'I{index}', 'checkbox'))
        # Configurar a cor de fundo como #333333

        tree.insert(item_id, 'end', values='', tags=(f'I{index}', 'checkbox'))

    # Chamar tree.update() após um curto intervalo
    tree.update()

    # Adicionar esta linha para atualizar a exibição imediatamente
    tree.update_idletasks()

def enviar_notificacoes_intervalo():
    data_inicial = data_inicial_entry.get()
    data_final = data_final_entry.get()
    try:
        data_inicial = datetime.datetime.strptime(data_inicial, "%Y-%m-%d").date()
        data_final = datetime.datetime.strptime(data_final, "%Y-%m-%d").date()
    except ValueError:
        messagebox.showerror("Erro", "Formato de data inválido. Use YYYY-MM-DD.")
        return
    for aluno in dados.iterrows():
        data_vencimento = aluno['data_vencimento']
        if data_inicial <= data_vencimento <= data_final:
            enviar_todas_mensagens()


    

def exibir_alunos():
    global dados
    global data_atual

    if dados is None or dados.empty:
        messagebox.showinfo("Sem Dados", "Nenhum dado disponível para exibir.")
        return

    # Criar uma nova janela para exibir todos os alunos e suas informações
    janela_alunos = tk.Toplevel(root)
    janela_alunos.title("Informações dos Clientes")

    # Criar o widget Treeview
    tree = ttk.Treeview(janela_alunos)
    tree['columns'] = tuple(dados.columns)

    # Definir as colunas
    for col in dados.columns:
        tree.column(col, anchor='center')
        tree.heading(col, text=col, anchor='center', command=lambda c=col: sort_column(tree, c))  # Adicionando opção de ordenar
        # Adicionar formatação negrito usando a propriedade font
        tree.heading(col, text=col)

    # Criar uma instância de ttk.Style
    style = ttk.Style()
    # Configurar a fonte para o estilo 'Treeview.Heading'
    style.configure('Treeview.Heading', font=('Helvetica', 12, 'bold'))
    # Aplicar o estilo à coluna
    tree.tag_configure('Treeview.Heading', foreground='white', background='#333333', font=('Helvetica', 10, 'bold'))

    # Adicionar uma barra de rolagem vertical separada ao Treeview
    scrollbar_y = ttk.Scrollbar(janela_alunos, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=scrollbar_y.set)

     # Preencher a tabela com os dados dos alunos
    for index, row in dados.iterrows():
        # Traduzir os valores booleanos para as strings desejadas
        status_pagamento = "Pagou" if row['Pagou Mensalidade'] else "Não Pagou"
        
        # Formatar a data de vencimento usando a função formatar_data
        data_vencimento_formatada = formatar_data(row['Data de Vencimento'])

        # Adicionar uma caixa de verificação para cada aluno
        item_id = tree.insert('', 'end', values=(row['Nome'], row['Email'], row['Telefone'], data_vencimento_formatada, status_pagamento, row['Mensalidades Atrasadas'], row['Disciplina']), tags=(f'I{index}', 'checkbox'))
        tree.tag_configure(f'I{index}', background='#333333', foreground='white')  # Ajuste opcional para a cor de fundo
        tree.insert(item_id, 'end', values='', tags=(f'I{index}', 'checkbox'))

    # Configurar o formato de exibição da coluna de Data de Vencimento
    tree.column('Data de Vencimento', anchor='center', width=120)  # Ajustar a largura conforme necessário
    tree.heading('Data de Vencimento', text='Data de Vencimento', anchor='center', command=lambda: sort_column(tree, 'Data de Vencimento'))
    tree.pack(padx=10, pady=10, side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)

    # Criar uma figura para o gráfico de pizza
    fig = Figure(figsize=(3, 4), dpi=100)
    ax = fig.add_subplot(111)
    ax.set_frame_on(False)

    # Calcular a porcentagem de alunos que pagaram e nao pagaram
    pagou = dados['Pagou Mensalidade'].sum()
    nao_pagou = len(dados) - pagou

    labels = ['Pagou', 'Não\nPagou']
    values = [pagou, nao_pagou]

        # Tamanho especifico
    fig, ax = plt.subplots()
    ax.set_aspect('equal')
    plt.figure(figsize=(6, 6))  # Setando de 6 por 6 polegadas

    # Adicionar o gráfico de pizza à figura
    ax.pie(values, labels=labels, autopct='%1.1f%%', startangle=90, textprops={'color': 'white'})

    # Adicionar a figura ao canvas
    canvas = FigureCanvasTkAgg(fig, master=janela_alunos)
    canvas.draw()

    # Configurar a cor de fundo do Figure para ser transparente
    fig.patch.set_facecolor('none')

    # Configurar a largura da borda para 0 no widget FigureCanvasTkAgg
    canvas.get_tk_widget().config(borderwidth=0)

    # Configurar a largura da borda para 0 no contêiner (janela_alunos)
    janela_alunos.config(bd=0)

    # Posicionar o canvas no canto inferior direito abaixo da tabela
    canvas.get_tk_widget().pack(side=tk.BOTTOM, padx=2, pady=10)

    # Adicionar um botão para atualizar o pagamento
    botao_atualizar_pagamento = ttk.Button(janela_alunos, text="PAGOU", bootstyle='SUCCESS', command=lambda: atualizar_pagamento(tree, dados, file_path), width=23)
    botao_atualizar_pagamento.pack(side=tk.TOP, padx=10, pady=5, fill=tk.X)

    # Adicionar um botão para voltar o pagamento para False (unidade selecionada)
    botao_voltar_false_selecionado = ttk.Button(janela_alunos, text="NÃO PAGOU", bootstyle= 'danger', command=lambda: voltar_para_false_selecionado(tree, dados,file_path), width=23)
    botao_voltar_false_selecionado.pack(side=tk.TOP, padx=10, pady=5, fill=tk.X)

    # Adicionar um botão para voltar o pagamento para False (todos)
    botao_voltar_false_todos = tk.Button(janela_alunos, text="Início do Mês (Todos)", command=lambda: voltar_para_false_todos(tree, dados,file_path))
    botao_voltar_false_todos.pack(side=tk.RIGHT, padx=10, pady=5)

    # Adicionar um botão para subtrair -1 na coluna de mensalidades atrasadas do aluno selecionado
    botao_subtrair_um = ttk.Button(janela_alunos, text="Pagou 1 Mensalidade Atrasada", bootstyle='outline success', command=lambda: subtrair_mensalidade_atrasada(tree, dados, file_path))
    botao_subtrair_um.pack(side=tk.RIGHT, padx=10, pady=5)

    # Adicionar um botão para subtrair -1 na coluna de mensalidades atrasadas do aluno selecionado
    botao_adicionar_um = ttk.Button(janela_alunos, text="Atrasou 1 Mensalidade", bootstyle= 'outline danger', command=lambda: adicionar_mensalidade_atrasada(tree, dados, file_path))
    botao_adicionar_um.pack(side=tk.RIGHT, padx=10, pady=5)

    botao_recarregar = ttk.Button(janela_alunos, text="Recarregar", bootstyle= 'secondary', command=lambda: recarregar_janela_alunos(tree, dados, file_path))
    botao_recarregar.pack(side=tk.RIGHT, padx=10, pady=5)




def recarregar_janela_alunos(tree, dados, file_path):
    # Limpar todos os itens na Treeview
    for item_id in tree.get_children():
        tree.delete(item_id)

    # Preencher a tabela com os dados atualizados
    for index, row in dados.iterrows():
        # Traduzir os valores 'Pagou Mensalidade' para 'Não Pagou' ou 'Pagou'
        status_pagamento = "Pagou" if row['Pagou Mensalidade'] else "Não Pagou"
        
        # Formatar a data de vencimento usando a função formatar_data
        data_vencimento_formatada = formatar_data(row['Data de Vencimento'])

        values_list = [row['Nome'], row['Email'], row['Telefone'], data_vencimento_formatada, status_pagamento, str(row['Mensalidades Atrasadas']), row['Disciplina']]
        item_id = tree.insert('', 'end', values=values_list, tags=(f'I{index}', 'checkbox'))
        # Configurar a cor de fundo como #333333
        tree.tag_configure(f'I{index}', background='#333333', foreground='white')  # Ajuste opcional para a cor de texto
        tree.insert(item_id, 'end', values='', tags=(f'I{index}', 'checkbox'))

    # Chamar tree.update() após um curto intervalo
    tree.update()

    # Adicionar esta linha para atualizar a exibição imediatamente
    tree.update_idletasks()

# Adicionar a função para subtrair -1 na coluna de mensalidades atrasadas do aluno selecionado
def subtrair_mensalidade_atrasada(tree, dados, file_path):
    item_selecionado = tree.selection()

    if item_selecionado:
        # Obter o índice associado ao item selecionado
        tags = tree.item(item_selecionado[0], 'tags')

        if tags and tags[0].startswith('I'):  # Verificar se a tag começa com 'I' (indicando um índice)
            # O índice está na forma 'I00x00' onde 00x00 é o índice real
            index_str = tags[0][1:]

            # Verificar se o índice é uma string de dígitos
            if index_str.isdigit():
                index = int(index_str)  # Converta para inteiro
                # Verificar se o índice é válido
                if 0 <= index < len(dados):
                    # Subtrair -1 na coluna de mensalidades atrasadas
                    dados.at[index, 'Mensalidades Atrasadas'] -= 1

                    # Salvar o DataFrame de volta no arquivo JSON
                    try:
                        # Usar to_list() para obter uma lista em vez de to_dict()
                        contador_mensalidades = dados['Mensalidades Atrasadas'].to_list()
                        with open("contador_mensalidades.json", "w") as json_file:
                            json.dump(contador_mensalidades, json_file)
                    except Exception as e:
                        messagebox.showerror("Erro", f"Erro ao salvar o arquivo JSON:\n{str(e)}")

                    # Atualizar a exibição na árvore diretamente
                    tree.item(item_selecionado[0], values=(dados.iloc[index]['Nome'], dados.iloc[index]['Email'], dados.iloc[index]['Telefone'], dados.iloc[index]['Data de Vencimento'], "Pagou" if dados.iloc[index]['Pagou Mensalidade'] else "Não Pagou", dados.iloc[index]['Mensalidades Atrasadas'], dados.iloc[index]['Disciplina']))

                    #messagebox.showinfo("Mensalidade Atualizada", "A mensalidade atrasada foi paga com sucesso.")
                else:
                    messagebox.showwarning("Operação Inválida", "Índice de aluno inválido.")
            else:
                messagebox.showwarning("Operação Inválida", "Índice de aluno não é um número.")
    else:
        messagebox.showwarning("Nenhum Aluno Selecionado", "Selecione um aluno antes de pagar a mensalidade atrasada.")

    recarregar_janela_alunos(tree, dados, file_path)

def atualizar_exibicao(tree, index):
    # Atualizar a exibição na árvore diretamente
    item_selecionado = tree.selection()
    row = dados.iloc[index]
    
    # Traduzir o valor 'Pagou Mensalidade' para 'Pagou' ou 'Não Pagou'
    status_pagamento = "Pagou" if row['Pagou Mensalidade'] else "Não Pagou"

    # Atualizar os valores diretamente na árvore
    tree.item(item_selecionado[0], values=(row['Nome'], row['Email'], row['Telefone'], row['Data de Vencimento'], status_pagamento, row['Mensalidades Atrasadas'], row['Disciplina']))

def adicionar_mensalidade_atrasada(tree, dados, file_path):
    item_selecionado = tree.selection()

    if item_selecionado:
        # Obter o índice associado ao item selecionado
        tags = tree.item(item_selecionado[0], 'tags')

        if tags and tags[0].startswith('I'):  # Verificar se a tag começa com 'I' (indicando um índice)
            # O índice está na forma 'I00x00' onde 00x00 é o índice real
            index_str = tags[0][1:]

            # Verificar se o índice é uma string de dígitos
            if index_str.isdigit():
                index = int(index_str)  # Converta para inteiro
                # Verificar se o índice é válido
                if 0 <= index < len(dados):
                    # Subtrair -1 na coluna de mensalidades atrasadas
                    dados.at[index, 'Mensalidades Atrasadas'] += 1

                    # Salvar o DataFrame de volta no arquivo JSON
                    try:
                        # Usar to_list() para obter uma lista em vez de to_dict()
                        contador_mensalidades = dados['Mensalidades Atrasadas'].to_list()
                        with open("contador_mensalidades.json", "w") as json_file:
                            json.dump(contador_mensalidades, json_file)
                    except Exception as e:
                        messagebox.showerror("Erro", f"Erro ao salvar o arquivo JSON:\n{str(e)}")

                    # Atualizar a exibição na árvore diretamente
                    tree.item(item_selecionado[0], values=(dados.iloc[index]['Nome'], dados.iloc[index]['Email'], dados.iloc[index]['Telefone'], dados.iloc[index]['Data de Vencimento'], "Pagou" if dados.iloc[index]['Pagou Mensalidade'] else "Não Pagou", dados.iloc[index]['Mensalidades Atrasadas'], dados.iloc[index]['Disciplina']))

                    #messagebox.showinfo("Mensalidade Atualizada", "A mensalidade atrasada foi incrementada com sucesso.")
                else:
                    messagebox.showwarning("Operação Inválida", "Índice de aluno inválido.")
            else:
                messagebox.showwarning("Operação Inválida", "Índice de aluno não é um número.")
    else:
        messagebox.showwarning("Nenhum Aluno Selecionado", "Selecione um aluno antes de adicionar a mensalidade atrasada.")
    recarregar_janela_alunos(tree, dados, file_path)

def voltar_para_false_selecionado(tree, dados, file_path):
    item_selecionado = tree.selection()

    if item_selecionado:
        # Obter o índice associado ao item selecionado
        tags = tree.item(item_selecionado[0], 'tags')

        if tags and tags[0].startswith('I'):  # Verificar se a tag começa com 'I' (indicando um índice)
            # O índice está na forma 'I00x00' onde 00x00 é o índice real
            index_str = tags[0][1:]

            # Verificar se o índice é uma string de dígitos
            if index_str.isdigit():
                index = int(index_str)  # Converta para inteiro
                # Verificar se o índice é válido
                if 0 <= index < len(dados):
                    # Atualizar o valor 'Pagou Mensalidade' para False no DataFrame
                    dados.at[index, 'Pagou Mensalidade'] = False
                    # Salvar o DataFrame de volta no arquivo Excel
                    try:
                        dados.to_excel(file_path, index=False)
                    except Exception as e:
                        messagebox.showerror("Erro", f"Erro ao salvar o arquivo Excel:\n{str(e)}")
                    # Atualizar a exibição na árvore diretamente
                    row = dados.iloc[index]
                    # Criar uma lista com os valores e adicionar "Não Pago" no final
                    values_list = list(row[['Nome', 'Email', 'Telefone', 'Data de Vencimento']])
                    values_list.append("Não Pagou")
                    # Atualizar a exibição na árvore diretamente
                    tree.item(item_selecionado[0], values=tuple(values_list))
                else:
                    messagebox.showwarning("Operação Inválida", "Índice de aluno inválido.")
            else:
                messagebox.showwarning("Operação Inválida", "Índice de aluno não é um número.")
    else:
        messagebox.showwarning("Nenhum Aluno Selecionado", "Selecione um aluno antes de voltar para 'False'.")

    recarregar_janela_alunos(tree, dados, file_path)

def voltar_para_false_todos(tree, dados, file_path):

    resposta = messagebox.askquestion("Confirmação", "Deseja realmente iniciar o mês?")
    if resposta == 'yes':
        for index, row in dados.iterrows():
            if not row['Pagou Mensalidade']:
                dados.at[index, 'Mensalidades Atrasadas'] += 1  # Incrementa o contador
            dados.at[index, 'Pagou Mensalidade'] = False

        # Limpar todos os itens na Treeview
        for item_id in tree.get_children():
            tree.delete(item_id)

        # Preencher a tabela com os dados atualizados
        for index, row in dados.iterrows():
            values_list = list(row[['Nome', 'Email', 'Telefone', 'Data de Vencimento', 'Pagou Mensalidade', 'Mensalidades Atrasadas', 'Disciplina']].astype(str))
            item_id = tree.insert('', 'end', values=values_list, tags=(f'I{index}', 'checkbox'))
            # Configurar a cor de fundo como #333333
            tree.tag_configure(f'I{index}', background='#333333', foreground='white')  # Ajuste opcional para a cor de texto
            tree.insert(item_id, 'end', values='', tags=(f'I{index}', 'checkbox'))

        # Salvar o DataFrame de volta no arquivo Excel
        try:
            dados.to_excel(file_path, index=False)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar o arquivo Excel:\n{str(e)}")

        messagebox.showinfo("Pagamento Atualizado", "O pagamento de todos os alunos foi atualizado para 'Não Pago' e o contador de quem não pagou foi incrementado em 'Mensalidades Atrasadas'.")
    else:
        print("Operação Cancelada.)")

    recarregar_janela_alunos(tree, dados, file_path)

def atualizar_pagamento(tree, dados, file_path):
    item_selecionado = tree.selection()

    if item_selecionado:
        # Obter o índice associado ao item selecionado
        tags = tree.item(item_selecionado[0], 'tags')
        
        if tags and tags[0].startswith('I'):  # Verificar se a tag começa com 'I' (indicando um índice)
            # O índice está na forma 'I00x00' onde 00x00 é o índice real
            index_str = tags[0][1:]
            
            # Verificar se o índice é uma string de dígitos
            if index_str.isdigit():
                index = int(index_str)  # Converta para inteiro
                # Verificar se o índice é válido
                if 0 <= index < len(dados):
                    # Atualizar o valor 'Pagou Mensalidade' para True no DataFrame
                    dados.at[index, 'Pagou Mensalidade'] = not dados.at[index, 'Pagou Mensalidade']  # Alternar entre True e False

                    # Adicionar esta linha para atualizar a exibição imediatamente
                    tree.update_idletasks()

                    # Salvar o DataFrame de volta no arquivo Excel
                    try:
                        dados.to_excel(file_path, index=False)
                    except Exception as e:
                        messagebox.showerror("Erro", f"Erro ao salvar o arquivo Excel:\n{str(e)}")

                    # Atualizar a exibição na árvore diretamente
                    # Traduzir o valor 'Pagou Mensalidade' para 'Não Pagou' ou 'Pagou'
                    status_pagamento = "Pagou" if dados.at[index, 'Pagou Mensalidade'] else "Não Pagou"
                    tree.item(item_selecionado[0], values=(dados.iloc[index]['Nome'], dados.iloc[index]['Email'], dados.iloc[index]['Telefone'], dados.iloc[index]['Data de Vencimento'], status_pagamento, dados.iloc[index]['Mensalidades Atrasadas'], dados.iloc[index]['Disciplina']))

                    #messagebox.showinfo("Pagamento Atualizado", f"O pagamento foi atualizado para '{status_pagamento}' e o arquivo Excel foi salvo com sucesso.")
                else:
                    messagebox.showwarning("Operação Inválida", "Índice de aluno inválido.")
            else:
                messagebox.showwarning("Operação Inválida", "Índice de aluno não é um número.")
        else:
            messagebox.showwarning("Operação Inválida", "Selecione um aluno antes de atualizar o pagamento.")
    else:
        messagebox.showwarning("Nenhum Aluno Selecionado", "Selecione um aluno antes de atualizar o pagamento.")

    recarregar_janela_alunos(tree, dados, file_path)







# Configurar a interface gráfica        
      
root = ttk.Window(themename='superhero')
root.title("Prelúdio Produções")
root.geometry("400x650")


# Define a function to create a button with a given text, command, and style
def create_button(text, command, style):
    return ttk.Button(root, text=text, command=command, bootstyle=style)




# Função para carregar e redimensionar uma imagem para o tamanho desejado
def load_and_resize_image(image_path, width, height, master=None):
    original_image = Image.open(image_path)
    original_image.thumbnail((width, height))

    # Verifica se a janela principal está disponível
    if master is not None:
        tk_image = ImageTk.PhotoImage(original_image, master=master)
        master.update_idletasks()
        return tk_image, original_image
    else:
        return ImageTk.PhotoImage(original_image)
    
icon_path = "icons/"

# Crie ícones para os botões
carregar_icon = load_and_resize_image(icon_path + "carregar_icon.png", 32, 32)
enviar_icon = load_and_resize_image(icon_path + "enviar_icon.png", 32, 32)
exibir_vencendo_icon = load_and_resize_image(icon_path + "exibir_vencendo_icon.png", 32, 32)
exibir_todos_icon = load_and_resize_image(icon_path + "exibir_todos_icon.png", 32, 32)

# Função para criar um botão com um ícone abaixo do texto
def create_button_with_icon(text, command, style, icon_path):
    button = ttk.Button(root, text=text, command=command, style=style)

    # Carregar e redimensionar a imagem do ícone
    icon, _ = load_and_resize_image(icon_path, 32, 32, master=root)

    # Configurar o ícone abaixo do texto no botão
    button.image = icon
    button.config(image=icon, compound=tk.BOTTOM)

    return button

# Botão para carregar o arquivo Excel
carregar_button = create_button_with_icon("Selecionar Arquivo Excel", carregar_arquivo, "primary.large", icon_path + "carregar_icon.png")
carregar_button.pack(pady=20)

# Rótulo para exibir a mensagem de status
mensagem_status = tk.StringVar()
status_label = ttk.Label(root, textvariable=mensagem_status, bootstyle="info")
status_label.pack(pady=10)

# Frame para agrupar os botões
frame_botoes = ttk.Frame(root)
frame_botoes.pack(pady=20)

# Botão para enviar as mensagens
enviar_button = create_button_with_icon("Enviar Notificações", enviar_todas_mensagens, "success.large",  icon_path + "enviar_icon.png")
enviar_button.pack(pady=20)

# Botão para exibir os alunos vencendo hoje
exibir_alunos_vencendo_button = create_button_with_icon("Clientes com Vencimento Hoje", exibir_alunos_vencendo_hoje, "danger",  icon_path + "exibir_vencendo_icon.png")
exibir_alunos_vencendo_button.pack(pady=20)


# Botão para exibir todos os alunos e suas informações
exibir_alunos_button = create_button_with_icon("Ver Todos os Clientes", exibir_alunos, "secondary.large",  icon_path + "exibir_todos_icon.png")
exibir_alunos_button.pack(pady=20)

# Botao para escolher o intevalo de datas dos alunos
intervalo_data_button = create_button_with_icon("Enviar Notificações por Data", enviar_notificacoes_intervalo, "success.large", icon_path + "exibir_vencendo_icon.png")
intervalo_data_button.pack(pady=20)

# Obtenha as dimensões da tela
largura_tela = root.winfo_screenwidth()
altura_tela = root.winfo_screenheight()

# Variáveis globais
dados = None
data_atual = datetime.now().date()




# Iniciar o loop principal da interface gráfica
root.mainloop()