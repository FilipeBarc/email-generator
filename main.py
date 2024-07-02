import tkinter as tk
from tkinter import *
import pandas as pd
import numpy as np
import os
import win32com.client as win32

class Main:
    def __init__(self):

        # Root principal
        self.root = tk.Tk()
        self.root.geometry("1120x540")
        p = PhotoImage(file='Base//logo.png')
        self.root.iconphoto(False, p)
        self.root.title('Emails')
        self.root.config(bg='#ffffff')
        self.table = None

        # Texto padrão
        texto = f"""
        <h4>Olá Participante, seu resgate foi aprovado!</h4>
            
        <p>O arquivo com as informações do seu plano está anexado.<br />
        Para segurança das suas informações use os 4 últimos digitos do seu CPF para abrir este arquivo.</p>  
        
        <h4>Tenha um excelente dia.</h4><br />"""

        # Label
        titulo = tk.Label(self.root, text="Encaminhador Automático de Emails")
        titulo.config(font=("Arial", 15))
        titulo.place(x=180, y=15)
        titulo.config(bg='#ffffff')

        assunto = tk.Label(self.root, text="Assunto")
        assunto.config(font=("Arial", 10))
        assunto.place(x=180, y=55)
        assunto.config(bg='#ffffff')

        corpo = tk.Label(self.root, text="Texto")
        corpo.config(font=("Arial", 10))
        corpo.place(x=180, y=115)
        corpo.config(bg='#ffffff')

        # Assunto
        self.assunto = tk.Text(self.root, height=1.2, width=60, bg="#f2f2f2")
        self.assunto.insert(tk.END, 'Informativo de Resgate **Teste**')
        self.assunto.place(x=180, y=80)

        # Caixa de texto
        self.texto = tk.Text(self.root, height=20, width=100, bg="#f2f2f2", font=("Arial", 12))
        self.texto.insert(tk.END, texto)
        self.texto.place(x=180, y=140)
        self.texto.config(bg='#ffffff')

        # Botões
        self.botao_exportar = tk.Button(self.root, text="Enviar", command=self.coletar_valores)
        self.botao_exportar.place(x=30, y=446)
        self.botao_exportar.config(width=16, bg='#ffffff', activebackground="#e6e6e6", activeforeground="Black")
        self.botao_sair = tk.Button(self.root, text="Sair", command=self.root.destroy)
        self.botao_sair.config(width=16, bg='#ffffff', activebackground="#e6e6e6", activeforeground="Black")
        self.botao_sair.place(x=30, y=476)

        self.root.mainloop()

    def coletar_valores(self):
        try:
            df = pd.read_excel('Base/Dados_Participantes.xlsx')
            for i in df.index:
                address = df['Email'][i]
                nome = df['Nome'][i].split(' ')[0].capitalize()
                texto = self.texto.get("1.0", END)
                texto = ','.join(texto.split(',')[1:])
                sa = str(df['ParticipanteSA'][i])
                assinatura = """<html><head>
                                <meta charset="utf-8">
                                <title>Documento sem título</title>
                                </head><body>
                                <table width="500" border="0" cellspacing="0" cellpadding="0"><tbody><tr>
                                <td width="500"><table width="500" border="0" cellspacing="0" cellpadding="0">
                                <tbody><tr><td width="150" align="center"><a href="https://www.visaoprev.com.br/"><img src="https://www.visaoprev.com.br/images/2024/Verso-azul-alta.png" width="230" height="89" alt="Visão Prev"></a></td>
                                <td width="20"><img src="../../../../Materiais Visuais/Assinaturas/Atualização/divisor.jpg" width="1" height="118" alt=""></td>
                                <td width="230"><table width="330" border="0" cellspacing="0" cellpadding="0"><tbody><tr>
                                <td style="font-family: 'Trebuchet MS'; font-size: 20px; color: #004E65;"><strong>Visão Prev</strong></td>
                                </tr><tr><td style="font-family: 'Trebuchet MS'; font-size: 15px; color: #666666;">Central de Atendimento</td>
                                </tr><tr></tr><tr> </tr><tr></tr><tr></tr><tr>
                                <td style="font-family: 'Trebuchet MS'; font-size: 13px; color: #666666;"><a style="color: #1AA06D" href="https://www.visaoprev.com.br/"><strong>www.visaoprev.com.br </strong></a></td>
                                </tr><tr><td height="10"></td></tr></tbody></table><table width="110" border="0" cellspacing="0" cellpadding="0"><tbody><tr>
                                <td width="20" align="center"><a href="https://www.facebook.com/visaoprevprevidencia/"><img src="https://imagens.visaoprev.com.br/v2icon-fb.jpg" width="13" height="12" alt="Facebook Visão Prev"></a></td>
                                <td width="10">&nbsp;</td><td width="20" align="center"><a href="https://www.instagram.com/visaoprev/"><img src="https://imagens.visaoprev.com.br/v2icon-insta.jpg" width="13" height="12" alt="Instagram Visão Prev"></a></td>
                                <td width="10">&nbsp;</td><td width="20" align="center"><a href="https://www.linkedin.com/company/9341321"><img src="https://imagens.visaoprev.com.br/v2icon-lkdin.jpg" width="13" height="12" alt="Linkedin Visão Prev"></a></td>
                                <td width="10">&nbsp;</td><td width="20" align="center"><a href="https://www.youtube.com/visaoprevprevidencia"><img src="https://imagens.visaoprev.com.br/v2icon-yt.jpg" width="13" height="12" alt="You Tube Visão Prev"></a></td>
                                </tr></tbody></table></td></tr></tbody></table></td></tr><tr></tr><tr>
                                <td><a href="https://quem-somos.visaoprev.com.br/#nossas-conquistas"><img src="https://www.visaoprev.com.br/IMAGES/2023/banner-FINAL-MESMO.jpg" width="499" height="73" alt=""></a></td>
                                </tr><tr></tr><tr><td><a href="https://maisvisao.visaoprev.com.br/"><img src="https://imagens.visaoprev.com.br/v2bnn-01.jpg" width="500" height="70" alt="Mais Visão"></a></td>
                                </tr><tr><td height="10"></td></tr><tr><td style="font-family: 'Trebuchet MS'; font-size: 10px; color: #666666;"><p><strong></strong>Visão Prev - Previdência Complementar do Grupo Telefônica<br>
                                Alameda Santos, 787 – Conjuntos 11 e 12 - Jd. Paulista – São Paulo</p><p>AVISO LEGAL - Esta mensagem e seu conteúdo, inclusive anexos, pode conter informações confidenciais e/ou legalmente privilegiadas sobre os nossos negócios, para uso exclusivo de seu(s) destinatário(s). Qualquer modificação, retransmissão, disseminação, impressão ou utilização não autorizada fica estritamente proibida. Se você recebeu esta mensagem por engano, por favor, informe ao remetente e remova as informações de seu sistema.</p></td>
                                </tr></tbody></table></body></html>"""

                outlook = win32.Dispatch('outlook.application')
                email = outlook.CreateItem(0)

                email.To = address
                email.Subject = self.assunto.get("1.0", END)
                email.HTMLBody = f'Olá {nome}, {texto}\n\n{assinatura}'
                anexo = f'{os.path.abspath("Anexo")}\\'
                email.Attachments.Add(f'{anexo}{sa}.pdf')
                email.Send()
        except Exception:
            self.primeiro_aviso = 'Problemas com envio'
            self.segundo_aviso = ' '
            self.aviso()

        self.primeiro_aviso = 'Email enviado!'
        self.segundo_aviso = ' '
        self.aviso()

    def aviso(self):
        # Janela que gera os avisos
        aviso_janela = tk.Toplevel()
        p = PhotoImage(file='Base/logo.png')
        # Janela
        aviso_janela.iconphoto(False, p)
        aviso_janela.title("Contribuições Esporádicas Visão Prev")
        aviso_janela.config(width=300, height=200)
        aviso_janela.resizable(width=False, height=False)
        aviso_janela.config(bg='#ffffff')
        # Botão
        botao_aviso = tk.Button(aviso_janela, text="Fechar", command=aviso_janela.destroy)
        botao_aviso.place(x=120, y=150)
        botao_aviso.config(bg='#ffffff', activebackground="#e6e6e6", activeforeground="Black")
        # Label
        label_aviso = tk.Label(aviso_janela, text=str(self.primeiro_aviso) + '\n' + str(self.segundo_aviso))
        label_aviso.config(font=("Courier", 10))
        label_aviso.place(x=60, y=60)
        label_aviso.config(bg='#ffffff')


Main()


