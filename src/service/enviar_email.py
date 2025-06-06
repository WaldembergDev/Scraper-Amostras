# criar a função que envia e-mail
import smtplib
from email.mime.text import MIMEText
from dotenv import load_dotenv
from datetime import date
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os

load_dotenv()

class EnviarEmail:
    @classmethod
    def criar_corpo_str(cls, dados: list):
        # Anexar uma versão de texto simples é uma boa prática
        plain_text_body = f"Prezados!,\n\nSegue a listagem de amostras que deverão ser liberadas hoje::\n\n"
        if not dados:
            plain_text_body += "Nenhum dado encontrado para o período.\n"
        else:
            for status_os, amostra, cliente, prioridade in dados:
                plain_text_body += f"Status OS: {status_os}, Amostra: {amostra}, Cliente: {cliente}, Prioridade: {prioridade}\n"
        plain_text_body += "\nAtenciosamente,\nSua Automação"

        return plain_text_body

    @classmethod
    def criar_corpo_html(cls, dados: list):
        html_content = '<p>Nenhum dado encontrado para o período.</p>'
        if not dados:
            return html_content
        html_content = """
    <html>
        <head></head>
        <body>
            <p>Prezados,</p>
            <p>Segue a listagem de amostras que deverão ser liberadas hoje:</p>
            <table style="width:100%; border-collapse: collapse;">
                <thead>
                    <tr style="background-color: #f2f2f2;">
                        <th style="border: 1px solid black; padding: 8px; text-align: left;">Status OS</th>
                        <th style="border: 1px solid black; padding: 8px; text-align: left;">Amostra</th>
                        <th style="border: 1px solid black; padding: 8px; text-align: left;">Cliente</th>
                        <th style="border: 1px solid black; padding: 8px; text-align: left;">Prioridade</th>
                    </tr>
                </thead>
                <tbody>
    """ 
        for status_os, amostra, cliente, prioridade in dados:
            html_content += f"""
        <tr>
            <td style="border: 1px solid black; padding: 8px;">{status_os}</td>
            <td style="border: 1px solid black; padding: 8px;">{amostra}</td>
            <td style="border: 1px solid black; padding: 8px;">{cliente}</td>
            <td style="border: 1px solid black; padding: 8px;">{prioridade}</td>
        </tr>
    """
            
        html_content += """ 
        </tbody>
            </table>
            <p>Atenciosamente,<br>Sua Automação</p>
        </body>
    </html>
    """
        return html_content

    @classmethod
    def enviar_email(cls, dados: list):
        # Configurações do servidor SMTP
        servidor = os.getenv('SERVIDOR_ZOHO')
        porta = os.getenv('PORTA')
        usuario = os.getenv('USUARIO')
        senha = os.getenv('SENHA')

        # Destinatário, assunto e copor do e-mail
        # obtendo a data de hoje
        data_hoje = date.today()
        data_hoje_str = date.strftime(data_hoje, '%d-%m-%Y')
        remetente = usuario
        assunto = f'Liberações do dia {data_hoje_str}'
        # criando corpo html
        corpo_html =  cls.criar_corpo_html(dados)
        corpo_str = cls.criar_corpo_str(dados)

        # Cria a mensagem de e-mail
        msg = MIMEMultipart('alternative')
        msg['Subject'] = assunto
        msg['From'] = remetente
        msg['To'] = 'rayara@qualylab.com.br, gestaolab@qualylab.com.br, ti@grupoqualityambiental.com.br'

        msg.attach(MIMEText(corpo_html, 'html'))
        msg.attach(MIMEText(corpo_str, 'plain'))

        # Conecta ao servidor SMTP
        try:
            with smtplib.SMTP(servidor, porta) as server:
                server.starttls()
                server.login(usuario, senha) 
                server.send_message(msg)
                print('E-mail enviado com sucesso!')
        except Exception as e:
            print(f'Erro ao enviar e-mail: {e}')