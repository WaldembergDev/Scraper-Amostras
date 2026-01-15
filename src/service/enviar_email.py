import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from dotenv import load_dotenv
from datetime import date
import os

load_dotenv()

class EnviarEmail:
    
    @classmethod
    def criar_corpo_html_padrao(cls, dados: list):
        """
        Cria o corpo HTML para o relatório padrão (Geral, Pier, Atrasados).
        Espera dados no formato: (status_amostra, amostra, solicitante, cliente, entrega_prevista)
        """
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
                            <th style="border: 1px solid black; padding: 8px; text-align: left;">Status Amostra</th>
                            <th style="border: 1px solid black; padding: 8px; text-align: left;">Amostra</th>
                            <th style="border: 1px solid black; padding: 8px; text-align: left;">Solicitante</th>
                            <th style="border: 1px solid black; padding: 8px; text-align: left;">Cliente</th>
                            <th style="border: 1px solid black; padding: 8px; text-align: left;">Entrega Prevista</th>
                        </tr>
                    </thead>
                    <tbody>
        """ 
        for row in dados:
            # Garante que temos 5 elementos, preenchendo com vazio se faltar
            row = list(row) + [''] * (5 - len(row))
            status_amostra, amostra, solicitante, cliente, entrega_prevista = row[:5]
            
            html_content += f"""
            <tr>
                <td style="border: 1px solid black; padding: 8px;">{status_amostra}</td>
                <td style="border: 1px solid black; padding: 8px;">{amostra}</td>
                <td style="border: 1px solid black; padding: 8px;">{solicitante}</td>
                <td style="border: 1px solid black; padding: 8px;">{cliente}</td>
                <td style="border: 1px solid black; padding: 8px;">{entrega_prevista}</td>
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
    def criar_corpo_html_analise_ar(cls, dados: list):
        """
        Cria o corpo HTML específico para Análise de Ar.
        Espera dados no formato: (status_amostra, numero_amostra, solicitante, cliente, data_entrada)
        """
        html_content = '<p>Nenhum dado encontrado para o período.</p>'
        if not dados:
            return html_content
            
        html_content = """
        <html>
            <head></head>
            <body>
                <p>Prezados,</p>
                <p>Segue a relação das amostras que precisam de relatório de ar:</p>
                <table style="width:100%; border-collapse: collapse;">
                    <thead>
                        <tr style="background-color: #e6f7ff;"> <th style="border: 1px solid black; padding: 8px; text-align: left;">Status Amostra</th>
                            <th style="border: 1px solid black; padding: 8px; text-align: left;">Amostra</th>
                            <th style="border: 1px solid black; padding: 8px; text-align: left;">Solicitante</th>
                            <th style="border: 1px solid black; padding: 8px; text-align: left;">Cliente</th>
                            <th style="border: 1px solid black; padding: 8px; text-align: left;">Data Entrada</th>
                        </tr>
                    </thead>
                    <tbody>
        """ 
        for row in dados:
            # Garante que temos 5 elementos
            row = list(row) + [''] * (5 - len(row))
            status_amostra, amostra, solicitante, cliente, data_entrada = row[:5]
            
            html_content += f"""
            <tr>
                <td style="border: 1px solid black; padding: 8px;">{status_amostra}</td>
                <td style="border: 1px solid black; padding: 8px;">{amostra}</td>
                <td style="border: 1px solid black; padding: 8px;">{solicitante}</td>
                <td style="border: 1px solid black; padding: 8px;">{cliente}</td>
                <td style="border: 1px solid black; padding: 8px;">{data_entrada}</td>
            </tr>
            """
            
        html_content += """ 
                    </tbody>
                </table>
                <p>Atenciosamente,<br>Sua Automação de Ar</p>
            </body>
        </html>
        """
        return html_content

    @classmethod
    def enviar_email(cls, dados: list, complemento=None, destinatarios=None, tipo_relatorio='padrao'):
        """
        Envia e-mail selecionando o corpo adequado e os destinatários.
        
        :param dados: Lista de dados para o relatório.
        :param complemento: Texto para o assunto.
        :param destinatarios: Lista ou string de e-mails.
        :param tipo_relatorio: 'padrao' ou 'analise_ar'.
        """
        # Configurações SMTP
        servidor = os.getenv('SERVIDOR_ZOHO')
        porta = os.getenv('PORTA')
        usuario = os.getenv('USUARIO')
        senha = os.getenv('SENHA')

        # 1. Definição de Destinatários
        if destinatarios is None:
            lista_destinatarios = [
                'rayara@qualylab.com.br', 
                'gestaolab@qualylab.com.br', 
                'ti@grupoqualityambiental.com.br', 
                'financeiro@grupoqualityambiental.com.br', 
                'adm@qualylab.com.br', 
                'servicosanaliticos@qualylab.com.br', 
                'expedicao@qualylab.com.br'
            ]
            destinatarios_str = ", ".join(lista_destinatarios)
        elif isinstance(destinatarios, list):
            destinatarios_str = ", ".join(destinatarios)
        else:
            destinatarios_str = destinatarios

        # 2. Definição do Corpo do E-mail
        if tipo_relatorio == 'analise_ar':
            corpo_html = cls.criar_corpo_html_analise_ar(dados)
            # Corpo texto simples simplificado
            corpo_str = "Prezados,\nSegue a relação das amostras que precisam de relatório de ar.\nVerifique a versão HTML."
        else:
            corpo_html = cls.criar_corpo_html_padrao(dados)
            corpo_str = "Prezados,\nSegue a listagem de amostras que deverão ser liberadas hoje.\nVerifique a versão HTML."

        # 3. Montagem do E-mail
        data_hoje = date.today().strftime('%d-%m-%Y')
        remetente = usuario
        assunto = f'{complemento} - Amostras recebidas em {data_hoje}'

        msg = MIMEMultipart('alternative')
        msg['Subject'] = assunto
        msg['From'] = remetente
        msg['To'] = destinatarios_str
        
        msg.attach(MIMEText(corpo_html, 'html'))
        msg.attach(MIMEText(corpo_str, 'plain'))

        # 4. Envio
        try:
            with smtplib.SMTP(servidor, porta) as server:
                server.starttls()
                server.login(usuario, senha) 
                server.send_message(msg)
                print(f'E-mail ({tipo_relatorio}) enviado com sucesso para: {destinatarios_str}')
        except Exception as e:
            print(f'Erro ao enviar e-mail: {e}')