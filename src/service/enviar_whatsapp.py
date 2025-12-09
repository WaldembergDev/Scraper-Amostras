import requests
import os
from dotenv import load_dotenv
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

load_dotenv()

class EnviarWhatsapp:
    def __init__(self):
        self.token = os.getenv('TOKEN_WHAPI')
    

    def enviar_mensagem(self, numero: str, mensagem: str):

        url = 'https://gate.whapi.cloud/messages/text'

        payload = {
            "typing_time": 0,
            "to": numero,
            "body": mensagem }
        
        headers = {
            "accept": "application/json",
            "content-type": "application/json",
            'Authorization': f'Bearer {self.token}'
        }

        try:
            response = requests.post(url, json=payload, headers=headers)
            response.raise_for_status()
            logging.info(response.text)
            print('teste')
        except Exception as e:
            logging.error(f'Erro: {e}')

