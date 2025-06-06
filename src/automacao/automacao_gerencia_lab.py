from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from datetime import date, timedelta
from dotenv import load_dotenv
from selenium.webdriver.common.keys import Keys
from src.service.enviar_email import EnviarEmail
import os
import time

# carregando as variáveis de ambiente
load_dotenv()

class AutomacaoAmostras():
    @classmethod
    def retornar_periodo(cls):
        hoje = date.today()
        amanha = hoje + timedelta(days=1)
        hoje_str = date.strftime(hoje, '%d/%m/%Y')
        amanha_str = date.strftime(amanha, '%d/%m/%Y')
        periodo_str = (f'{hoje_str} - {amanha_str}')
        return periodo_str
    
    @classmethod
    def obter_driver(cls):
        try:
            # obtendo o usuário logado
            usuario = os.getlogin()
            chrome_options = Options()
            chrome_options.add_argument(
                f'user-data-dir=C:/Users/{usuario}/AppData/Local/Google/Chrome/Selenium'
            )
            service = Service(ChromeDriverManager().install())
            driver = webdriver.Chrome(service=service, options=chrome_options)

            # maximixando a tela
            driver.maximize_window()
            return driver
        except Exception as e:
            raise RuntimeError(f"Erro ao obter o driver: {e}")
        
    @classmethod
    def logar(cls, driver):
        try:
            # indo para a tela de login
            driver.get("https://qualylab.gerencialab.com.br/")

            # aguardando elemento aparecer na tela
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, '//img[@src="/assets/images/Logo_Gerencialab-azul.png"]'))
                )
            time.sleep(4)

            # obtendo as variáveis de ambiente
            login = os.getenv('LOGIN')
            password = os.getenv('PASSWORD')

            if not login or not password:
                raise ValueError('Variáveis de ambiente LOGIN ou PASSWORD não definidas.')

            # preenchendo o campo usuário
            driver.find_element(By.XPATH, '//*[@id="loginsite"]').send_keys(login)

            # preenchendo o campo senha
            driver.find_element(By.XPATH, '//*[@id="senhasite"]').send_keys(password)

            # clicando em acessar
            driver.find_element(By.XPATH, '//*[@id="authLogin"]').click()

            # aguardando logo GerenciaLab Aparecer na tela
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, '//img[@src="/assets/images/gerencialab-logo-n.png"]'))
                )
            time.sleep(4)

            print('Login realizado com sucesso!')
        except Exception as e:
            raise RuntimeError(f'Erro ao realizar login no sistema: {e}')

    @classmethod
    def aplicar_configuracoes(cls, driver):
        colunas = ['Ordem Serviço', 'Status O.S', 'Referência', 'Prioridade', 'Cliente', 'Data de Entrega']
        try:
            # indo para ordens de serviço
            driver.get("https://qualylab.gerencialab.com.br/service-order")
            # aguardando label Ordem serviços aparecer na tela
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, '//h4[text()="Ordem serviços"]'))
                )
            time.sleep(4)
            # clicando no botão Limpar Pesquisa
            for i in range(2):
                WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, '//span[text()="Limpar Pesquisa"]'))
                    ).click()
                time.sleep(3)
            # clicando em visualizar colunas
            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//button[@aria-controls="tableOrdemdeServico"]//span[text()="Visualizar Colunas"]'))
                ).click()
            elementos = driver.find_elements(By.XPATH, '//div[@class="dt-button-collection dropdown-menu"]//a')
            # indo até a coluna Data de Entrega - posição 15 
            for index, elemento in enumerate(elementos):
                if index > 15:
                    break; 
                coluna = elemento.text
                status = elemento.get_attribute('class')
                if coluna in colunas and not 'active' in status:
                    elemento.click()
                if not coluna in colunas and 'active' in status:
                    elemento.click()
            # fechando a tela das visualizações de colunas
            driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
            time.sleep(3)
            # obtendo o campo da data de entrega
            campo_data = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="dataPrazoEntregaOSForm"]'))
                )
            # preenchendo o campo data de entrega
            campo_data.send_keys(cls.retornar_periodo())
            time.sleep(3)
            # clicando em aplicar
            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//div[@class="drp-buttons"]//button[text()="Aplicar" and not(@disabled)]'))
                ).click()
            time.sleep(3)
            # mostrar 200 resultados por página
            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//div[@class="dataTables_length"]//select[@name="tableOrdemdeServico_length"]'))
                ).send_keys(200)
            time.sleep(3)
        except Exception as e:
            raise RuntimeError(f'Erro ao aplicar filtros: {e}')
    
    @classmethod
    def obter_dados(cls, driver):
        try:        
            # obtendo a lista de elementos
            lista_elementos = driver.find_elements(By.CSS_SELECTOR, '#tableOrdemdeServico tr.even, #tableOrdemdeServico tr.odd')

            if len(lista_elementos) < 1:
                print('Nenhum elemento foi encontrado')
                return []

            amostras = []
            for linha in lista_elementos:
                cliente = linha.find_elements(By.TAG_NAME, 'td')[4].text
                status_os = linha.find_elements(By.TAG_NAME, 'td')[1].text
                amostra = linha.find_elements(By.TAG_NAME, 'td')[2].text
                prioridade = linha.find_elements(By.TAG_NAME, 'td')[3].text
                amostras.append((status_os, amostra, cliente, prioridade))
                
            return amostras
        except Exception as e:
            raise RuntimeError(f'Erro ao obter os dados: {e}')
    
    @classmethod
    def sair_sistema(cls, driver):
        try:
            driver.find_element(By.XPATH, '//a[@href="/sair"]').click()
            time.sleep(4)
            driver.quit()
            return True
        except Exception as e:
            raise RuntimeError(f'Erro ao sair do sistema utilizando o botão "sair"')
        finally:
            driver.quit()
            print('Navegador Fechado')
    
    @classmethod
    def iniciar_automacao(cls):
        driver = None
        try:
            # iniciando o driver
            driver = cls.obter_driver()
            print('Driver Inicializado')

            # Realizando login
            cls.logar(driver)
            print('Logado')

            # Aplicando as configurações
            cls.aplicar_configuracoes(driver)
            print('Aplicado filtros e configurações')

            # Obtendo os dados
            dados = cls.obter_dados(driver)
            print('Dados obtidos')
            
            # enviando os dados por e-mail
            EnviarEmail.enviar_email(dados)
            
            print('Automação finalizada!')
        except Exception as e:
            print(f'Erro: {e}')
        finally:
            if driver:
                cls.sair_sistema(driver)
