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
        # obtendo os dados em formato date
        hoje = date.today()
        amanha = hoje + timedelta(days=1)
        depois_de_amanha = hoje + timedelta(days=2)
        terca = hoje + timedelta(days=4) 
        # obtendo os dados em formato str
        amanha_str = date.strftime(amanha, '%d/%m/%Y')
        depois_de_amanha_str = date.strftime(depois_de_amanha, '%d/%m/%Y')
        terca_str = date.strftime(terca, '%d/%m/%Y')  
        if hoje.weekday() == 4: # verificando se hoje é sexta
            periodo = f'{amanha_str} - {terca_str}'
        else:
            periodo = f'{amanha_str} - {depois_de_amanha_str}'
        return periodo
    
    @classmethod
    def retornar_periodo_pier(cls):
        # obtendo os dados em formato date
        hoje = date.today()
        amanha = hoje + timedelta(days=1)
        hoje_mais_tres = hoje + timedelta(days=3)
        # obtendo os dados em formato str
        amanha_str = date.strftime(amanha, '%d/%m/%Y')
        hoje_mais_tres_str = date.strftime(hoje_mais_tres, '%d/%m/%Y')
        # convertendo para o formato do gerencia
        periodo = f'{amanha_str} - {hoje_mais_tres_str}'
        return periodo
    
    @classmethod
    def retornar_periodo_atrasados(cls):
        hoje = date.today()
        uma_semana_atras = hoje - timedelta(days=7)
        hoje_str = date.strftime(hoje, '%d/%m/%Y')
        uma_semana_atras_str = date.strftime(uma_semana_atras, '%d/%m/%Y')
        periodo = f'{uma_semana_atras_str} - {hoje_str}'
        return periodo
        
    
    @classmethod
    def obter_driver(cls):
        try:
            # obtendo o usuário logado
            usuario = os.getlogin()
            chrome_options = Options()
            chrome_options.add_argument(
                f'C:/Users/{usuario}/AppData/Local/Google/Chrome/Selenium'
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
    def aplicar_configuracoes(cls, driver, periodo=None):
        colunas = ['Ordem Serviço', 'Status O.S', 'Referência', 'Prioridade', 'Cliente', 'Data de Entrega', 'Solicitante']
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
            campo_data.send_keys(cls.retornar_periodo() if periodo is None else periodo)
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
    def obter_dados(cls, driver, cliente_excluido=None, cliente_selecionado=None):
        try:        
            amostras = []
            while True:
                lista_elementos = driver.find_elements(By.CSS_SELECTOR, '#tableOrdemdeServico tr.even, #tableOrdemdeServico tr.odd')
                if (len(lista_elementos) == 1) and (lista_elementos[0].text == 'Nenhum registro encontrado'):
                            print('Nenhum registro encontrado')
                            break
                lista_status_os = ['Laboratorio', 'Em Revisão', 'Assinatura']
                for linha in lista_elementos:
                    status_os = linha.find_elements(By.TAG_NAME, 'td')[1].text
                    if not status_os in lista_status_os:
                        continue
                    solicitante = linha.find_elements(By.TAG_NAME, 'td')[4].text
                    cliente = linha.find_elements(By.TAG_NAME, 'td')[5].text
                    if 'Gerencialab' in cliente:
                        continue
                    if cliente_selecionado is not None and cliente_selecionado != cliente:
                            continue
                    if cliente_excluido is not None and cliente_excluido == cliente:
                        continue
                    amostra = linha.find_elements(By.TAG_NAME, 'td')[2].text
                    data_entrega = linha.find_elements(By.TAG_NAME, 'td')[6].text[0:10]
                    amostras.append((status_os, amostra, solicitante, cliente, data_entrega))
                # verificando se é possível passar para a próxima página:
                elementos_de_navegacao = driver.find_elements(By.XPATH, "//li[contains(@class, 'paginate_button ')]")
                if len(elementos_de_navegacao) == 3:
                    break
                # passando para a próxima página
                driver.find_element(By.XPATH, "//li[@class='paginate_button page-item next']//a").click()
                time.sleep(3)
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
    def iniciar_automacao_geral(cls):
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
            dados = cls.obter_dados(driver, cliente_excluido='PIER MAUA S/A ( )')
            print('Dados obtidos')
            
            # enviando os dados por e-mail
            complemento = 'Geral'
            EnviarEmail.enviar_email(dados, complemento=complemento)
            
            print('Automação finalizada!')
        except Exception as e:
            print(f'Erro: {e}')
        finally:
            if driver:
                cls.sair_sistema(driver)
    
    @classmethod
    def iniciar_automacao_pier(cls):
        driver = None
        try:
            # iniciando o driver
            driver = cls.obter_driver()
            print('Driver Inicializado')

            # Realizando login
            cls.logar(driver)
            print('Logado')

            # Aplicando as configurações
            cls.aplicar_configuracoes(driver, periodo=cls.retornar_periodo_pier())
            print('Aplicado filtros e configurações')

            # Obtendo os dados
            dados = cls.obter_dados(driver, cliente_selecionado='PIER MAUA S/A ( )')
            print('Dados obtidos')
            
            # enviando os dados por e-mail
            complemento = 'Pier'
            EnviarEmail.enviar_email(dados, complemento=complemento)
            
            print('Automação finalizada!')
        except Exception as e:
            print(f'Erro: {e}')
        finally:
            if driver:
                cls.sair_sistema(driver)
    
    @classmethod
    def iniciar_automacao_atrasados(cls):
        driver = None
        try:
            # iniciando o driver
            driver = cls.obter_driver()
            print('Driver Inicializado')

            # Realizando login
            cls.logar(driver)
            print('Logado')

            # Aplicando as configurações
            cls.aplicar_configuracoes(driver, periodo=cls.retornar_periodo_atrasados())
            print('Aplicado filtros e configurações')

            # Obtendo os dados
            dados = cls.obter_dados(driver)
            print('Dados obtidos')
            
            # enviando os dados por e-mail
            # assunto
            complemento = 'Atrasados'
            EnviarEmail.enviar_email(dados, complemento=complemento)
            
            print('Automação finalizada!')
        except Exception as e:
            print(f'Erro: {e}')
        finally:
            if driver:
                cls.sair_sistema(driver)