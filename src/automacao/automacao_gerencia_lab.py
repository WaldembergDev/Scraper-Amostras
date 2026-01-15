import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from datetime import datetime, date, timedelta
from dotenv import load_dotenv
from selenium.webdriver.common.keys import Keys
from src.service.enviar_email import EnviarEmail
from src.service.enviar_whatsapp import EnviarWhatsapp
import os
import time

# carregando as variáveis de ambiente
load_dotenv()

destinatario_whatsapp = os.getenv('DESTINATARIO_WHATSAPP')

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
        hoje_menos_cinco = hoje - timedelta(days=5)
        hoje_mais_cinco = hoje + timedelta(days=5)
        # obtendo os dados em formato str
        hoje_menos_cinco_str = date.strftime(hoje_menos_cinco, '%d/%m/%Y')
        hoje_mais_cinco_str = date.strftime(hoje_mais_cinco, '%d/%m/%Y')
        # convertendo para o formato do gerencia
        periodo = f'{hoje_menos_cinco_str} - {hoje_mais_cinco_str}'
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
    def retornar_periodo_coleta_ar(cls):
        hoje = date.today()
        uma_semana_atras = hoje - timedelta(days=7)
        uma_semana_atras_mais_um = uma_semana_atras + timedelta(days=1)
        uma_semana_atras_str = date.strftime(uma_semana_atras, '%d/%m/%Y')
        uma_semana_atras_mais_um_str = date.strftime(uma_semana_atras_mais_um, '%d/%m/%Y')
        periodo = f'{uma_semana_atras_str} - {uma_semana_atras_mais_um_str}'
        return periodo

        
    
    @classmethod
    def obter_driver(cls):
        try:
            nome_sistema = os.name
            # verificando se é Windows
            if nome_sistema == 'nt':
                # obtendo o usuário logado
                usuario = os.getlogin()
                chrome_options = Options()
                chrome_options.add_argument(
                    f'C:/Users/{usuario}/AppData/Local/Google/Chrome/Selenium'
                )
                service = Service(ChromeDriverManager().install())
                driver = webdriver.Chrome(service=service, options=chrome_options)
            else:
                driver = webdriver.Chrome()

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
    def aplicar_configuracoes(cls, driver, periodo=None, tipo_periodo=None):
        colunas = ['Ordem Serviço', 'Status O.S', 'Status Amostra', 'Referência', 'Prioridade', 'Cliente', 'Solicitante']
        if tipo_periodo is None:
            colunas.append('Data de Entrega')
        else:
            colunas.append(tipo_periodo)
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
            # obtendo o campo da data
            if 'Data da Coleta' in colunas:
                # caso seja a data de coleta
                campo_data = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="dataPrazoColetaOSForm"]'))
                    )
            elif 'Data de Entrega' in colunas:
                # caso seja a data da entrega
                campo_data = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="dataPrazoEntregaOSForm"]'))
                    )
            else:
                # caso seja a data da entrega
                campo_data = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="dataPrazoEntradaOSForm"]'))
                    )
            # preenchendo o campo data
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
    def obter_lista_os(cls, driver):
        lista_os = []
        while True:
            # Esperar a tabela carregar para evitar StaleElement
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, '#tableOrdemdeServico tbody tr'))
            )
            
            lista_elementos = driver.find_elements(By.CSS_SELECTOR, '#tableOrdemdeServico tr.even, #tableOrdemdeServico tr.odd')
            
            # Verifica se não há registros
            if len(lista_elementos) > 0:
                primeira_coluna = lista_elementos[0].find_elements(By.TAG_NAME, 'td')
                if len(primeira_coluna) > 0 and primeira_coluna[0].text == 'Nenhum registro encontrado':
                    print('Nenhum registro encontrado na listagem inicial.')
                    break

            for linha in lista_elementos:
                colunas = linha.find_elements(By.TAG_NAME, 'td')
                if colunas:
                    os_num = colunas[0].text
                    if os_num not in lista_os:
                        lista_os.append(os_num)
            
            # Paginação
            elementos_de_navegacao = driver.find_elements(By.XPATH, "//li[contains(@class, 'paginate_button ')]")
            # Lógica simples de paginação baseada no seu código (ajuste se necessário)
            if len(elementos_de_navegacao) <= 3: 
                break
            
            # Tenta ir para próxima página
            try:
                next_btn = driver.find_element(By.XPATH, "//li[@class='paginate_button page-item next']//a")
                next_btn.click()
                time.sleep(3)
            except:
                break # Não tem botão next ou falhou
        return lista_os
        


    @classmethod
    def obter_dados(cls, driver, cliente_excluido=None, cliente_selecionado=None):
        try:
            amostras = []
            while True:
                lista_elementos = driver.find_elements(By.CSS_SELECTOR, '#tableOrdemdeServico tr.even, #tableOrdemdeServico tr.odd')
                if (len(lista_elementos) == 1) and (lista_elementos[0].text == 'Nenhum registro encontrado'):
                            print('Nenhum registro encontrado')
                            break
                ignorar_status_amostra = ['Concluído', 'Cancelada', 'Aguardando']
                for linha in lista_elementos:
                    status_amostra = linha.find_elements(By.TAG_NAME, 'td')[3].text.strip()
                    if status_amostra in ignorar_status_amostra:
                        continue
                    solicitante = linha.find_elements(By.TAG_NAME, 'td')[5].text
                    cliente = linha.find_elements(By.TAG_NAME, 'td')[6].text
                    data_entrega = linha.find_elements(By.TAG_NAME, 'td')[7].text[0:10]
                    if 'Gerencialab' in cliente:
                        continue
                    if cliente_selecionado is not None and cliente_selecionado != cliente:
                            continue
                    if cliente_excluido is not None and cliente_excluido == cliente:
                        continue
                    if cliente == 'PIER MAUA S/A ( )':
                        formato_string = '%d/%m/%Y'
                        data_convertida = datetime.strptime(data_entrega, formato_string).date()
                        data_convertida_mais_dois = data_convertida + timedelta(days=2)
                        data_entrega = datetime.strftime(data_convertida_mais_dois, '%d/%m/%Y')
                    amostra = linha.find_elements(By.TAG_NAME, 'td')[2].text
                    amostras.append((status_amostra, amostra, solicitante, cliente, data_entrega))
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
            dados = cls.obter_dados(driver, cliente_selecionado=None, cliente_excluido='PIER MAUA S/A ( )')
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
            cls.aplicar_configuracoes(driver, periodo=cls.retornar_periodo_pier(), tipo_periodo='Data da Coleta')
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
    

    @classmethod
    def iniciar_automacao_pier_whatsapp(cls):
        driver = None
        try:
            # iniciando o driver
            driver = cls.obter_driver()
            print('Driver Inicializado')

            # Realizando login
            cls.logar(driver)
            print('Logado')

            # Aplicando as configurações
            cls.aplicar_configuracoes(driver, periodo=cls.retornar_periodo_pier(), tipo_periodo='Data da Coleta')
            print('Aplicado filtros e configurações')

            # Obtendo os dados
            dados = cls.obter_dados(driver, cliente_selecionado='PIER MAUA S/A ( )')
            print('Dados obtidos')
            
            # enviando os dados por whatsapp
            whatsapp = EnviarWhatsapp()

            if dados:
                mensagem = f'*Notificação Pier*: \nQuantidade de amostras a serem liberadas: {len(dados)}\nRelação da(s) amostra(s):'
                for amostra in dados:
                    mensagem += f"\n\nStatus Amostra - {amostra[0]}\nAmostra - {amostra[1]}\nCliente - {amostra[3]}\nData de entrega - {amostra[4]}\n\n"
            else:
                mensagem = "*Notificação Pier*:\nNão há laudo a ser liberado"
            whatsapp.enviar_mensagem(destinatario_whatsapp, mensagem)
            
            print('Automação finalizada!')
        except Exception as e:
            print(f'Erro: {e}')
        finally:
            if driver:
                cls.sair_sistema(driver)
    

    @classmethod
    def obter_dados_ar(cls, driver, lista_os):
        try:
            amostras = []

            # 2. Iterar sobre cada OS coletada
            driver.get('https://qualylab.gerencialab.com.br/service-order')
            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//h4[text()="Ordem serviços"]')))
            time.sleep(2)

            for os_num in lista_os:
                try:
                    # Limpar pesquisa anterior
                    driver.find_element(By.XPATH, '//span[text()="Limpar Pesquisa"]').click()
                    time.sleep(1)
                    
                    # Pesquisar a OS específica
                    input_os = driver.find_element(By.XPATH, '//thead//th[contains(text(),"Ordem Serviço ")]//input')
                    input_os.clear()
                    input_os.send_keys(os_num)
                    input_os.send_keys(Keys.ENTER)
                    time.sleep(2)

                    # Pegar a linha resultante
                    linhas = driver.find_elements(By.XPATH, '//table[@id="tableOrdemdeServico"]//tbody//tr')
                    if not linhas:
                        continue

                    cols = linhas[0].find_elements(By.TAG_NAME, 'td')
                    if len(cols) < 12: 
                        continue

                    # Coletar dados iniciais da tabela principal
                    data_entrada_str = cols[12].text[0:10]
                    if data_entrada_str == 'Aguardando':
                        continue
                    
                    numero_amostra = cols[5].text
                    solicitante = cols[8].text
                    status_amostra = cols[6].text
                    cliente = cols[9].text
                    
                    # Entrar nos detalhes
                    linhas[0].click()
                    driver.find_element(By.XPATH, '//div[@class="btn-group"]//span[text()="Visualizar"]').click()
                    time.sleep(1)
                    driver.find_element(By.XPATH, '//div[@class="dt-button-collection dropdown-menu"]//a[1]').click() # Visualizar O.S
                    time.sleep(3)

                    # Aumentar visualização da tabela interna
                    try:
                        driver.find_element(By.XPATH, '//div[@class="dataTables_length"]//select[@name="tableListaOrdemServico_length"]').send_keys('200')
                        time.sleep(2)
                    except:
                        pass

                    # Varrer tabela interna para achar Grupo == 'Ar'
                    lista_amostras_interna = driver.find_elements(By.XPATH, '//table[@id="tableListaOrdemServico"]//tbody//tr')
                    
                    for row_amostra in lista_amostras_interna:
                        cols_interna = row_amostra.find_elements(By.TAG_NAME, 'td')
                        if len(cols_interna) > 3:
                            amostra_grupo = cols_interna[3].text
                            
                            if amostra_grupo == 'Ar':
                                data_entrada = datetime.strptime(data_entrada_str, '%d/%m/%Y')
                                # data_entrega = data_entrada + timedelta(days=7)
                                data_entrada_str = data_entrada.strftime('%d/%m/%Y')
                                
                                amostras.append((status_amostra, numero_amostra, solicitante, cliente, data_entrada_str))
                    
                    # Voltar para a lista principal
                    driver.find_element(By.XPATH, '//button//span[text()="Voltar"]').click()
                    time.sleep(2)

                except Exception as e_loop:
                    print(f"Erro ao processar OS {os_num}: {e_loop}")
                    # Tenta voltar para garantir que o loop continue
                    driver.get('https://qualylab.gerencialab.com.br/service-order')
                    time.sleep(3)

            return amostras

        except Exception as e:
            raise RuntimeError(f'Erro ao obter dados de Análise de Ar: {e}')

    @classmethod
    def iniciar_automacao_analise_ar(cls):
        driver = None
        try:
            print('--- Iniciando Automação Análise de Ar ---')
            driver = cls.obter_driver()
            print('Driver Inicializado')

            cls.logar(driver)
            print('Logado')

            periodo = cls.retornar_periodo_coleta_ar()
            cls.aplicar_configuracoes(driver, periodo=periodo, tipo_periodo='Data da Entrada')
            print(f'Filtros aplicados. Período: {periodo}')

            # obtendo a lista de os
            lista_os = cls.obter_lista_os(driver)

            # obtendo os dados
            dados = cls.obter_dados_ar(driver, lista_os)
            print(f'Dados obtidos: {len(dados)} registros de Ar encontrados.')

            # Envio de e-mail
            complemento = 'Análise de Ar'
            lista_emails_ar = ['rayara@qualylab.com.br',
                               'relatorios@grupoqualityambiental.com.br',
                               'gestaolab@qualylab.com.br']
            EnviarEmail.enviar_email(dados, 
                complemento=complemento, 
                destinatarios=lista_emails_ar,
                tipo_relatorio='analise_ar'
                )
            
            print('Automação de Ar finalizada com sucesso!')

        except Exception as e:
            print(f'Erro na automação de Ar: {e}')
        finally:
            if driver:
                cls.sair_sistema(driver)