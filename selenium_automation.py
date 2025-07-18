import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select 
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import WebDriverException, TimeoutException, NoSuchElementException
from selenium.webdriver.remote.webelement import WebElement

class SeleniumHandler:
    def __init__(self):
        self.driver = None
        self.wait = None

    def initialize_driver(self) -> bool:
        try:
            options = webdriver.ChromeOptions()
            self.driver = webdriver.Chrome(options=options)
            self.wait = WebDriverWait(self.driver, 30)
            return True
        except WebDriverException as e:
            print(f"Erro ao inicializar o driver: {e}")
            return False
        except Exception as e:
            print(f"Erro inesperado ao inicializar o driver: {e}")
            return False

    def _login(self, email: str, password: str, mes, ano) -> bool:
        try:
            self.initialize_driver()
            aba_original = self.driver.current_window_handle
            self.driver.get("http://financeiro.postalis.org.br/ControleAcesso/login/Login.aspx")
            self.wait.until(EC.presence_of_element_located((By.ID, "MainContent_lgnLogin_UserName"))).send_keys(email)
            self.driver.find_element(By.ID, "MainContent_lgnLogin_Password").send_keys(password)
            self.driver.find_element(By.ID, "MainContent_lgnLogin_LoginButton").click()
            self.wait.until(EC.url_contains("Default.aspx"))
            
            self.wait.until(EC.element_to_be_clickable((By.ID, "MainContent_imbOrcamento"))).click()
            time.sleep(2)
            self.wait.until(EC.element_to_be_clickable((By.ID, "MainContent_imbOrcamento"))).click() 

            self.wait.until(EC.element_to_be_clickable((By.ID, "MenuContent_tvwMenut55"))).click()
            self.wait.until(EC.element_to_be_clickable((By.ID, "MenuContent_tvwMenut65"))).click()
            
            self.wait.until(EC.presence_of_element_located((By.ID, "MenuContent_tvwMenut67"))).click()
            
            self.wait.until(EC.presence_of_element_located((By.ID, "MainContent_MainContent_dbData"))).click() 
            for i in range(6):
                self.wait.until(EC.presence_of_element_located((By.ID, "MainContent_MainContent_dbData"))).send_keys(Keys.BACKSPACE)
            self.wait.until(EC.presence_of_element_located((By.ID, "MainContent_MainContent_dbData"))).send_keys(f"{mes}{ano}")
            time.sleep(10)

            dropdown_versao = self.wait.until(EC.presence_of_element_located((By.ID, "MainContent_MainContent_ddlVersao")))
            select = Select(dropdown_versao)
            select.select_by_visible_text("Exercício: 2025; Versão: 1; Orçamento 2025")
            time.sleep(15)

            dropdown_balancete = self.wait.until(EC.presence_of_element_located((By.ID, "MainContent_MainContent_ddlBalancete")))
            select = Select(dropdown_balancete)
            select.select_by_visible_text("96-PLANO DE GESTÃO ADMINISTRATIVA")
            time.sleep(10)

            self.wait.until(EC.presence_of_element_located((By.ID, "MainContent_MainContent_rblTipoRealizacao_0"))).click() 
            time.sleep(1)

            self.wait.until(EC.presence_of_element_located((By.ID, "MainContent_MainContent_btnImprimir"))).click() 
            time.sleep(20) 

            print("Antes da troca de aba")
            todas_as_abas = self.driver.window_handles
            nova_aba = None
            for handle in todas_as_abas:
                if handle != aba_original:
                    nova_aba = handle
                    break

            if nova_aba:
                self.driver.switch_to.window(nova_aba)
                print(f"Foco mudado para a nova aba: {nova_aba}")
                print(f"Título da nova aba: {self.driver.title}")
                
                dropdown_trigger = self.wait.until(EC.element_to_be_clickable((By.ID, "rtbControles_Menu_ITCNT11_SaveFormat_I")))
                dropdown_trigger.click()
                time.sleep(1)

                xlsx_option = self.wait.until(EC.element_to_be_clickable((By.XPATH, "//td[text()='Excel sem formatação (*.xlsx)']")))
                xlsx_option.click()
                time.sleep(2)

                self.wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/form/div[3]/div[2]/div[1]/table/tbody/tr/td/table/tbody/tr/td[34]'))).click() 
                time.sleep(15)
                self.driver.close()
                print("Nova aba fechada.")

            else:
                print("Nenhuma nova aba foi detectada.")

            return True
        except Exception as e:
            print(f"Erro: {e}")
            return False

    def quit_driver(self):
        if self.driver:
            self.driver.quit()
            self.driver = None
            self.wait = None