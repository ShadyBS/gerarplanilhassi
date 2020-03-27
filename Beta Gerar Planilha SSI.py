#! python3
#Atualizador SSI

#Importando requisitos e configurando webdriver chrome
import re
import pandas as pd
from datetime import datetime
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options as ChromeOptions
opcoes = ChromeOptions()#Opções do Chrome
preferencias = {"profile.managed_default_content_settings.images": 2}
opcoes.headless = False #Chrome em modo silencioso
opcoes.add_experimental_option("prefs", preferencias)
browser = webdriver.Chrome(options=opcoes) #Definindo avegador


#Variaveis
meuUsuarioJIRA = 'renato.rosa' #usuário JIRA
minhaSenhaJIRA = '!japao@FESTA%' #senha JIRA
pagPrincipalJIRA = 'https://jira.mv.com.br/servicedesk/customer/portal/301/user/login' #Página de login SSI

meuUsuarioSSI = 'rkrosa' #usuário SSI
minhaSenhaSSI = 'Metal1837' #senha SSI
pagPrincipalSSI = 'https://ssi.caxias.rs.gov.br/index.php' #lista de SSI

totalSSI = 0

#Garante o carregamento de uma pagina e espera a pagina carregar
def carregarPagina(pagina, seletor, tempo):
	i=0
	while True:
		try:
			browser.get('about:blank')
			browser.get(pagina)
			esperarElemento(seletor, tempo)
			break
		except:
			i=i+1
			if i<=4:
				print('Falha na tentativa ' + str(i) + ' de carregar a pagina. Tentando novamente...')
			else:
				break
#Escrever em caixa de texto
def escreverCaixa(caixa, texto, submit):
	elemento = browser.find_element_by_id(caixa)
	elemento.send_keys(texto)
	if submit:
		elemento.submit()

#Esperar que um elemento na pagina carregue
def esperarElemento(seletor, tempo):
	elemento = WebDriverWait(browser, tempo).until(EC.presence_of_element_located((By.CSS_SELECTOR, seletor)))

#Estágio de login
def loginSSI():
	print('Fazendo login no SSI com o usuário ' + meuUsuarioSSI)
	carregarPagina(pagPrincipalSSI, "#nm_senha", 120)
	escreverCaixa('nm_login', meuUsuarioSSI, False)
	escreverCaixa('nm_senha', minhaSenhaSSI, True)
	esperarElemento("body > table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(2) > div:nth-child(1) > font:nth-child(1)", 120)

#Estágio de login JIRA
def loginJIRA():
	print('Fazendo login no JIRA com o usuário ' + meuUsuarioJIRA)
	carregarPagina(pagPrincipalJIRA, "#os_password", 120)
	escreverCaixa('os_username', meuUsuarioJIRA, False)
	escreverCaixa('os_password', minhaSenhaJIRA, True)
	esperarElemento(".cv-requests-nav__text", 120)

#Encontra a tabela dentro de 'alvo' e enumera suas colunas e células.
def sopaTabela(alvo):
	print('Fazendo download da tabela de SSI fila MV Sistemas')
	carregarPagina(alvo, "#fila_88", 120)
	SSIs = []
	Tecnicos = []
	Situacoes = []
	Dtransfs = []
	sopa = BeautifulSoup(browser.page_source, 'html.parser') #Criando elemento BS
	tabela = sopa.find('tbody', id = "fila_88")
	linhas = tabela.find_all('tr')
	for linha in linhas:
		celulas= linha.find_all('td')
		SSI = celulas[0].text.strip()
		SSIs.append(SSI)
		Tecnico = celulas[4].text.strip()
		Tecnicos.append(Tecnico)
		Situacao = celulas[5].text.strip()
		Situacoes.append(Situacao)
		Dtransf = re.sub(' ..:..', '', str(celulas[6].text.strip()))
		Dtransfs.append(Dtransf)
	global totalSSI
	totalSSI = len(SSIs)
	print('Encontrados ' + str(totalSSI) + ' SSIs')
	global df1
	df1 = pd.DataFrame(
	{
	'SSI': SSIs,
	'Tecnico': Tecnicos,
	'Situacao': Situacoes,
	'Data Transf.': Dtransfs
	}
	)

#Encontra chamados JIRA dentro de cada SSI e coletar situação e versão corrigida do JIRA
def encontrarJIRA():
	print('Entrando em cada SSI e encontrando o chamado JIRA no formato SDMVCAC-XXXXXX')
	JIRAs=[]
	analise=[]
	i=1
	global totalSSI
	for SSI in df1.loc[:, "SSI"]:
		print('Extraindo chamados JIRA do SSI o chamado ' + SSI + '. Tarefa (' + str(i) + '/' + str(totalSSI) + ')')
		try:
			carregarPagina('https://ssi.caxias.rs.gov.br/auxiliar.php?option=abrir_solicitacao&cod_solicitacao='+SSI+'#tabs-7', "#flat2", 120)
			i=i+1
		except:
			i=i+1
			continue

		JIRA = re.findall(r'\b(\w*SDMVCAC.\w*)\b', str(browser.page_source), flags=re.IGNORECASE)
		carregarPagina('https://jira.mv.com.br/issues/?jql=text%20~%20%22'+SSI+'%22', "#issue-content > div > div > p, .focused > a:nth-child(1) > div:nth-child(1) > div:nth-child(2) > span:nth-child(2), .no-results", 60)
		JIRA = JIRA+re.findall(r'\b(\w*SDMVCAC-\w*)\b', str(browser.page_source), flags=re.IGNORECASE)
		if JIRA != []:
			JIRA = [J.replace(' ', '-', re.IGNORECASE) for J in JIRA]
			JIRA = [J.replace('SDMVCAC_', 'SDMVCAC-', re.IGNORECASE) for J in JIRA]
			JIRA = [J.replace('_', '', re.IGNORECASE) for J in JIRA]
			JIRA = list(set(JIRA))
			k=1
			for chamado in JIRA:
				print('└Extraindo dados do chamado JIRA ' + chamado + '. Tarefa (' + str(k) + '/' + str(len(JIRA)) + ')')
				try:
					carregarPagina('https://jira.mv.com.br/projects/SDMVCAC/queues/issue/' + chamado, "#issue-content > div > div > h1, #sd-customer-visible-status, .sd-error-panel-message", 120)
					k=k+1
				except:
					analise = [chamado, "Inválido"]
					k=k+1
					continue
				sopaJIRA = BeautifulSoup(browser.page_source, 'html.parser')
				try:
					situacaoJIRA = sopaJIRA.find('dd', id="sd-customer-visible-status").text.strip()
					versaocorJIRA = sopaJIRA.find('span', id="fixfor-val").text.strip()
				except:
					situacaoJIRA = []
					versaocorJIRA = []
				analise.append(chamado)
				analise.append(situacaoJIRA)
				analise.append(versaocorJIRA)
		else:
			print('└O SSI ' + SSI + ' não possui chamado JIRA. Encerrando análise')
		JIRAs.append(analise)
		analise=[]
		JIRA=[]
	df2 = df1.join(pd.DataFrame(JIRAs))
	df2.to_excel("output.xlsx")

loginJIRA()
loginSSI()
sopaTabela(pagPrincipalSSI)
encontrarJIRA()

#Encerrar navegador apropriadamente ao final da execução.
browser.quit()
print('Tarefa encerrada.')