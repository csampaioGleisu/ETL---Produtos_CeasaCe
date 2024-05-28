# Importação das bibliotecas necessárias
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import pandas as pd
from Functions import formatar_primeira_coluna, preencher_celulas_em_branco, drop_colunm
import os
import time
from datetime import datetime


# Configuração de execução do WebDriver
options = webdriver.ChromeOptions()
options.add_argument("--headless")
driver = webdriver.Chrome(options=options)

# Verifica se o arquivo 'teste_produtos.xlsx' existe e excluí-lo se existir
file_path = os.path.join(os.path.dirname(__file__), 'teste_produtos.xlsx')
if os.path.exists(file_path):
    os.remove(file_path)
    print("Arquivo teste_produtos.xlsx existente excluído.")
    
# Variavél que armazena a data atual
now = datetime.now()

# Variavel que armazena o Ano atual
ano_atual = datetime.now().year

# Variavel que armazena o Mês atual
mes_atual = now.month

# Adiciona o link para o ano atual
Link = {
    str(ano_atual):f'https://files.ceasa-ce.com.br/nuple/principais_produtos/ppmensais-{ano_atual}.html'
} 

# Mapeamento de siglas de meses em Português
meses = {
    1: 'JAN',
    2: 'FEV',
    3: 'MAR',
    4: 'ABR',
    5: 'MAIO',
    6: 'JUN',
    7: 'JUL',
    8: 'AGO',
    9: 'SET',
    10: 'OUT',
    11: 'NOV',
    12: 'DEZ'
}

# Obtém a sigla do mês atual em Português
sigla_mes_atual = meses[mes_atual]

# Cria um novo dicionário months com os meses até o mês atual
months = {mes: str(mes) for mes in range(1, mes_atual)}

# Função que extrai os dados dos principais produtos do site da ceasa
def extract_principais_produtos(Link, months):
    
    df_main_products = pd.DataFrame()
    
    
    for ano, link in Link.items():
        try:
            driver.get(link)
            time.sleep(2)

            driver.switch_to.frame("frTabs")

            for month, xpath in months.items():
                aba_element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, f'/html/body/table/tbody/tr/td[{xpath}]/b/small/small/a'))
                )
                aba_element.click()

                driver.switch_to.default_content()

                frame = driver.find_element(By.NAME, 'frSheet')
                driver.switch_to.frame(frame)

                time.sleep(2)
                html = driver.page_source
                soup = BeautifulSoup(html, 'html.parser')
                table = soup.find('table')
                data = []
                headers = [header.text.strip() for header in table.find_all('th')]
                for row in table.find_all('tr')[1:]:
                    cells = row.find_all('td')
                    row_data = [cell.get_text(strip=True) for cell in cells]
                    data.append(row_data)
                
                dataframe = pd.DataFrame(data,
                            columns=['PRODUTOS', 'PROCEDÊNCIA','Volume (Toneladas)', 'Volume Total ', '(%)', 't', 't'])
                
                dataframe.drop('t', axis=1, inplace=True)

                
                dataframe['ANO'] = ano
                dataframe['MÊS'] = month
                
                df_main_products = pd.concat([df_main_products, dataframe])
                driver.switch_to.default_content()
                driver.switch_to.frame("frTabs")
                
                
        except Exception as e:
            print(f"Erro ao processar o ano {ano} no mês {month}: {e}")

    driver.quit() 
    try:
        df_main_products = df_main_products[df_main_products['PRODUTOS'] != 'P R O D U T O S']
        df_main_products = df_main_products[df_main_products['PRODUTOS'] != 'TONELADAS']
        df_main_products = df_main_products[df_main_products['PRODUTOS'] != 'FRUTAS']
        df_main_products = df_main_products[df_main_products['PRODUTOS'] != '']
        df_main_products = df_main_products[df_main_products['PRODUTOS'] != 'OUTROS EST.']
        df_main_products = df_main_products[df_main_products['PRODUTOS'] != 'HORTALIÇAS']
        df_main_products = df_main_products[df_main_products['PRODUTOS'] != 'T O T A LG E R A L']
        df_main_products = df_main_products[~df_main_products['PRODUTOS'].str.contains('CEASA')]
        df_main_products = df_main_products[~df_main_products['PRODUTOS'].str.contains('Fonte:')]
        df_main_products = df_main_products[~df_main_products['PRODUTOS'].str.contains('OUTROS')]   
        df_main_products = df_main_products[~df_main_products['PRODUTOS'].str.contains('TOTAL')]     
    except Exception as e:
        print(f"Erro na formatação da tabela no ano {ano} no mês {month}")
    return df_main_products


# Executa a função utilizando como parâmetros os dados dos dicionários (Links e months)
df_main = extract_principais_produtos(Link, months)

# Variável filename2 guarda o valor com o nome que será dado ao arquivo xlsx 
filename = 'teste_produtos.xlsx'
#

# Obtém o caminho do diretório atual
current_dir = os.getcwd()

# Constroi o caminho completo para a pasta CeasaCe
ceasa_ce_path = os.path.join(current_dir, 'CeasaCe')

# Constroi o caminho completo para o arquivo
file_path = os.path.join(ceasa_ce_path, 'teste_produtos.xlsx')

# Criar o arquivo xlsx passando os parâmetros de nome do arquivo (que esta salvo na variavel filename2)

df_main.to_excel(filename, index=False, sheet_name='Produtos')
#

# Executa a função que limpa a base removendo linhas desnecessárias
# Parâmetros passados é a (Variavél que armazena o nome do arquivo)
#remover_linhas(filename=filename2)
#

# Executa a função que formata do de uma maneira expecifica a primeira coluna
# Parâmetros passados é a (Variavél que armazena o nome do arquivo)
formatar_primeira_coluna(filename)
#

# Executa a função que preenche as celulas em branco com os dados necessários para que o xlsx fique bem formatado
# Parâmetros passados é a (Variavél que armazena o nome do arquivo)
preencher_celulas_em_branco(filename)

drop_colunm(filename)