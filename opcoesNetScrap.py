from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
import time

'''
FUNCIONALIDADES:

--le a tabela do site opcoes.net e os filtra crindo um arquivo .xlsx com os dados selecionados

--Necessario os parametros acao(nome da acao ex:ABEV3), tipo(contendo o tipo de filtragem dos dados) e vencimento(contendo a data de vencimento para pesquisa)

--Necessario o parametro nome para selecionar o nome do arquivo .xlsx a ser criado com os dados
'''

driver = webdriver.Chrome()

def obterDados(acao, tipo, vencimento):
    global df_filtrado, valorAtualstr

    url = f"https://opcoes.net.br/opcoes/bovespa/{acao}"
    driver.get(url)
    time.sleep(5)

    # Desmarca os vencimentos prviamente marcados
    checkbox_container_xpath = '//*[@id="grade-vencimentos-dates"]'
    checkboxes = driver.find_elements(By.XPATH, f"{checkbox_container_xpath}//input[@type='checkbox']")
    for index, checkbox in enumerate(checkboxes, start=1):
        if checkbox.is_selected():
            checkbox.click()

    # Clica no vencimento selecionado
    driver.find_element(By.XPATH, f'//*[@id="v{vencimento}"]').click()
    time.sleep(5)

    # Clica no tipo selecionado (tpTodas, tpCalls, tpPuts)
    driver.find_element(By.XPATH, f'//*[@id="{tipo}"]').click()
    time.sleep(5)

    # Busca o valor atual
    valorAtualstr = driver.find_element(By.XPATH, '//*[@id="divCotacaoAtual"]/span[2]').text
    # Transforma o valorAtual em float
    valorAtual = valorAtualstr.replace("R$", "").replace(",", ".")
    valorAtual = float(valorAtual)
    time.sleep(5)

    # https://iqss.github.io/dss-webscrape/finding-web-elements.html
    tabela = []
    # Busca os dados da tabela
    table_body = driver.find_element(By.XPATH, '//*[@id="tblListaOpc"]')
    entries = table_body.find_elements(By.TAG_NAME, 'tr')

    for i in range(1, len(entries)):
        cols = entries[i].find_elements(By.TAG_NAME, 'td')
        table_row = ''
        for j in range(len(cols)):
            col = cols[j].text
            if col == "":
                col = "-"
            if j == len(cols) - 1:
                table_row = table_row + col
            else:
                table_row = table_row + col + " "
        tabela.append(table_row.split())

    # Busca o cabecalho da tabela
    if tipo == 'tpTodas':
        cabecalho = ['Ticker', 'Tipo', 'F.M.', 'Mod.', 'Strike', 'A/I/OTM', 'Dist. (%) do Strike', 'Último', 'Var. (%)',
                     'Data/Hora', 'Núm. de Neg.', 'Vol. Financeiro',
                     'Vol. Impl. (%)', 'Delta', 'Gamma', 'Theta ($)', 'Theta (%)', 'Vega', 'IQ', 'Coberto', 'Travado',
                     'Descob.', 'Tit.', 'Lanç.']
    else:
        cabecalho = ['Ticker', 'F.M.', 'Mod.', 'Strike', 'A/I/OTM', 'Dist. (%) do Strike', 'Último', 'Var. (%)',
                     'Data/Hora', 'Núm. de Neg.', 'Vol. Financeiro',
                     'Vol. Impl. (%)', 'Delta', 'Gamma', 'Theta ($)', 'Theta (%)', 'Vega', 'IQ', 'Coberto', 'Travado',
                     'Descob.', 'Tit.', 'Lanç.']


    # Criar o DataFrame
    df = pd.DataFrame(tabela, columns=cabecalho)

    # Remove as colunas desnecessarias
    df = df.drop(columns=['Gamma', 'Theta ($)', 'Theta (%)', 'Vega', 'IQ', 'Coberto', 'Travado',
                     'Descob.', 'Tit.', 'Lanç.'])


    # Transforma os valores Strike em float
    df['Strike'] = df['Strike'].str.replace(',', '.').astype(float)

    # Transforma os valores Último em float
    df['Último'] = df['Último'].str.replace(',', '.').replace('-', '0').astype(float)

    # Calculando o resultado do Premio (divisão entre as colunas Último e Strike)
    df['Premio'] = df['Último'] / df['Strike']

    # Obtendo o índice da coluna "Último"
    indice_ultimo = df.columns.get_loc('Último')

    # Inserindo a coluna "Resultado" após a coluna "Último"
    df.insert(indice_ultimo + 1, 'Premio', df.pop('Premio'))

    # Filtra entre os 5 maiores e 5 menores Strike a partir do valorAtual
    df_filtrado = pd.concat([df[df['Strike'] <= valorAtual].nlargest(5, 'Strike'), df[df['Strike'] >= valorAtual].nsmallest(5, 'Strike')])


def montaExcel(nome):
    global df_filtrado, valorAtualstr

    # Cria um arquivo .xlsx com os dados do dataframe
    df_filtrado.to_excel(f'{nome}.xlsx', sheet_name='filto opcoes.net')


# vencimentos ex:    2024-02-23   , 2024-03-01     (consulte as opcoes no site)
# tipos =   tpTodas    ,   tpCalls   ,   tpPuts
obterDados('ABEV3', 'tpTodas', '2024-04-19')

# Escreva o nome do arquivo que desejar
montaExcel('opcoesNet')
