import streamlit as st
import pandas as pd
import json
from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import tempfile
import os


def extrair_produtos_nf(chave):
    options = webdriver.ChromeOptions()
    options.add_argument("--headless=new")                  # Headless moderno
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    options.add_argument('--disable-blink-features=AutomationControlled')
    options.add_argument(
        '--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')
    driver = webdriver.Chrome(options=options)

    driver.get(
        "https://consultadfe.fazenda.rj.gov.br/consultaDFe/paginas/consultaChaveAcesso.faces")
    input_chave = driver.find_element(By.ID, "chaveAcesso")
    input_chave.send_keys(chave)
    driver.find_element(By.XPATH, "//input[@value='Consultar']").click()
    time.sleep(5)
    driver.find_element(By.NAME, "j_idt16").click()
    driver.find_element(By.ID, "tab_3").click()
    tabelas = driver.find_elements(
        By.XPATH, '//*[@id="Prod"]/fieldset/div/table')

    produtos = []
    for i in range(1, len(tabelas), 2):
        if i + 1 < len(tabelas):
            tabela_produto = tabelas[i]
            tabela_info = tabelas[i + 1]
            linha = tabela_produto.find_element(By.XPATH, ".//tbody/tr")
            numero = linha.find_element(
                By.CLASS_NAME, "fixo-prod-serv-numero").text.strip()
            descricao = linha.find_element(
                By.CLASS_NAME, "fixo-prod-serv-descricao").text.strip()
            quantidade = linha.find_element(
                By.CLASS_NAME, "fixo-prod-serv-qtd").text.strip()
            unidade = linha.find_element(
                By.CLASS_NAME, "fixo-prod-serv-uc").text.strip()
            valor = linha.find_element(
                By.CLASS_NAME, "fixo-prod-serv-vb").text.strip()
            try:
                info_adicional = tabela_info.find_element(
                    By.XPATH, './/tbody/tr/td/table[2]/tbody/tr[3]/td[1]/span')
                info_adicional = info_adicional.get_attribute('innerHTML')
            except Exception:
                info_adicional = "Não encontrado"
            produto = {
                "numero": numero,
                "descricao": descricao,
                "quantidade": quantidade,
                "unidade_comercial": unidade,
                "valor": valor,
                "codigo_barras": info_adicional
            }
            produtos.append(produto)

    driver.quit()
    return produtos


st.title('Consulta Tabela de Produtos NF-e')

# Campo para o usuário
chave = st.text_input('Digite a chave de acesso da NF-e (44 dígitos):')

if st.button('Consultar e gerar Excel'):
    if len(chave) != 44 or not chave.isdigit():
        st.error('Chave deve ter 44 dígitos numéricos.')
    else:
        with st.spinner('Consultando NF-e e extraindo dados...'):
            produtos = extrair_produtos_nf(chave)
        if produtos:
            df = pd.DataFrame(produtos)
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
                df.to_excel(tmp.name, index=False)
                tmp.seek(0)
                st.success(f"{len(produtos)} produtos extraídos!")
                st.download_button(
                    label="Baixar Excel",
                    data=tmp.read(),
                    file_name="produtos_nf.xlsx",
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
            os.unlink(tmp.name)
        else:
            st.warning('Nenhum produto encontrado.')
