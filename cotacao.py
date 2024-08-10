from base import *


urls = 'https://www.remessaonline.com.br/cotacao/cotacao-dolar'


driver, wait = iniciar_driver()
driver.get(urls)
sl(2)

valor_dolar_atual = wait.until(CondicaoExperada.visibility_of_all_elements_located((By.XPATH,'//input[@class="style__Input-sc-zzw8vh-0 gITcsK"]')))
valor_dolar = valor_dolar_atual[1].get_attribute('value')
valor = str(valor_dolar)

wait.until(CondicaoExperada.element_to_be_clickable((By.XPATH, '//button[@data-testid="quotation-page:chart-cta"]'))).click()
sl(1)

driver.execute_script('window.scrollTo(0, 300)')
sl(1)

driver.save_screenshot('Grafico.png')
sl(1)

data = data_atual()
sl(1)

os.system('cls')
titulo = 'Monitoramento de Cambio'
site = urls
grafico = 'Grafico.png'
print(f"\033[;33mCriando arquivo Word \033[1;36m{titulo}\033[;m")
criar_arquivo_word(titulo, valor, data, grafico, site)
sl(2)
print(f"\033[1;36m{titulo}\033[;33m criado com sussesso\033[;m")
sl(2)
print(f'\033[;32mConvertendo o documento Word em PDF\033[;m')

word_path = "teste.docx"
pdf_path = "file.pdf"
convert_pdf(word_path, pdf_path)
sl(1)

print('\033[;31mFim do programa')
