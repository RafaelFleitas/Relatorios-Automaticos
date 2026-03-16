# 📧 Automação de E-mails e Relatórios com Python

Este projeto é um script em Python desenvolvido para automatizar a criação de relatórios individuais de vendas e o envio desses arquivos por e-mail utilizando o Microsoft Outlook. Ele lê uma planilha geral de vendas, calcula os valores totais, separa os dados cliente por cliente, gera planilhas individuais e deixa os e-mails prontos no Outlook com os arquivos anexados.

## 🛠️ Tecnologias Utilizadas
* **Python 3**
* **Pandas:** Para manipulação de dados, cálculos e criação das novas planilhas.
* **PyWin32:** Para integração e automação do Microsoft Outlook no Windows.
* **OpenPyXL:** Biblioteca auxiliar necessária para ler e salvar os arquivos `.xlsx`.

## 📊 Estrutura da Planilha
A planilha principal de dados deve se chamar `base_clientes.xlsx`. Ela **não deve conter fórmulas** e precisa ter a primeira linha configurada como cabeçalho contendo obrigatoriamente as seguintes colunas:
* `Clientes`
* `Produto`
* `Quantidade`
* `Valor Unitário`
* `Valor Total`
* `Email`

## ⚙️ Baixando as Bibliotecas
Para que o código funcione, é necessário instalar as dependências do projeto. Abra o terminal do seu editor (como o VS Code) ou o Prompt de Comando na pasta do projeto e execute o comando abaixo:

```bash
pip install pandas pywin32 openpyxl
