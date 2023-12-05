# Relatório de Vendas por Loja

Este é um script Python que gera um relatório de vendas por loja a partir de um arquivo Excel e envia o relatório por e-mail usando o Outlook.

## Como usar

1. Certifique-se de que você tem as bibliotecas 'pandas' e 'win32com.client' instaladas no seu ambiente Python. Se não, você pode instalá-las usando pip:

```bash
pip install pandas pywin32
```

2. Atualize o endereço de e-mail no script para o endereço de e-mail para o qual você deseja enviar o relatório.

3. Execute o script Python.

## O que o script faz

1. Lê os dados de vendas do arquivo 'Vendas.xlsx'.
2. Calcula o faturamento total e a quantidade total de vendas por loja.
3. Calcula o ticket médio do produto por loja.
4. Envia um e-mail com o relatório de vendas para o endereço de e-mail especificado.

