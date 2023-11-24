# PlanilhaBKO-
# Copiar e Adicionar na Base Cadastro

Este script em Google Apps Script copia dados da planilha "PENDENTES DE CADASTRO" para a planilha "BASE CADASTRO" e adiciona apenas as linhas não copiadas anteriormente.

## Configuração

1. Abra a planilha no Google Sheets.

2. Vá para "Extensões" > "Apps Script".

3. Cole o código fornecido no editor de scripts.

4. Salve o script clicando no ícone de disquete no canto superior esquerdo.

## Uso

1. Na planilha, clique em "Exibir" > "Painel de controle do editor de scripts" para abrir o editor de scripts.

2. Execute a função `copiarEAdicionarNaPagina2` clicando no botão de play ou configurando um gatilho de execução automática.

3. A função copiará as novas linhas da planilha "PENDENTES DE CADASTRO" para a planilha "BASE CADASTRO".

4. As linhas serão adicionadas abaixo das já existentes na "BASE CADASTRO".

## Observações

- Certifique-se de ter as planilhas "PENDENTES DE CADASTRO" e "BASE CADASTRO" com as mesmas colunas e na mesma ordem.

- Ajuste as variáveis `linhasParaCopiar` e `colunasParaCopiar` conforme necessário.

- Se desejar, atribua a função a um botão na planilha para facilitar a execução.

# Funções Google Apps Script

Este conjunto de funções em Google Apps Script é projetado para ser usado em uma planilha específica do Google Sheets. Cada função executa uma ação específica na planilha.

## Funções Disponíveis

### 1. `myFunction`

Ativa a faixa de células 'E3:F7' e muda para a aba '>60 dias'.

### 2. `_60dias`

Ativa a faixa de células 'E9:F13' e muda para a aba '>60 dias'.

### 3. `primeiro`

Ativa a faixa de células 'E9:F13' e muda para a aba '>60 dias'.

### 4. `segundo`

Ativa a faixa de células 'E3:F7' e muda para a aba 'Zerados'.

### 5. `terceiro`

Ativa a célula 'B46' e muda para a aba 'PENDENTES DE CADASTRO'.

### 6. `quarto`

Ativa a faixa de células 'E21:F25' e muda para a aba 'BASE CADASTRO'.

### 7. `quinto`

Ativa a faixa de células 'B27:D31' e muda para a aba '-11 vendas'.

### 8. `sexto`

Ativa a faixa de células 'E33:F37' e muda para a aba 'ATIVAR!'.

### 9. `cadastroBase`

Ativa a célula 'K13' e muda para a aba 'BASE CADASTRO'.

## Como Usar

1. Abra a planilha no Google Sheets.

2. Vá para "Extensões" > "Apps Script".

3. Cole o código fornecido no editor de scripts.

4. Salve o script clicando no ícone de disquete no canto superior esquerdo.

5. Execute qualquer uma das funções conforme necessário.

## Observações

- Certifique-se de ter a planilha e as abas conforme esperado pelo script.

- Ajuste as células ou faixas de células conforme necessário nas funções.

- Se desejar, atribua as funções a botões na planilha para facilitar a execução.

## Autor

[Nicolas e Guilherme]



