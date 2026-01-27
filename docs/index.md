# Bem-vindo à Documentação de googleSheetsLib

Esta biblioteca serve como um wrapper robusto para a Google Sheets API.
A ideia é facilitar drasticamente a operação com as Spreadsheets usando Python,
abstraíndo a complicação de dominar os endpoints e a sintaxe do Google.

## Instalação

Baixe o arquivo e extraia em alguma pasta de módulos.
Vá até a pasta em que extraiu os arquivos e execute o pip install.

```bash
cd pasta_de_modulo/googleSheetsLib
pip install .
```

Agora é possível importar o módulo normalmente.
```python
from googleSheetsLib import Spreadsheet

# Instanciando um objeto Spreadsheet
ss = Spreadsheet(spreadsheet_id = '1WUBUMIw0fk_dnFO_jnMUTCMS_t_esnYKPYndZIFXhIs',
                 token_fp = 'auth/token.json',
                 cred_fp = 'auth/cred.json')

# Acessando uma aba:
sales = ss['Sales']

# Acessando um range:
values = sales['A1:G12']

# Atualizando uma célula:
sales['B1'] = 'Produto'
```