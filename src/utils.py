import re

def column_to_number(column_string:str) -> int:
  "Função que transforma uma string de letras de coluna em um índice de coluna (base 0)"
  x = 0
  for i in range(len(column_string)):
    x += 26**(len(column_string) - i - 1)*(ord(column_string[i])-64)
  return x-1

def number_to_column(column_int:int) -> str:
  "Função que transforma um índice de coluna base 0 em uma string de letras de coluna"
  start_index = 0   #  it can start either at 0 or at 1
  letter = ''
  while column_int > 25 + start_index:   
    letter += chr(65 + int((column_int-start_index)/26) - 1)
    column_int = column_int - (int((column_int-start_index)/26))*26
  letter += chr(65 - start_index + (int(column_int)))
  return letter

def xrange_to_grid_range(xrange:str) -> dict:
  "Função que transforma um range do tipo A1:B2 em um dicionário grid_range com índice base 0 exclusivo"
  if not validate_xrange(xrange):
    print("Erro ao converter xrange. Xrange inválido.")
    return None
  columns = re.findall(r'[A-Z]+', xrange) + [0,0]      # Colocando valores extras para não dar xabu
  lines = re.findall(r'[0-9]+', xrange) + [0,0]        # Colocando valores extras para não dar xabu

  start_column = column_to_number(columns[0])
  end_column = column_to_number(columns[1])+1          # Offset para converter de range inclusivo para exclusivo.
  start_row = int(lines[0])-1                          # Offset do base 0
  end_row = int(lines[1])                              # Como é exclusivo, mantém linha original.
    
  grid_range = {"startRowIndex": start_row,
    "endRowIndex": end_row,
    "startColumnIndex": start_column,
    "endColumnIndex": end_column}

  return grid_range

def grid_range_to_xrange(grid_range:dict) -> str:
    if not validate_grid_range(grid_range):
        print('Erro ao converter grid range para xrange: grid range inválido.')
        return ""
    first_row = grid_range['startRowIndex']+1                           # Offset de um pela base 0
    second_row = grid_range['endRowIndex']
    first_column = number_to_column(grid_range['startColumnIndex'])
    second_column = number_to_column(grid_range['endColumnIndex']-1)    # Offset de um pelo range exclusivo.
    return f'{first_column}{first_row}:{second_column}{second_row}'

def validate_xrange(xrange_str: str):
    "Função que valida ranges do formato do excel (A1:B2). São aceitos no formato A1:B2, A1, A:A, 1:1"
    # Algo importante é que eu pressuponho que o segundo range é maior que o primeiro. Não aceito B2:A1 p.e
    xrange_str = xrange_str.upper().strip()
    
    if '!' in xrange_str:
        xrange_str = xrange_str.split('!')[-1]

    patterns = {
        "full": r"^(?P<col1>[A-Z]+)(?P<row1>[1-9][0-9]*):(?P<col2>[A-Z]+)(?P<row2>[1-9][0-9]*)$",
        "single": r"^[A-Z]+[1-9][0-9]*$",
        "col_only": r"^(?P<col1>[A-Z]+):(?P<col2>[A-Z]+)$",
        "row_only": r"^(?P<row1>[1-9][0-9]*):(?P<row2>[1-9][0-9]*)$"
    }

    # 1. Caso: Célula Única (A1)
    if re.fullmatch(patterns["single"], xrange_str):
        return True

    # 2. Caso: Range Completa (A1:B2)
    match = re.fullmatch(patterns["full"], xrange_str)
    if match:
        d = match.groupdict()
        # Validação de ordem (opcional, dependendo da sua necessidade)
        row_ok = int(d['row1']) <= int(d['row2'])
        col_ok = column_to_number(d['col1']) <= column_to_number(d['col2'])
        return row_ok and col_ok

    # 3. Caso: Colunas (A:B)
    match = re.fullmatch(patterns["col_only"], xrange_str)
    if match:
        d = match.groupdict()
        return column_to_number(d['col1']) <= column_to_number(d['col2'])

    # 4. Caso: Linhas (1:10)
    match = re.fullmatch(patterns["row_only"], xrange_str)
    if match:
        d = match.groupdict()
        return int(d['row1']) <= int(d['row2'])

    return False

def validate_grid_range(grid_range:dict, expect_sheet_id = False) -> bool:
  "Função que valida um objeto do tipo grid range."
  # Schema esperado: gridRange = {
  #   "sheetId": int[opcional],
  #   "startRowIndex": int,
  #   "endRowIndex": int,
  #   "startColumnIndex": int,
  #   "endColumnIndex": int
  # }

  if not isinstance(grid_range, dict):
    print('Grid Range Object is not a dictionary.')
  
  expected_fields = {"startRowIndex","endRowIndex","startColumnIndex","endColumnIndex"}
  if expect_sheet_id:
    expected_fields.add('sheetId')

  grid_range_keys = set(grid_range.keys())

  if grid_range_keys != expected_fields:
    print(f"Unexpected keys in Grid Range. Expected: {expected_fields}. Actual: {grid_range_keys}.")
    return False
  
  # Checando se os índices são inteiros.
  for key, value in grid_range.items():
    if key != 'sheetId' and not isinstance(value, int):
      print(f'Unexpected value for {key}. Expected int. Got {type(value)}.')
      return False
    elif value < 0:
      print(f'{key} value out of bounds: {value}.')
  
  # Checando se os índices fazem sentido.
  if grid_range['startRowIndex'] >= grid_range['endRowIndex']:
    print(f'Index error: start row index overlaps end index.')
    return False
  elif grid_range['startColumnIndex'] >= grid_range['endColumnIndex']:
    print(f'Index error: start column index overlaps end index.')
    return False

  return True
