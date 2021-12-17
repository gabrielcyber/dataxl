from .get import choice


def A1_error(cell, workbook_active):
    if cell[0] == '\x41':
        raise ValueError("Coluna A não pode ser usada")
    elif cell[1:] == '\x31':
        raise ValueError("Linha 1 não pode ser usada")


def edge_error(start, finish):
    if finish is not None:
        from get import get_column_index
        if get_column_index(start[0]) > get_column_index(finish[0]):
            raise ValueError(f"Coluna {finish[0]} tem que ser maior que à coluna inicial")
        elif int(start[1:]) > int(finish[1:]):
            raise ValueError(f"Linha {finish[1]} tem que ser maior que à linha inicial")


def values_package(cell, value):
    if value is tuple():
        raise ValueError(f"É necessário informar valores para o pacote da célula {cell}")


def permutation(characters, length):
    try:
        return ''.join(choice(characters) for i in range(int(length)))
    except IndexError:
        raise ValueError('Ao menos um tipo de dado deve estar habilitado')
