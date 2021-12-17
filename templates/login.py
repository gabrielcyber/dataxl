DEFAULT_TEXT_COLOR = None
DEFAULT_BACKGROUND_COLOR = 'B9B9B9'
DEFAULT_COLUMN_WIDTH = 9.2
DEFAULT_LINE_HEIGHT = 15


class Login:
    def __init__(self, exc, intelligence):
        cellFormat = lambda cell, value: exc.cell(cell, value, text_size=20, text_font='Arial',
                                                  bold=True, text_color='FFFFFF', background_color='0F243E',
                                                  column_width=20, line_height=30, alignment='center')
        exc.border(6, istemplate=True)

        # Merge to cells
        exc.privateFunction().merge_cells('B2:D2')
        exc.privateFunction().merge_cells('G2:I2')
        exc.privateFunction().merge_cells('L2:N2')

        cellFormat('B2', 'Sites')
        cellFormat('B3', 'Nome')
        cellFormat('C3', 'Email')
        cellFormat('D3', 'Senha')

        cellFormat('G2', 'Aplicativos')
        cellFormat('G3', 'Nome')
        cellFormat('H3', 'Email')
        cellFormat('I3', 'Senha')

        cellFormat('L2', 'Jogos')
        cellFormat('l3', 'Nome')
        cellFormat('M3', 'Email')
        cellFormat('N3', 'Senha')

        if intelligence is True:
            from dataxl.inner_engine.validation import validation_input, password_generator

            def fill_data(category, starting_cell):
                social_media = validation_input(f'Informe o nome do {category}: ')
                account = validation_input('Informe o email da conta: ')
                password = validation_input('Informe a senha da conta: ',
                                            'A senha não foi informada. Deseja gerar uma senha (y/n)? ',
                                            input_options=('y', 'n'),
                                            input_option_true=('Senha gerada: ',
                                                               '. Gostaria de gerar uma nova senha(y/n)? '),
                                            action=lambda:
                                            password_generator(validation_input('Digite o comprimento da senha: '),
                                                               lowercase=True, uppercase=True,
                                                               numbers=True, special=True))
                exc.insert_unlimited(social_media, account, password, starting_cell=starting_cell)

            print(f"{'-'*6}Categorias{'-'*6}")
            for c, e in enumerate(['Sites', 'Aplicativos', 'Jogos']):
                print(f'{c+1} - {e}')
            category_type = validation_input('Selecione uma categoria acima: ')
            if category_type == '1':
                fill_data('Site', 'B3')
            elif category_type == '2':
                fill_data('Aplicativo', 'G3')
            elif category_type == '3':
                fill_data('Jogo', 'L3')
