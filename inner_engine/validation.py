from .get import ascii_lowercase, ascii_uppercase, digits
from .errors import permutation

get_column_index = lambda i: ascii_uppercase.find(i)
get_column_by_index = lambda i: ascii_uppercase[i]


def validation_input(prompt, input_optional_question=None, input_options=None, input_option_true=None, action=None):
    if input_optional_question is not None:
        obj = input(prompt).strip()
        if len(obj) == 0:
            while True:
                ask = input(input_optional_question).lower().strip()
                if ask in [input_options[0], input_options[1]]:
                    if ask == input_options[0]:
                        obj = action()
                        if input_option_true is not None:
                            while True:
                                option_true = input(f'{input_option_true[0]} {obj}'
                                                    f'{input_option_true[1]}').lower().strip()
                                if option_true == input_options[0]:
                                    obj = action()
                                elif option_true == input_options[1]:
                                    return obj
                        return obj
                    elif ask == input_options[1]:
                        break
        return obj
    else:
        while True:
            ask = input(prompt).strip()
            if len(ask) == 0:
                continue
            break
        return ask


def password_generator(length, lowercase=False, uppercase=False, numbers=False, special=False):
    lowercase = ascii_lowercase if lowercase is True else ''
    uppercase = ascii_uppercase if uppercase is True else ''
    numbers = digits if numbers is True else ''
    special = '!@#$%&*()-+_´^~{}:;.,<>' if special is True else ''

    characters = lowercase + uppercase + numbers + special
    return permutation(characters, length)
