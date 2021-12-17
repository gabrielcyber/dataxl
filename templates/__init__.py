from .login import *


def standard_template(exc, template: str, intelligence=False, get=False):
    template = template.lower() if template is not None else template

    if template == 'login':
        if get:
            return login
        else:
            return Login(exc, intelligence)
