from dataxl.inner_engine.get import standard_name


def excel(file_xlsx=standard_name, template=None, intelligence: bool = False):
    from dataxl.inner_engine.applications import DataAnalysis
    global file_name
    file_name = DataAnalysis(file_xlsx).file_name
    return DataAnalysis(file_name, template, intelligence)
