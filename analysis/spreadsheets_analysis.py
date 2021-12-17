def sheet(file_xlsx, sheet_name):
    from dataxl.inner_engine.get import pd

    df = pd.read_excel(file_xlsx, sheet_name).drop('#', axis=1).dropna(how='any', axis=1)
    return df
