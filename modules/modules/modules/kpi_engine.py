import pandas as pd

def safe_div(numerador, denominador):
    if denominador == 0 or pd.isna(denominador):
        return 0
    return numerador / denominador
