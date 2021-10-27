import pandas as pd;
import numpy as np;




def percentage_convert(handler):
    print(handler)
    df = handler
    df = df.round(1)
    # df['Ach%'] = pd.Series([round(val, 2) for val in df['Ach%']], index = df.index)
    # df['Ach%'] = pd.Series(["{0:.2f}%".format(val * 100) for val in df['Ach%']], index = df.index)
    return df

def table3_convert(handler):
    data = handler
    data = data[["Financial HighLights","Q2 FY18","Q2 FY19","YoY%","Ach%","Bgt Gap","PF Gap"]]
    print(data)
    return data

def test(handler):
    print(handler)
    return handler


def bold_color(handler, data, color):
    import pdb; pdb.set_trace()
    return "#000000"