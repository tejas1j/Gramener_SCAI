# coding: utf-8
# -*- coding: utf-8 -*-

import pandas as pd

def annote(handler):
    if handler=='slide_1':
       data=pd.read_excel('Services_Revenue_Excl_VBR.xlsx')
       x=data.loc[0]['YoY%']
       print(x)
       y=data['YoY%']
       c=0
       string=""
       j=2
       for i in y[2:7]:
           
           if i>0:
              c=c+1
           else:
              string=string+" "+data.loc[j]['Services_Revenue_Excl_VBR'].lstrip('-')+" ("+str(data.loc[j]['YoY%'])+"% YoY)"
           j=j+1
       print(c)
       if c<5:
         grw = "lagging behind"
       else:
         grw = "is"
       
    return [grw,str(x),str(c),string]


def color(handler, logic_mapping):
   logic = { 
      'insights': {
         'color': { 'red': "#ff00000", 'green': '#228B22' }
         },    
      'YOY': {
         "regions": ["--AM","--AP","--AU","--EU","--MEA"],
         'color': { 'red': "#ff00000", 'green': '#00ff00' }
         }
   }
   print("----------------test----------------")
   print(handler)
   print("----------------test----------------")
   print(logic_mapping)
   color = logic[logic_mapping]['color'] 
   data  = color['red'] if handler < 0 else color['green']

   if logic_mapping == 'YOY':
      # df = ["--AM","--AP","--AU","--EU","--MEA"]
      # df = handler[handler['Services_Revenue_Excl_VBR'].isin(df)]
      print("testing color")
      print(handler)
      print("testing---------")
      data = '#fff'
       
   # handler = handler["Services_Revenue_Excl_VBR"!='Group Services',"Services_Revenue_Excl_VBR"!='Group Services',"Services_Revenue_Excl_VBR"!='Regions'] 
   # print(handler["Services_Revenue_Excl_VBR"!='Group Services'] & "Services_Revenue_Excl_VBR"!='Group Services',"Services_Revenue_Excl_VBR"!='Regions'] )

   return data