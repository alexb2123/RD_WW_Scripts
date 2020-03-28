import pandas as pd

'''data frame being used'''

df = pd.read_excel(open('/Users/alexandrubordei/Downloads/Cambro Sales 2019 Nationally.xlsx','rb'), sheet_name='Sheet1', usecols="A")



'''print statement'''

with open('output_INVQ_CAMBRO_Retail.txt', 'a+') as f:
    for _, row in df.iterrows():
        item = row
        INVQ_syntax = ['type invq', 'key Enter', 'type ' + str(int(item)), 'key Enter',
                       'key CursorUp',
                       'key CursorUp',
                       'key CursorUp',
                       'key CursorUp',
                       'key CursorUp',
                       'key CursorUp',
                       'key CursorUp',
                       'key CursorUp',
                       'key CursorUp',
                       'key CursorUp',
                       'key CursorLeft',
                       'key CursorLeft',
                       'key CursorLeft',
                       'key CursorLeft',
                       'key CursorLeft',
                       'key CursorLeft',
                       'key CursorLeft',
                       'key CursorLeft',
                       'key CursorLeft',
                       'key CursorLeft',
                       'key CursorLeft',
                       'key CursorLeft',
                       'key CursorLeft',
                       'key CursorLeft',
                       'key CursorLeft',
                       'key CursorLeft',
                       'key CursorLeft',
                       'key CursorLeft',
                       'key CursorLeft',
                       'key CursorLeft',
                       'key CursorLeft',
                       'key CursorLeft',
                       'key CursorLeft',
                       'key CursorLeft',
                       'key CursorLeft',
                       'key CursorLeft',
                       'key CursorLeft',
                       'key CursorLeft',
                       'key CursorLeft',
                       'key CursorLeft',
                       'key CursorLeft',
                       'EditSelect 14,48,14,52',
                       'wait 1000',
                       'key EditCopy',
                       'wait 1000',
                       r'FileSpec clipboard,C:\Users\abordei\Desktop\INVQ_CAMBRO_retail.csv,append',
                       'key EditSaveClipboard',
                       'key clear']
        f.write(('\n'.join(INVQ_syntax) + '\n'))