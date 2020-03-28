import pandas as pd

'''data frame being used'''

df = pd.read_excel(open('/Users/alexandrubordei/Downloads/Cambro Sales 2019 Nationally-2.xlsx','rb'), sheet_name='Sheet1', usecols="A")


'''print statement'''

with open('output_CAMBRO_VND_Item.txt', 'a+') as f:
    for _, row in df.iterrows():
        item = row
        INVM_syntax = [
                        'type invm',
                        'key enter',
                        'type ' + str(int(item)),
                        'key Enter',
                        'type e',
                        'key Enter',
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
                        'key CursorUp',
                        'key CursorUp',
                        'key CursorUp',
                        'key CursorUp',
                        'key CursorUp',
                        'key CursorUp',
                        'key CursorUp',
                        'key CursorRight',
                        'key CursorRight',
                        'key CursorRight',
                        'key CursorRight',
                        'key CursorRight',
                        'key CursorRight',
                        'key CursorRight',
                        'key CursorRight',
                        'key CursorRight',
                        'key CursorRight',
                        'key CursorRight',
                        'key CursorRight',
                        'key CursorRight',
                        'key CursorRight',
                        'key CursorRight',
                        'key CursorRight',
                        'key CursorRight',
                        'key CursorRight',
                        'key CursorRight',
                        'key CursorRight',
                        'key CursorRight',
                        'key CursorRight',
                        'key CursorRight',
                        'key CursorRight',
                        'key CursorRight',
                        'EditSelect 08,19,08,36',
                        'wait 1000',
                        'key EditCopy',
                        'wait 500',
                        r'FileSpec clipboard,C:\Users\abordei\Desktop\VND_Item_Code_CAMBRO.csv,append',
                        'key EditSaveClipboard', 'key PA2']
        f.write(('\n'.join(INVM_syntax) + '\n'))

