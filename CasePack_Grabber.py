import pandas as pd

'''data frame being used'''

df = pd.read_excel(open('/Users/alexandrubordei/Downloads/2019 Vendor Product Sales.xlsx','rb'), sheet_name='Sheet1', usecols="A")



'''print statement'''

with open('output_Krowne.txt', 'a+') as f:
    for _, row in df.iterrows():
        item = row
        INVM_syntax = ['type invm', 'key Enter', 'type ' + str(int(item)), 'key Enter', 'key CursorDown', 'key CursorDown', 'key CursorDown',
              'key CursorDown', 'key CursorDown', 'key CursorDown', 'key CursorDown', 'key CursorDown', 'key CursorDown', 'key CursorDown',
              'key CursorDown', 'key CursorLeft', 'key CursorLeft', 'key CursorLeft', 'key CursorLeft', 'EditSelect 16,44,16,46', 'wait 2000','key EditCopy', 'wait 2000', r'FileSpec clipboard,C:\Users\abordei\Desktop\INVM_Krowne.csv,append',
                'key EditSaveClipboard', 'key PA2']
        f.write(('\n'.join(INVM_syntax) + '\n'))