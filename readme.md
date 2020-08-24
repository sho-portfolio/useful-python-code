### Write dataframes to multiple Excel worksheets

```python

# from: https://xlsxwriter.readthedocs.io/example_pandas_multiple.html

import pandas as pd

# Create some Pandas dataframes from some data.
df1 = pd.DataFrame({'Data': [11, 12, 13, 14]})
df2 = pd.DataFrame({'Data': [21, 22, 23, 24]})
df3 = pd.DataFrame({'Data': [31, 32, 33, 34]})

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('pandas_multiple.xlsx', engine='xlsxwriter')

# Write each dataframe to a different worksheet.
df1.to_excel(writer, sheet_name='Sheet1')
df2.to_excel(writer, sheet_name='Sheet2')
df3.to_excel(writer, sheet_name='Sheet3')

# Close the Pandas Excel writer and output the Excel file.
writer.save()

```


### Send Outlook Email from Python

```python

# from: https://gist.github.com/ITSecMedia/b45d21224c4ea16bf4a72e2a03f741af

import win32com.client
from win32com.client import Dispatch, constants

pSubject = "My Subject"
pAttachment = r"C:\Temp\example.pdf"
pBody = <br/>my body text</br>

const=win32com.client.constants
olMailItem = 0x0
obj = win32com.client.Dispatch("Outlook.Application")
newMail = obj.CreateItem(olMailItem)
newMail.Subject = pSubject
# newMail.Body = "I AM\nTHE BODY MESSAGE!"
newMail.BodyFormat = 2 # olFormatHTML https://msdn.microsoft.com/en-us/library/office/aa219371(v=office.11).aspx
newMail.HTMLBody = pBody #"<HTML><BODY>Enter the <span style='color:red'>message</span> text here.</BODY></HTML>"
newMail.To = "abc@abc.com; xyz@xyz.com"
if pAttachment != None:
    attachment1 =  pAttachment  # r"C:\Temp\example.pdf"
    newMail.Attachments.Add(Source=attachment1)
newMail.display(False) #This can be True if you want to see the email
newMail.Send()

  ```
