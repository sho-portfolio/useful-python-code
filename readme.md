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



### Save dataframe as html file (formated with css)

```python

# get css stylesheet
with open('css/myCssFile.css', 'r') as myfile:
     style = myfile.read()

# build html
h = '<html><head><style>' + style + ' </style></head><body>'

# global variables report
h += "</br><h4> My Title </h2>" 

h += df.to_html(classes='general') + "</br>" #change this to whatever the style sheet specifies it as

h += '</body></html>'

# create html report file
text_file = open("myHtmlFile.html", "w")
text_file.write(h)
text_file.close()
```

### Make a file read-only but make it writeable when you need to write to it

```python

# https://en.wikipedia.org/wiki/Chmod
# https://stackoverflow.com/questions/28492685/change-file-to-read-only-mode-in-python

import os
from stat import S_IREAD, S_IRGRP, S_IROTH, S_IWOTH, S_IWUSR, S_IWRITE, S_IWGRP

# if file exists make it writeable
if os.path.exists("myFile.txt"):
    os.chmod('myFile.txt', S_IWUSR|S_IWRITE|S_IWGRP|S_IWOTH)

# create file and write, or append to file if it exists
text_file = open("myFile.txt", "a")
text_file.write("hello")
text_file.close()

# make file read-only
os.chmod('myFile.txt', S_IREAD|S_IRGRP|S_IROTH)
```



### How to handle arguments/parameters passed to python

```python
# PART 1
import argparse

if __name__ == '__main__':  # this ensures won't be run when imported otherwise it is by default
    
    parser = argparse.ArgumentParser()

    # create 2 arguments to pass in, one of type int and of type string (but to be used as a list) and set default values
    parser.add_argument("-a", "--paramA", type=str, default = "A,B,CDE")
    parser.add_argument("-b", "--paramB", type=int, default = 99)

    args = parser.parse_args("") # <-- THE EMPTY QUOTES ARE NEEDED FOR JUPYTER NOTEBOOK

    # print out the arguments passed in (convert the string to a list)
    if args.paramA:
        print ('u:', args.paramA.split(','))
        print (type(args.paramA.split(',')))

    if args.paramB:
        print ('m', args.paramB)
        


```python
# PART 2
# simulate passing in different argument values than the defualt
args = parser.parse_args("-a x,y,z -b 5".split())

print(args)
print (args.paramA.split(','))
print (args.paramB)


```python
# PART 3

# simulate passing in different argument values than the defualt for only one of the arguments
args = parser.parse_args("-b 10".split())

print(args)
print (args.paramA.split(','))
print (args.paramB)





