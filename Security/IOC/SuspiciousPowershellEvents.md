Run the script as administrator on the affected machine.

- By default the script will export to "c:\", but it can be changed with the $csvDir variable

Collect the csv output files from the root of the C drive and transfer them to your machine.


Using Notepad++ will make analysis much easier then notepad or Excel

# encodedEvents.csv
Open the file in Notepad++

Search for " -e " or "-encoded" using "Find all in current document"

![](https://i.imgur.com/YJUinEf.png)

It is extremely unusual for this command to be run on a computer. You can find what the command did by pasting the encoded text into a site like https://www.base64decode.org/

# downloadstringEvents.csv
Open the file in Notepad++

Search for "downloadstring" using "Find all in current document"

![](https://i.imgur.com/WjbCknV.png)

Look through the results for suspicious entries. This command is not usually run by users/applications.

# iexEvents.csv
Open the file in Notepad++

Search for "ies" or "invoke-expression" using "Find all in current document"

![](https://i.imgur.com/xuGViTo.png)

Look through the results for suspicious entries. This command is not usually run by users/applications.