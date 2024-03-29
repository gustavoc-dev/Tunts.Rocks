<img src="https://tunts.rocks/_next/static/media/logoTuntsRocksHeader.c8146752.png">
<br>

<h1>Welcome</h1>

Hi! Welcome to my selection process for Tunts.rock

I'm in the technical challenge part, where I have to create an application in a programming language of my choice, I chose to use Javascript. The application must be able to read a Google Sheets spreadsheet, search for the necessary information, calculate and write the result in the spreadsheet.

I used the xlsx.js library for this application.
<br>
<a href='https://www.npmjs.com/package/xlsx'>xlsx.js</a>

I also created a copy of the spreadsheet used in the challenge, according to the instructions.
<br>
<a href='https://docs.google.com/spreadsheets/d/13XjO37lhhZVJfvaRuYOc60YJWA8hnbiiJoI0MPaCeic/edit?usp=drive_link'>Google Drive link.</a>

<h1>Instructions</h1>

Firstly I started npm in the application
``` javascript
  npm init --y
```
Then I installed the xlsx.json library
``` javascript
  npm install xlsx
```

I used a VsCode extension to view spreadsheets in the IDE called Excel Viewer
<br>
<a href='https://marketplace.visualstudio.com/items?itemName=GrapeCity.gc-excelviewer'>Excel Viewer</a>

<h1>To Test</h1>
The code is commented to make it easier to understand, the spreadsheet in the file folder will be edited, but there will be a folder with the unedited spreadsheet, I used node.js to run the index.js file, creating an output file with the name [TEST]
This line of code will be commented

``` javascript
//xlsx.writeFile(workbook, '[TEST]Engenharia de Software - Desafio Gustavo Carvalho.xlsx');
```
To test the code with node, the output file will be named [RESULT] 

``` javascript
  xlsx.writeFile(workbook, '[RESULT]Engenharia de Software - Desafio Gustavo Carvalho.xlsx')
```
And created a new spreadsheet with the name [RESULT] after running the node in the index.js file
``` javascript
  node index.js
```

<h1>Thanks!</h1>

I would like to thank you for the opportunity and I hope to have achieved the expected quality 🤘
