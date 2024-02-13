//Importing the library
const xlsx = require('xlsx')

const workbook = xlsx.readFile('./Cópia de Engenharia de Software - Desafio Gustavo Carvalho.xlsx')

//Accessing the spreadsheets and cells, with the file loaded, you can access the spreadsheets and the information contained in each cell
const worksheet = workbook.Sheets[workbook.SheetNames[0]];

//Loop to read and write spreadsheet cells in ascending order
for (let i = 4; i <28; i++) {
    const array = [worksheet[`D${i}`].v,worksheet[`E${i}`].v,worksheet[`F${i}`].v]
    const initialValue = 0;
    const sumWithInitial = array.reduce((accumulator, currentValue) => accumulator + currentValue/3,initialValue)
    const result = Math.floor(sumWithInitial)
    
    let cellAddress = `G${i}`
    
    //Conditional structure to be able to edit empty cells
    if (!worksheet[cellAddress]) {
        worksheet[cellAddress] = { t: 's', v: '' }; 
    }

    const fouls = worksheet[`C${i}`].v
    const classes = 60*0.25

    if(result<50 ){
        worksheet[cellAddress].v = 'Reprovado por Nota';
        if(fouls>=classes){
            worksheet[cellAddress].v = 'Reprovado por Falta';
        }
    }else if(result<70){
        worksheet[cellAddress].v = 'Exame Final';
        if(fouls>=classes){
            worksheet[cellAddress].v = 'Reprovado por Falta';
        }
    }else if(result>=70){
        worksheet[cellAddress].v = 'Aprovado';
        if(fouls>=classes){
            worksheet[cellAddress].v = 'Reprovado por Falta';
        }
    }

}

//Loop to read and write another spreadsheet column
for(let i = 4;i<28;i++){
    const situation = worksheet[`G${i}`].v

    const array = [worksheet[`D${i}`].v,worksheet[`E${i}`].v,worksheet[`F${i}`].v]
    const initialValue = 0;
    const sumWithInitial = array.reduce((accumulator, currentValue) => accumulator + currentValue/3,initialValue)
    const result = Math.floor(sumWithInitial)

    let cellAddress = `H${i}`
    if (!worksheet[cellAddress]) {
        worksheet[cellAddress] = { t: 's', v: '' }; 
    }

    if(situation !=='Exame Final'){
        worksheet[cellAddress].v = 0
    }else{
        const calcule = (100 - result)
        worksheet[cellAddress].v = calcule
    }
}

//Library function to edit spreadsheet in the project folder.
xlsx.writeFile(workbook, '[RESULT]Engenharia de Software - Desafio Gustavo Carvalho.xlsx')
//xlsx.writeFile(workbook, '[TEST]Engenharia de Software - Desafio Gustavo Carvalho.xlsx');