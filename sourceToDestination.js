const XLSX = require('xlsx');
const FS = require('fs');

const fileArrayToUse = {fileName: '', outputName: ''};

const manageFile = function (filename, outputName) {

    if (fileArrayToUse.fileName.length === 0 || outputName.length === 0) {
        manageArgsNames(fileArrayToUse);
    }
    const workBook = XLSX.readFile(filename);
    const sheetNames = workBook.SheetNames;

    sheetNames.forEach(sheetName => {
        const workSheet = workBook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(workSheet, {header: 1});

        const resultOutPut = data.map(row => {
            if (row[0] === 'headerName') {
                return;
            }
            return (`desired string ${row[0].trim()} with desired column managed`)
        });

        const finalResult = resultOutPut.join('\n');
        const outputFile = outputName;

        FS.appendFileSync(outputFile, finalResult, (error) => {
            if (error) {
                console.error('error while writing output file : {}', error);
            } else {
                console.info('file {} was correctly created', outputFile);
            }
        })
    });

}

const manageArgsNames = function (objectToUse) {
// Récupérer les arguments de la ligne de commande
    const args = process.argv;

// Chercher l'index de l'argument '--prop'
    let propertyArgIndex = args.findIndex(arg => arg.startsWith('--source='));

    if (propertyArgIndex !== -1) {
        // Extraire la valeur de l'argument '--prop'
        objectToUse.fileName = args[propertyArgIndex].split('=')[1];

        console.log('La valeur de --source est :', objectToUse.fileName);

    } else {
        console.error('L\'argument --property n\'a pas été fourni ou est mal formaté.');
    }

    propertyArgIndex = args.findIndex(arg => arg.startsWith('--destination='));

    if (propertyArgIndex !== -1) {
        objectToUse.outputName = args[propertyArgIndex].split('=')[1];

        console.log('La valeur de --destination est :', objectToUse.outputName);
    } else {
        console.error('L\'argument --property n\'a pas été fourni ou est mal formaté.');
    }
}

const array = [fileArrayToUse];
array.forEach(content => {
    manageFile(content.fileName, content.outputName);
});