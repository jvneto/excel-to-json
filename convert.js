const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

const inputFolderPath = 'input';
const outputFolderPath = 'output';

// Lista de arquivos na pasta de entrada
const files = fs.readdirSync(inputFolderPath);

files.forEach(file => {
  // Verifica se o arquivo é um arquivo Excel
  if (path.extname(file) === '.xlsx') {
    const filePath = path.join(inputFolderPath, file);

    // Lê o arquivo Excel
    const workbook = xlsx.readFile(filePath);

    // Obtém os dados da primeira planilha (índice 0)
    const sheetName = workbook.SheetNames[0];
    const jsonData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);

    // Aplica trim nas chaves do JSON
    const trimmedData = jsonData.map(obj => {
      const trimmedObj = {};
      for (const key in obj) {
        if (obj.hasOwnProperty(key)) {
          trimmedObj[key.trim()] = obj[key];
        }
      }
      return trimmedObj;
    });

    // Cria o caminho do arquivo de saída
    const outputFilePath = path.join(outputFolderPath, `${path.parse(file).name}.json`);

    // Escreve os dados no arquivo JSON de saída
    fs.writeFileSync(outputFilePath, JSON.stringify(trimmedData, null, 2));

    console.log(`Arquivo JSON exportado com sucesso: ${outputFilePath}`);
  }
});

console.log('Concluído.');
