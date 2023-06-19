const fs = require("fs");
const path = require("path");
const xlsx = require("xlsx");
const csv = require("csv-parser");
const glob = require("glob");
const readlineSync = require("readline-sync");

const local = __dirname;
console.log(local);

function converterPlanilhas() {
  const planilhasDir = path.join(local, 'planilhas');
  if (fs.existsSync(planilhasDir) && fs.readdirSync(planilhasDir).length >= 2) {
    const files = fs.readdirSync(planilhasDir);
    let num = 1;
    let listDfs = null;
    let validFiles = [];

    files.forEach((file) => {
      const filePath = path.join(planilhasDir, file);
      let df = null;
      if (path.extname(filePath) === '.xlsx') {
        const workbook = xlsx.readFile(filePath);
        if (workbook.SheetNames.length > 0 && workbook.Sheets[workbook.SheetNames[0]]) {
          validFiles.push(filePath);
        }
      } else if (path.extname(filePath) === '.csv') {
        validFiles.push(filePath);
      }
    });

    validFiles.forEach((file) => {
      let df = null;
      if (path.extname(file) === '.xlsx') {
        const workbook = xlsx.readFile(file);
        const sheetNames = workbook.SheetNames.slice(0, 0xFFFF);
        sheetNames.forEach((sheetName) => {
          const worksheet = workbook.Sheets[sheetName];
          const jsonData = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
          if (jsonData.length > 0) {
            const newDf = xlsx.utils.json_to_sheet(jsonData);
            if (!listDfs) {
              listDfs = newDf;
            } else {
              xlsx.utils.book_append_sheet(listDfs, newDf, sheetName);
            }
          }
        });
      } else if (path.extname(file) === '.csv') {
        const jsonData = [];
        fs.createReadStream(file)
          .pipe(csv())
          .on('data', (data) => jsonData.push(data))
          .on('end', () => {
            if (jsonData.length > 0) {
              const newDf = xlsx.utils.json_to_sheet(jsonData);
              if (!listDfs) {
                listDfs = newDf;
              } else {
                xlsx.utils.book_append_sheet(listDfs, newDf, path.basename(file, '.csv'));
              }
            }
          });
      }

      console.log(file);
      if (df) {
        if (num === 1) {
          listDfs = df;
        } else {
          xlsx.utils.book_append_sheet(listDfs, df);
        }
        num++;
      }
    });

    if (num > 1) {
      const nome = readlineSync.question('Um nome para a nova planilha: ');
      const outputFile = path.join(local, `${nome}.xlsx`);
      xlsx.writeFile(listDfs, outputFile, { bookType: 'xlsx', type: 'buffer' });
      console.log('Planilha criada com sucesso!');
    } else {
      console.log('Não foi possível criar a planilha. Verifique se os arquivos de entrada têm conteúdo válido.');
    }
  } else {
    console.log('Coloque mais de um arquivo .xlsx ou .csv na pasta "planilhas".');
  }
}


if (fs.existsSync(path.join(local, "planilhas"))) {
  converterPlanilhas();
} else {
  console.log('Criando pasta "planilhas".');
  fs.mkdirSync("planilhas");
  converterPlanilhas();
}
