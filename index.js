const fs = require("fs");
const path = require("path");
const xlsx = require("xlsx");
const csv = require("csv-parser");
const readlineSync = require("readline-sync");

const local = __dirname;

function converterPlanilhas() {
  const planilhasDir = path.join(local, "planilhas");
  if (fs.existsSync(planilhasDir) && fs.readdirSync(planilhasDir).length >= 2) {
    const files = fs.readdirSync(planilhasDir);
    let mergedData = [];

    files.forEach((file) => {
      const filePath = path.join(planilhasDir, file);
      if (path.extname(filePath) === ".xlsx") {
        const workbook = xlsx.readFile(filePath);
        workbook.SheetNames.forEach((sheetName) => {
          const worksheet = workbook.Sheets[sheetName];
          const jsonData = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
          mergedData = mergedData.concat(jsonData);
        });
      } else if (path.extname(filePath) === ".csv") {
        const jsonData = [];
        fs.createReadStream(filePath)
          .pipe(csv())
          .on("data", (data) => jsonData.push(data))
          .on("end", () => {
            mergedData = mergedData.concat(jsonData);
          });
      }
    });

    if (mergedData.length > 0) {
      const nome = readlineSync.question("Um nome para a nova planilha: ");
      const outputFile = path.join(local, `${nome}.xlsx`);
      const workbook = xlsx.utils.book_new();
      const worksheet = xlsx.utils.json_to_sheet(mergedData);
      xlsx.utils.book_append_sheet(workbook, worksheet, "Sheet1");
      xlsx.writeFile(workbook, outputFile);
      console.log("Planilha criada com sucesso!");
    } else {
      console.log(
        "Não foi possível criar a planilha. Verifique se os arquivos de entrada têm conteúdo válido."
      );
    }
  } else {
    console.log(
      'Coloque mais de um arquivo .xlsx ou .csv na pasta "planilhas".'
    );
  }
}

if (fs.existsSync(path.join(local, "planilhas"))) {
  converterPlanilhas();
} else {
  console.log('Criando pasta "planilhas".');
  fs.mkdirSync("planilhas");
  converterPlanilhas();
}
