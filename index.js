const fs = require("fs");
const XLSX = require("xlsx");

const processFile = async () => {
  try {
    let compa = XLSX.readFile("companias.xlsx");
    compa = XLSX.utils.sheet_to_json(compa.Sheets[compa.SheetNames[0]]);
    compa = JSON.stringify(compa);
    await fs.writeFile("companias.json", compa, "utf-8");
    const workbook = XLSX.readFile("consolidado.xlsx");
    let datosUnificados = [];

    for (const sheetName of workbook.SheetNames) {
      const worksheet = workbook.Sheets[sheetName];
      const jsonSheet = XLSX.utils.sheet_to_json(worksheet, {
        blankrows: false,
      });
      datosUnificados = datosUnificados.concat(
        jsonSheet.filter((row) => row["Valor deuda"] !== "Servicio")
      );
    }
    const newWorkbook = XLSX.utils.book_new();
    workbook.SheetNames.forEach((sheetName) => {
      const worksheet = workbook.Sheets[sheetName];
      XLSX.utils.book_append_sheet(newWorkbook, worksheet, sheetName);
    });
    await XLSX.writeFile(newWorkbook, "consolidado_modificado.xlsx");
    console.log(
      "Se ha creado una copia de 'consolidado.xlsx' con los cambios en consolidado_modificado.xlsx"
    );
    const datosJSON = JSON.stringify(datosUnificados);
    await fs.writeFile("consolidado.json", datosJSON, "utf-8");
    compa = await fs.readFile("companias.json", "utf-8");
    compa = JSON.parse(compa);
    let conso = await fs.readFile("consolidado.json", "utf-8");
    conso = JSON.parse(conso);
    console.log(compa.length, "compa", conso.length, "conso");
    const final = [];
    compa.forEach((c) => {
      const foundConso = conso.find((cs) => c["DNI"] === cs["Documento"]);
      if (foundConso) {
        const newData = Object.assign({}, foundConso, c);
        newData.RUC = String(newData.RUC);
        delete newData["Documento"];
        final.push(newData);
      }
    });
    const finalJSON = JSON.stringify(final);
    await fs.writeFile("final.json", finalJSON, "utf8");
    console.log("El archivo JSON final se ha guardado correctamente.");
  } catch (e) {
    console.log('Ha ocurrido un error intesperado:', e);
  }
};

processFile();