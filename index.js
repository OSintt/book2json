const fs = require("fs");
const XLSX = require("xlsx");

try {
  let compa = XLSX.readFile("companias.xlsx");
  compa = XLSX.utils.sheet_to_json(compa.Sheets[compa.SheetNames[0]]);
  compa = JSON.stringify(compa);

  fs.writeFile("companias.json", compa, "utf8", (err) => {
    if (err) {
      console.log(err);
    } else {
      const workbook = XLSX.readFile("consolidado.xlsx");
      let datosUnificados = [];

      workbook.SheetNames.forEach((sheetName) => {
        const worksheet = workbook.Sheets[sheetName];
        const jsonSheet = XLSX.utils.sheet_to_json(worksheet, {
          blankrows: false,
        });
        datosUnificados = datosUnificados.concat(
          jsonSheet.filter((row) => row["Valor deuda"] !== "Servicio")
        );
      });

      const newWorkbook = XLSX.utils.book_new();
      workbook.SheetNames.forEach((sheetName) => {
        const worksheet = workbook.Sheets[sheetName];
        XLSX.utils.book_append_sheet(newWorkbook, worksheet, sheetName);
      });
      XLSX.writeFile(newWorkbook, "consolidado_modificado.xlsx");

      console.log(
        "Se ha creado una copia de 'consolidado.xlsx' con los cambios en consolidado_modificado.xlsx"
      );
      const datosJSON = JSON.stringify(datosUnificados);

      fs.writeFile("consolidado.json", datosJSON, "utf8", (err) => {
        if (err) {
          console.error("Ocurrió un error al escribir el archivo JSON:", err);
        } else {
          compa = fs.readFileSync("companias.json", "utf-8");
          compa = JSON.parse(compa);

          let conso = fs.readFileSync("consolidado.json", "utf-8");
          conso = JSON.parse(conso);
          console.log(compa.length, "compa", conso.length, "conso");
          if (Array.isArray(conso)) {
            const final = [];
            compa.forEach((c) => {
              const foundConso = conso.find(
                (cs) => c["Documento"] === cs["DNI"]
              );

              if (foundConso) {
                const newData = Object.assign({}, foundConso, c);
                newData.RUC = String(newData.RUC)
                delete newData["DNI"];
                final.push(newData);
              }
            });

            const finalJSON = JSON.stringify(final);

            fs.writeFile("final.json", finalJSON, "utf8", (err) => {
              if (err) {
                console.log(err);
              } else {
                console.log(
                  "El archivo JSON final se ha guardado correctamente."
                );
              }
            });
          } else {
            console.error(
              "El contenido del archivo 'consolidado.json' no es un objeto JSON válido."
            );
          }
        }
      });
    }
  });
} catch (err) {
  console.error("Error al leer el archivo JSON:", err);
}
