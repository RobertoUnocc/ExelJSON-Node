var XLSX = require("xlsx");


const ExcelAJSON = () => {
  const excel = XLSX.readFile(
    "C:\\Users\\jespiritu.CHSALT\\Desktop\\CARLOSS\\Express\\ExelJson\\VRTPvacio.xlsx"
  );
  var nombreHoja = excel.SheetNames; // regresa un array
  let datos = XLSX.utils.sheet_to_json(excel.Sheets[nombreHoja[0]]);
  //console.log(datos);

  const jDatos = [];
  
  for (let i = 0; i < datos.length; i++) {
    const dato = datos[i];
    var calendario = new Date((dato.FECHA - (25567 + 1)) * 86400 * 1000);

    jDatos.push({
      //...dato,
      Origen:dato.ORIGEN,
      Transito:dato.TRANSITO,
      Destino:dato.DESTINO,
      Fecha: calendario.toLocaleDateString(),
      ANIO: calendario.getFullYear(),
      Fuente: dato.FUENTE
      
    });
  }
  console.log(jDatos);
};
ExcelAJSON();


/*     "C:\\Users\\jespiritu.CHSALT\\Desktop\\Express\\ExelJson\\prueba.xlsx" */