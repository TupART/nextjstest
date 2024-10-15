import { useState } from 'react';
import XLSX from 'xlsx';
import { saveAs } from 'file-saver';

export default function ExcelHandler() {
  const [data, setData] = useState(null);  // Datos del archivo .xlsx
  const [template, setTemplate] = useState(null);  // PlantillaSTEP4.xlsx
  const [selectedRows, setSelectedRows] = useState([]);

  // Cargar archivo de datos .xlsx
  const handleFileUpload = (file) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const workbook = XLSX.read(e.target.result, { type: 'binary' });
      const sheetName = workbook.SheetNames[0];
      const sheet = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });
      setData(sheet);  // Guarda el contenido del archivo para su manipulación
    };
    reader.readAsBinaryString(file);
  };

  // Cargar la plantilla PlantillaSTEP4.xlsx
  const handleTemplateUpload = (file) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const workbook = XLSX.read(e.target.result, { type: 'binary' });
      setTemplate(workbook);  // Guardamos la plantilla cargada
    };
    reader.readAsBinaryString(file);
  };

  // Rellenar la plantilla
  const handleFillTemplate = () => {
    if (!template || !data) {
      alert('Por favor, sube ambos archivos antes de continuar.');
      return;
    }

    const sheet = template.Sheets['Sheet1'];
    const updatedSheet = fillTemplate(sheet, data, selectedRows);  // Llenar la plantilla

    const newWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWorkbook, updatedSheet, 'Sheet1');
    const wbout = XLSX.write(newWorkbook, { bookType: 'xlsx', type: 'binary' });

    saveAs(new Blob([s2ab(wbout)], { type: 'application/octet-stream' }), 'PlantillaSTEP4_Modificada.xlsx');
  };

  // Lógica de rellenado
  const fillTemplate = (templateSheet, data, selectedRows) => {
    selectedRows.forEach((row, index) => {
      // Rellenar columnas "Name" y "Surname"
      templateSheet[`C${index + 7}`] = { v: data[row][0] };  // Columna "Name"
      templateSheet[`D${index + 7}`] = { v: data[row][1] };  // Columna "Surname"
      templateSheet[`E${index + 7}`] = { v: data[row][14] };  // Columna "Primary email"

      const vaSerPCC = data[row][4];  // Columna E del archivo
      const market = data[row][2];  // Columna C del archivo

      // Lógica para "Primary phone" (Columna F)
      if (vaSerPCC === 'Y') {
        if (market === 'DACH') templateSheet[`F${index + 7}`] = { v: "/+4940210918145 /+43122709858 /+41445295828" };
        else if (market === 'France') templateSheet[`F${index + 7}`] = { v: "/+33180037979" };
        else if (market === 'Spain') templateSheet[`F${index + 7}`] = { v: "/+34932952130" };
        else if (market === 'Italy') templateSheet[`F${index + 7}`] = { v: "/+390109997099" };
      }

      // Lógica para "Workgroup" (Columna G)
      if (vaSerPCC === 'Y') {
        if (market === 'DACH') templateSheet[`G${index + 7}`] = { v: "D_PCC" };
        else if (market === 'France') templateSheet[`G${index + 7}`] = { v: "F_PCC" };
        else if (market === 'Spain') templateSheet[`G${index + 7}`] = { v: "E_PCC" };
        else if (market === 'Italy') templateSheet[`G${index + 7}`] = { v: "I_PCC" };
      } else if (vaSerPCC === 'N') {
        if (market === 'DACH') templateSheet[`G${index + 7}`] = { v: "D_Outbound" };
        else if (market === 'France') templateSheet[`G${index + 7}`] = { v: "F_Outbound" };
        else if (market === 'Spain') templateSheet[`G${index + 7}`] = { v: "E_Outbound" };
      }

      // Lógica para "Team" (Columna H)
      if (vaSerPCC === 'Y') {
        if (market === 'DACH') templateSheet[`H${index + 7}`] = { v: "Team_D_CCH_PCC_1" };
        else if (market === 'France') templateSheet[`H${index + 7}`] = { v: "Team_F_CCH_PCC_1" };
        else if (market === 'Spain') templateSheet[`H${index + 7}`] = { v: "Team_E_CCH_PCC_1" };
        else if (market === 'Italy') templateSheet[`H${index + 7}`] = { v: "Team_I_CCH_PCC_1" };
      } else if (vaSerPCC === 'N') {
        if (market === 'DACH') templateSheet[`H${index + 7}`] = { v: "Team_D_CCH_B2C_1" };
        else if (market === 'France') templateSheet[`H${index + 7}`] = { v: "Team_F_CCH_B2C_1" };
        else if (market === 'Spain') templateSheet[`H${index + 7}`] = { v: "Team_E_CCH_B2C_1" };
      }

      // Columna L "Is PCC"
      templateSheet[`L${index + 7}`] = { v: vaSerPCC === 'Y' ? 'Y' : 'N' };

      // Columna Q "CTI User" y columna R "TTG UserID 1"
      templateSheet[`Q${index + 7}`] = { v: data[row][15] };  // "B2E User Name"
      templateSheet[`R${index + 7}`] = { v: data[row][15] };

      // Columna V "Campaign Level"
      templateSheet[`V${index + 7}`] = { v: vaSerPCC === 'TL' ? 'Team Leader' : 'Agent' };
    });
    return templateSheet;
  };

  // Convertir string a ArrayBuffer para la descarga
  const s2ab = (s) => {
    const buf = new ArrayBuffer(s.length);
    const view = new Uint8Array(buf);
    for (let i = 0; i !== s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
    return buf;
  };

  return (
    <div>
      <h1>Sube los archivos Excel</h1>
      <h2>1. Sube el archivo de datos .xlsx</h2>
      <input type="file" accept=".xlsx" onChange={(e) => handleFileUpload(e.target.files[0])} />
      
      <h2>2. Sube la plantilla PlantillaSTEP4.xlsx</h2>
      <input type="file" accept=".xlsx" onChange={(e) => handleTemplateUpload(e.target.files[0])} />

      {data && template ? (
        <div>
          <h3>Selecciona los nombres que deseas incluir:</h3>
          <form>
            {data.map((row, index) => (
              <div key={index}>
                <label>
                  <input type="checkbox" value={index} onChange={(e) => {
                    const checked = e.target.checked;
                    setSelectedRows(prev => checked ? [...prev, index] : prev.filter(r => r !== index));
                  }} />
                  {row[0]} {row[1]}
                </label>
              </div>
            ))}
          </form>
          <button onClick={handleFillTemplate}>Generar archivo</button>
        </div>
      ) : (
        <p>Por favor, sube ambos archivos antes de continuar.</p>
      )}
    </div>
  );
}
