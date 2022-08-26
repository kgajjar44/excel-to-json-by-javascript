let selectedFiles;
document.getElementById("input").addEventListener("change", (event) => {
  selectedFiles = event.target.files;
});

document.getElementById("button").addEventListener("click", () => {
  const response = { data: { finalData: [] }, headerData: { finalData: [] } };
  readExcelFile(0, response);
});

function readExcelFile(fileIndex, response) {
  const fileReader = new FileReader();
  fileReader.readAsBinaryString(selectedFiles[fileIndex]);
  fileReader.onload = (event) => {
    const excelData = event.target.result;
    const workbook = XLSX.read(excelData, { 
        type: "binary", cellDates: true, dateNF: 'dd-mm-yyyy hh:MM'
    });
    response.data[fileIndex + 1] = [];
    response.headerData[fileIndex + 1] = [];
    workbook.SheetNames.forEach((sheet, sheetIndex) => {
      const rowObject = XLSX.utils.sheet_to_json(workbook.Sheets[sheet], {
        raw: false,
      });
      if (rowObject && rowObject.length > 0) {
        response.headerData[fileIndex + 1] = Object.keys(rowObject[0]);
        response.headerData.finalData = [
          ...new Set([
            ...response.headerData.finalData,
            ...Object.keys(rowObject[0]),
          ]),
        ];
      }
      response.data[fileIndex + 1] =
        response.data[fileIndex + 1].concat(rowObject);
      response.data.finalData = response.data.finalData.concat(rowObject);
      if (workbook.SheetNames.length - 1 == sheetIndex) {
        if (selectedFiles.length - 1 == fileIndex) {
          exportAsExcelFile(response.data, response.headerData, "Merged.xlsx");
          console.log(response);
        } else {
          readExcelFile(++fileIndex, response);
        }
      }
    });
  };
}

function exportAsExcelFile(json, headerName, excelFileName) {
  const objectMaxLength = [];
  const workSheet = {};
  for (const key in json) {
    Object.keys(headerName[key]).map((headerKey) => {
      const value =
        headerName[key][headerKey] === undefined || headerName[key][headerKey] === null
          ? ""
          : headerName[key][headerKey];
      objectMaxLength[headerKey] = value.length + 3;
    });
    workSheet[key] = XLSX.utils.json_to_sheet(json[key], {
      header: headerName[key],
    });
    json[key].map((arr) => {
      Object.keys(headerName[key]).map((headerKey) => {
        const value =
          arr[headerName[key][headerKey]] === undefined || arr[headerName[key][headerKey]] === null
            ? ""
            : arr[headerName[key][headerKey]];
        objectMaxLength[headerKey] = Math.max(
          objectMaxLength[headerKey] || 0,
          value.toString().length
        );
      });
    });
    workSheet[key]["!cols"] = objectMaxLength.map((width) => {
      return {
        width,
      };
    });
  }
  const workbook = { Sheets: workSheet, SheetNames: Object.keys(json) };

  const excelBuffer =
    "data:@file/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64," +
    XLSX.write(workbook, { bookType: "xlsx", type: "base64", cellDates: true });

  const downloadLink = document.createElement("a");
  downloadLink.href = excelBuffer;
  downloadLink.download = excelFileName;
  downloadLink.click();
}
