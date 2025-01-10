function convertCSVtoExcel() {
  const csvFileInput = document.getElementById('csvFileInput');

  if (csvFileInput.files.length > 0) {
    const csvFile = csvFileInput.files[0];

    Papa.parse(csvFile, {
      complete: function (result) {
        const worksheet = XLSX.utils.json_to_sheet(result.data);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");

        XLSX.writeFile(workbook, 'Convertido.xlsx');
      },

      header: true
    });
  }
  else {
    alert("Por favor, selecione um arquivo .csv.");
  }

}