const XLSX = require('xlsx');

const readExcel = async (file) => {
  return new Promise((resolve, reject) => {
    const fileData = XLSX.read(file, {
      type: 'buffer',
    });

    if (fileData.isBuffer) {
      const workBook = XLSX.utils.book_new_from_buffer(fileData);
      const wsNames = workBook.SheetNames;

      if (wsNames.length > 2) {
        const firstSheet = workBook.Sheets[wsNames[0]];
        const secondSheet = workBook.Sheets[wsNames[1]];
        const thirdSheet = workBook.Sheets[wsNames[2]];

        resolve({
          firstSheet,
          secondSheet,
          thirdSheet,
        });
      } else {