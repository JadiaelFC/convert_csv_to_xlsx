const fs = require('fs');
const { parse } = require('csv-parse');
const ExcelJS = require('exceljs');

const delimiter = ';';

const convertCsvToXlsx = async (csvFilePath, xlsxFilePath) => {
    const parser = parse({ delimiter: delimiter, columns: true, relax_column_count: true });
    const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: xlsxFilePath });
    const worksheet = workbook.addWorksheet('Sheet1');

    let headerWritten = false;
    let rowCount = 0;

    fs.createReadStream(csvFilePath)
        .pipe(parser)
        .on('data', (record) => {
            if (!headerWritten) {
                worksheet.columns = Object.keys(record).map(key => ({ header: key, key }));
                headerWritten = true;
            }
            worksheet.addRow(record).commit();
            rowCount++;
        })
        .on('end', async () => {
            worksheet.commit();
            await workbook.commit();
            console.log(`File converted: ${xlsxFilePath}`);
        })
        .on('error', (err) => {
            console.error(err);
        });
};

// Caminhos dos arquivos CSV de entrada e XLSX de saída
const csvFilePath = './input.csv';
const xlsxFilePath = './output.xlsx';

// Executa a conversão
convertCsvToXlsx(csvFilePath, xlsxFilePath);
