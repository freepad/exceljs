const ExcelJS = verquire('exceljs');

const TEST_XLSX_FILE_NAME = './spec/out/streaming-formula-report.xlsx';
const SHEET_NAME = 'forecast-report';
const DATA_ROW_COUNT = 100000;

const NUM_FMT_MONEY = '0.00';
const NUM_FMT_PERCENT = '0.00%';

// the report layout: some columns carry source data, others are formulas
// built from the letters of the columns they reference
const columns = [
  {field: 'code', title: 'Item Code'},
  {
    field: 'salesH1',
    title: 'Sales H1',
    numFmt: NUM_FMT_MONEY,
    value: index => index * 10.25,
  },
  {
    field: 'baseH1',
    title: 'Base H1',
    numFmt: NUM_FMT_MONEY,
    value: index => index * 5.75,
  },
  {
    field: 'growthH1',
    title: 'Growth H1',
    numFmt: NUM_FMT_PERCENT,
    value: index => (index % 20) / 100,
  },
  {
    field: 'growthH2',
    title: 'Growth H2',
    numFmt: NUM_FMT_PERCENT,
    value: index => ((index % 17) + 1) / 100,
  },
  {
    field: 'forecastH2',
    title: 'Forecast H2',
    // forward reference: avgGrowth lives to the right of this column
    buildFormula: row =>
      `IFERROR(${cellRef('salesH1', row)}*(1+${cellRef(
        'avgGrowth',
        row
      )}), "")`,
  },
  {
    field: 'totalYear',
    title: 'Total Year',
    // reference to another formula column
    buildFormula: row =>
      `IFERROR(${cellRef('salesH1', row)}+${cellRef('forecastH2', row)}, "")`,
  },
  {
    field: 'growthYoY',
    title: 'Growth YoY',
    numFmt: NUM_FMT_PERCENT,
    buildFormula: row =>
      `IFERROR(${cellRef('baseH1', row)}/${cellRef('salesH1', row)}-1, "")`,
  },
  {
    field: 'rangeSum',
    title: 'Range Sum',
    buildFormula: row =>
      `SUM(${cellRef('salesH1', row)}:${cellRef('baseH1', row)})`,
    buildResult: index => columns[1].value(index) + columns[2].value(index),
  },
  {
    field: 'avgGrowth',
    title: 'Avg Growth',
    buildFormula: row =>
      `AVERAGE(${cellRef('growthH1', row)}:${cellRef('growthH2', row)})`,
  },
];

const columnOffsetByField = new Map(
  columns.map((column, index) => [column.field, index + 1])
);

function columnLetter(field) {
  const column = columnOffsetByField.get(field);
  if (column === undefined) {
    throw new Error(`Export field column is not set: ${field}`);
  }
  let columnNumber = column;
  let letters = '';
  while (columnNumber > 0) {
    const remainder = (columnNumber - 1) % 26;
    letters = String.fromCharCode(65 + remainder) + letters;
    columnNumber = Math.floor((columnNumber - 1) / 26);
  }
  return letters;
}

function cellRef(field, row) {
  return `${columnLetter(field)}${row}`;
}

describe('WorkbookWriter streaming formula report', () => {
  it('writes and reads back a 100k-row report with cross-column formulas', async function() {
    this.timeout(300000);

    const wb = new ExcelJS.stream.xlsx.WorkbookWriter({
      filename: TEST_XLSX_FILE_NAME,
      useSharedStrings: false,
      useStyles: true,
    });
    const ws = wb.addWorksheet(SHEET_NAME);

    const headerRow = ws.getRow(1);
    columns.forEach((column, columnIndex) => {
      headerRow.getCell(columnIndex + 1).value = column.title;
    });
    headerRow.commit();

    for (let index = 1; index <= DATA_ROW_COUNT; index++) {
      const rowNumber = index + 1;
      const row = ws.getRow(rowNumber);
      columns.forEach((column, columnIndex) => {
        const cell = row.getCell(columnIndex + 1);
        if (column.buildFormula) {
          cell.value = {
            formula: column.buildFormula(rowNumber),
            date1904: false,
            ...(column.buildResult && {result: column.buildResult(index)}),
          };
        } else if (column.field === 'code') {
          cell.value = `SKU-${index}`;
        } else {
          cell.value = column.value(index);
        }
        if (column.numFmt) {
          cell.numFmt = column.numFmt;
        }
      });
      row.commit();
    }
    ws.commit();

    return wb.commit().then(() => {
      const wb2 = new ExcelJS.Workbook();
      return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME).then(() => {
        const ws2 = wb2.getWorksheet(SHEET_NAME);
        expect(ws2).to.exist();
        expect(ws2.actualRowCount).to.equal(DATA_ROW_COUNT + 1);

        const header = ws2.getRow(1);
        columns.forEach((column, columnIndex) => {
          expect(header.getCell(columnIndex + 1).value).to.equal(column.title);
        });

        const sampleRowNumbers = [
          2,
          2 + Math.floor(DATA_ROW_COUNT / 2),
          DATA_ROW_COUNT + 1,
        ];
        sampleRowNumbers.forEach(rowNumber => {
          const index = rowNumber - 1;
          const row = ws2.getRow(rowNumber);

          expect(row.getCell(1).value).to.equal(`SKU-${index}`);
          expect(row.getCell(2).value).to.equal(columns[1].value(index));
          expect(row.getCell(3).value).to.equal(columns[2].value(index));
          expect(row.getCell(4).value).to.equal(columns[3].value(index));
          expect(row.getCell(5).value).to.equal(columns[4].value(index));

          expect(row.getCell(6).value.formula).to.equal(
            `IFERROR(B${rowNumber}*(1+J${rowNumber}), "")`
          );
          expect(row.getCell(7).value.formula).to.equal(
            `IFERROR(B${rowNumber}+F${rowNumber}, "")`
          );
          expect(row.getCell(8).value.formula).to.equal(
            `IFERROR(C${rowNumber}/B${rowNumber}-1, "")`
          );
          expect(row.getCell(9).value.formula).to.equal(
            `SUM(B${rowNumber}:C${rowNumber})`
          );
          expect(row.getCell(10).value.formula).to.equal(
            `AVERAGE(D${rowNumber}:E${rowNumber})`
          );

          // only the rangeSum column was written with a cached result
          expect(row.getCell(9).value.result).to.equal(
            columns[8].buildResult(index)
          );
          [6, 7, 8, 10].forEach(columnIndex => {
            expect(row.getCell(columnIndex).value.result).to.be.undefined();
          });

          expect(row.getCell(2).numFmt).to.equal(NUM_FMT_MONEY);
          expect(row.getCell(3).numFmt).to.equal(NUM_FMT_MONEY);
          expect(row.getCell(4).numFmt).to.equal(NUM_FMT_PERCENT);
          expect(row.getCell(5).numFmt).to.equal(NUM_FMT_PERCENT);
          expect(row.getCell(8).numFmt).to.equal(NUM_FMT_PERCENT);
        });
      });
    });
  });
});
