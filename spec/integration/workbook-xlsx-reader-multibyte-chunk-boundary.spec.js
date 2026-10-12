const {PassThrough} = require('stream');

const ExcelJS = verquire('exceljs');

/**
 * A UTF-8 sequence of a multi-byte character gets cut
 * by a chunk boundary roughly every other chunk. Decoding each
 * chunk on its own replaced every cut character with U+FFFD (�)
 * so cyrillic symbols were corrupted.
 */
const MIXED_TEXT_FILLER = 'Пример текста with ASCII 123 and кириллицей. ';
const VALUE_LENGTH = 600000;

describe('WorkbookReader', () => {
  describe('when a multi-byte character is split across chunk boundaries', () => {
    it('reads the shared string without replacement characters', async () => {
      const expected = MIXED_TEXT_FILLER.repeat(
        Math.ceil(VALUE_LENGTH / MIXED_TEXT_FILLER.length)
      ).slice(0, VALUE_LENGTH);

      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Кириллица');
      worksheet.getCell('A1').value = expected;
      const buffer = await workbook.xlsx.writeBuffer();

      const input = new PassThrough();
      input.end(Buffer.from(buffer));
      const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader(input, {
        sharedStrings: 'cache',
        styles: 'ignore',
        hyperlinks: 'ignore',
      });

      let actual = null;
      for await (const worksheetReader of workbookReader) {
        for await (const row of worksheetReader) {
          actual = row.getCell(1).value;
        }
      }

      expect(actual).to.equal(expected);
      expect(actual).to.not.contain('\uFFFD');
    });
  });
});
