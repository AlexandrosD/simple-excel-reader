package io.github.alexandrosd.excelhelper

import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * ExcelDocument
 * Represents an excel file
 */
class ExcelDocument {

    protected FileInputStream file
    protected Workbook workbook

    /**
     * Initialize an ExcelDocument
     * @param filename
     */
    public ExcelDocument(String filename) {
        file = new FileInputStream(new File(filename));

        if (filename.endsWith(".xlsx")) {
            // XSSFWorkbook
            workbook = new XSSFWorkbook(file);
        }
        if (filename.endsWith(".xls")) {
            // HSSFWorkbook
            workbook = new HSSFWorkbook(file);
        }
    }

    /**
     * Get the sheets a document contains
     * @return
     */
    public Collection<ExcelSheet> getSheets() {
        int numberOfSheets = workbook.getNumberOfSheets()
        Collection<ExcelSheet> sheets = new ArrayList<ExcelSheet>();
        for (int i = 0; i < numberOfSheets; i++) {
            sheets.add(new ExcelSheet(workbook.getSheetAt(i)))
        }
        return sheets
    }

    /**
     * Get a sheet by name
     * @param sheetName
     * @return
     */
    public ExcelSheet getSheet(String sheetName) {
        return new ExcelSheet(workbook.getSheet(sheetName))
    }

}
