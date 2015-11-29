package io.github.alexandrosd.excelhelper

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet

/**
 * ExcelSheet
 * Represents an Excel Sheet
 */
class ExcelSheet {

    protected Sheet sheet

    /**
     * Initialize the Excel Sheet, providing a Sheet object
     * @param sheet
     */
    public ExcelSheet(Sheet sheet) {
        this.sheet = sheet
    }

    /**
     * Get a Row by index
     * @param i
     * @return
     */
    private Row getRow(int i) {
        return sheet.iterator().getAt(i)
    }

    /**
     * Get a Sheet's name
     * @return
     */
    public String getName() {
        return sheet.getSheetName()
    }

    /**
     * Get value of cell at x,y
     * @param columnId
     * @param rowId
     * @return
     */
    public String getCellValue(int columnId, int rowId) {
        return getCell(columnId, rowId).getStringCellValue()
    }

    /**
     * Get a Cell by index
     * @param columnId
     * @param rowId
     * @return
     */
    private Cell getCell(int columnId, int rowId) {
        Row row = sheet.iterator().getAt(rowId)
        Cell cell = row.getAt(columnId)
        return cell
    }

    /**
     * Get Column values
     * @param columnId
     * @param firstRow
     * @param lastRow
     * @return
     */
    public List<String> getColumnCellValues(int columnId, int firstRow, int lastRow) {
        List<String> values = new ArrayList<String>()
        for (int i = firstRow; i <= lastRow; i++) {
            values.add(getCellValue(columnId, i))
        }
        return values
    }

    /**
     * Get Key/Value pairs by providing the indexes of the key column, value column, and first and last rows
     * @param firstRow
     * @param lastRow
     * @param keyColumnId
     * @param valueColumnId
     * @return
     */
    public List<Map<String, String>> getKeyValuePairs(int firstRow, int lastRow, int keyColumnId, int valueColumnId) {
        List< Map<String, String>> pairs = new ArrayList<Map<String, String>>()

        for (int i = firstRow; i <= lastRow; i++) {
            Map<String, String> pair = new HashMap<String, String>();
            String key = getCellValue(keyColumnId, i)
            String value = getCellValue(valueColumnId, i)
            pair.put(key, value)
            pairs.add(pair)
        }
        return pairs
    }

    /**
     * Get Key/Value pairs by providing the indexes of the key column, value column, and first row. Iteration continues until the last available row
     * @param firstRow
     * @param keyColumnId
     * @param valueColumnId
     * @return
     */
    public List<Map<String, String>> getKeyValuePairs(int firstRow, int keyColumnId, int valueColumnId) {
        return getKeyValuePairs(firstRow, sheet.getLastRowNum(), keyColumnId, valueColumnId)
    }
}
