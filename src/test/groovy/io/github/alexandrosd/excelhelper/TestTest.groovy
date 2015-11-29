package io.github.alexandrosd.excelhelper

class TestTest extends GroovyTestCase{

    String xlsxPath

    TestTest() {
        xlsxPath = this.getClass().getClassLoader().getResource("test.xlsx").getPath()
        System.out.println("XLSX Path: " + xlsxPath)
    }

    void testFull() {
        ExcelDocument xlsx = new ExcelDocument(xlsxPath)
        ExcelSheet sheet = xlsx.getSheet("Sheet1")
        assert(sheet.getKeyValuePairs(1, 0, 2).equals(sheet.getKeyValuePairs(1, 15, 0, 2)))
    }

}
