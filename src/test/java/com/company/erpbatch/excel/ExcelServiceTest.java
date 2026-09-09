package com.company.erpbatch.excel;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import javax.sql.rowset.CachedRowSet;
import javax.sql.rowset.RowSetMetaDataImpl;
import javax.sql.rowset.RowSetProvider;
import java.nio.file.Path;
import java.sql.Types;
import java.time.LocalDateTime;

import static org.junit.jupiter.api.Assertions.assertEquals;

class ExcelServiceTest {
    @TempDir
    Path tempDir;

    @Test
    void replacesTimestampAndAddsExtension() {
        ExcelService service = new ExcelService();
        Path path = service.resolveOutputPath(Path.of("C:/output"), "report_{timestamp}", LocalDateTime.of(2026, 9, 8, 1, 2, 3));
        assertEquals(Path.of("C:/output/report_20260908_010203.xlsx"), path);
    }

    @Test
    void writesErpLayoutAndTypedValues() throws Exception {
        ExcelService service = new ExcelService();
        ExcelService.ExcelWriteResult result = service.write(totalRow(), tempDir, "sales.xlsx");

        try (XSSFWorkbook workbook = new XSSFWorkbook(result.outputFile().toFile())) {
            var sheet = workbook.getSheet("Report");
            assertEquals("일자별판매실적분석", sheet.getRow(0).getCell(0).getStringCellValue());
            assertEquals(1, sheet.getNumMergedRegions());
            assertEquals("A1:AG1", sheet.getMergedRegion(0).formatAsString());
            assertEquals("FF03B051", sheet.getRow(1).getCell(0).getCellStyle().getFillForegroundColorColor().getARGBHex());
            assertEquals("FFACFAB8", sheet.getRow(2).getCell(0).getCellStyle().getFillForegroundColorColor().getARGBHex());
            assertEquals(CellType.NUMERIC, sheet.getRow(2).getCell(12).getCellType());
            assertEquals("#,##0", sheet.getRow(2).getCell(12).getCellStyle().getDataFormatString());
            for (int index = 17; index <= 21; index++) {
                assertEquals(CellType.NUMERIC, sheet.getRow(2).getCell(index).getCellType());
                assertEquals("#,##0.00", sheet.getRow(2).getCell(index).getCellStyle().getDataFormatString());
            }
            assertEquals(CellType.NUMERIC, sheet.getRow(2).getCell(2).getCellType());
            assertEquals("yyyy-MM-dd", sheet.getRow(2).getCell(2).getCellStyle().getDataFormatString());
            assertEquals(25 * 256, sheet.getColumnWidth(8));
        }
    }

    private CachedRowSet totalRow() throws Exception {
        CachedRowSet rowSet = RowSetProvider.newFactory().createCachedRowSet();
        RowSetMetaDataImpl metadata = new RowSetMetaDataImpl();
        metadata.setColumnCount(33);
        for (int index = 1; index <= 33; index++) {
            metadata.setColumnName(index, "열" + index);
            metadata.setColumnLabel(index, "열" + index);
            metadata.setColumnType(index, Types.VARCHAR);
        }
        metadata.setColumnType(3, Types.DATE);
        metadata.setColumnType(13, Types.INTEGER);
        for (int index = 18; index <= 22; index++) {
            metadata.setColumnLabel(index, switch (index) {
                case 18 -> "판매기준가";
                case 19 -> "판매단가";
                case 20 -> "판매금액";
                case 21 -> "부가세액";
                default -> "판매금액계";
            });
            metadata.setColumnType(index, Types.DECIMAL);
        }
        rowSet.setMetaData(metadata);
        rowSet.moveToInsertRow();
        for (int index = 1; index <= 33; index++) {
            rowSet.updateString(index, "");
        }
        rowSet.updateString(1, "TOTAL");
        rowSet.updateDate(3, java.sql.Date.valueOf("2026-09-01"));
        rowSet.updateInt(13, 7405);
        rowSet.updateDouble(18, 369000);
        rowSet.updateDouble(19, 369000);
        rowSet.updateDouble(20, 2215054.5);
        rowSet.updateDouble(21, 4027372);
        rowSet.updateDouble(22, 123.45);
        rowSet.insertRow();
        rowSet.moveToCurrentRow();
        rowSet.beforeFirst();
        return rowSet;
    }
}
