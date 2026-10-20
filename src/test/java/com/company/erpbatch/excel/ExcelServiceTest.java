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
import static org.junit.jupiter.api.Assertions.assertTrue;

class ExcelServiceTest {
    @TempDir
    Path tempDir;

    @Test
    void replacesTimestampAndAddsExtension() {
        ExcelService service = new ExcelService();
        Path path = service.resolveOutputPath(Path.of("C:/output"), "GALA_일자별판매실적분석_{timestamp}.xlsx", LocalDateTime.of(2026, 9, 9, 13, 51));
        assertEquals(Path.of("C:/output/GALA_일자별판매실적분석_260909.xlsx"), path);
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

    @Test
    void reusesCellStylesForThousandsOfRows() throws Exception {
        ExcelService service = new ExcelService();
        ExcelService.ExcelWriteResult result = service.write(largeRows(3_000), tempDir, "large-sales.xlsx");

        assertEquals(3_000, result.rowCount());
        try (XSSFWorkbook workbook = new XSSFWorkbook(result.outputFile().toFile())) {
            assertTrue(workbook.getNumCellStyles() <= 20,
                    () -> "행 수와 무관하게 제한된 스타일만 생성해야 합니다. 실제 스타일 수: " + workbook.getNumCellStyles());
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

    private CachedRowSet largeRows(int rowCount) throws Exception {
        CachedRowSet rowSet = RowSetProvider.newFactory().createCachedRowSet();
        RowSetMetaDataImpl metadata = new RowSetMetaDataImpl();
        metadata.setColumnCount(5);
        metadata.setColumnName(1, "구분");
        metadata.setColumnLabel(1, "구분");
        metadata.setColumnType(1, Types.VARCHAR);
        metadata.setColumnName(2, "일자");
        metadata.setColumnLabel(2, "일자");
        metadata.setColumnType(2, Types.DATE);
        metadata.setColumnName(3, "수량");
        metadata.setColumnLabel(3, "수량");
        metadata.setColumnType(3, Types.INTEGER);
        metadata.setColumnName(4, "판매금액");
        metadata.setColumnLabel(4, "판매금액");
        metadata.setColumnType(4, Types.DECIMAL);
        metadata.setColumnName(5, "비율");
        metadata.setColumnLabel(5, "비율");
        metadata.setColumnType(5, Types.DECIMAL);
        rowSet.setMetaData(metadata);

        for (int index = 0; index < rowCount; index++) {
            rowSet.moveToInsertRow();
            rowSet.updateString(1, "일반");
            rowSet.updateDate(2, java.sql.Date.valueOf("2026-09-01"));
            rowSet.updateInt(3, index);
            rowSet.updateDouble(4, 1_000.50);
            rowSet.updateDouble(5, 12.345);
            rowSet.insertRow();
        }
        rowSet.moveToCurrentRow();
        rowSet.beforeFirst();
        return rowSet;
    }
}
