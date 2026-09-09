package com.company.erpbatch.excel;

import com.company.erpbatch.exception.BatchException;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Timestamp;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Set;

/** ResultSet을 일반적인 Excel 통합 문서로 저장합니다. */
public final class ExcelService {
    private static final Logger log = LoggerFactory.getLogger(ExcelService.class);
    private static final DateTimeFormatter TIMESTAMP_FORMAT = DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss");
    private static final int REPORT_COLUMN_COUNT = 33;
    private static final int[] COLUMN_WIDTHS = {16, 16, 13, 18, 16, 16, 16, 16, 25, 16, 16, 8, 16, 8, 16, 8,
            16, 16, 16, 16, 16, 16, 16, 16, 16, 16, 16, 16, 16, 10, 15, 16, 10};
    private static final XSSFColor HEADER_COLOR = color(0x03, 0xB0, 0x51);
    private static final XSSFColor TOTAL_COLOR = color(0xAC, 0xFA, 0xB8);
    private static final XSSFColor BORDER_COLOR = color(0xAA, 0xAA, 0xAA);
    private static final Set<String> TWO_DECIMAL_HEADERS = Set.of("판매기준가", "판매단가", "판매금액", "부가세액", "판매금액계");

    public ExcelWriteResult write(ResultSet resultSet, Path outputDirectory, String fileNamePattern) throws SQLException {
        Path outputFile = resolveOutputPath(outputDirectory, fileNamePattern, LocalDateTime.now());
        log.info("Excel 생성을 시작합니다: {}", outputFile.toAbsolutePath());
        try {
            Files.createDirectories(outputDirectory);
            try (XSSFWorkbook workbook = new XSSFWorkbook()) {
                XSSFSheet sheet = workbook.createSheet("Report");
                List<ExcelColumn> columns = columns(resultSet.getMetaData());
                writeTitle(sheet, workbook);
                writeHeader(sheet, columns, workbook);
                long rowCount = writeRows(sheet, columns, resultSet, workbook);
                applyColumnWidths(sheet);
                try (var output = Files.newOutputStream(outputFile)) {
                    workbook.write(output);
                }
                log.info("Excel 생성이 완료되었습니다. 조회 건수: {}, 파일: {}", rowCount, outputFile.toAbsolutePath());
                return new ExcelWriteResult(outputFile, rowCount);
            }
        } catch (IOException exception) {
            throw new BatchException("Excel 파일을 저장할 수 없습니다: " + outputFile.toAbsolutePath(), exception);
        }
    }

    public Path resolveOutputPath(Path outputDirectory, String pattern, LocalDateTime timestamp) {
        if (pattern == null || pattern.isBlank()) throw new BatchException("Excel 파일명 패턴이 비어 있습니다.");
        String fileName = pattern.replace("{timestamp}", TIMESTAMP_FORMAT.format(timestamp));
        if (!fileName.toLowerCase().endsWith(".xlsx")) fileName += ".xlsx";
        return outputDirectory.resolve(fileName);
    }

    private List<ExcelColumn> columns(ResultSetMetaData metadata) throws SQLException {
        List<ExcelColumn> columns = new ArrayList<>();
        for (int index = 1; index <= metadata.getColumnCount(); index++) {
            columns.add(new ExcelColumn(index, metadata.getColumnLabel(index)));
        }
        return columns;
    }

    private void writeTitle(XSSFSheet sheet, XSSFWorkbook workbook) {
        Row row = sheet.createRow(0);
        row.setHeightInPoints(50);
        Cell cell = row.createCell(0);
        cell.setCellValue("일자별판매실적분석");
        cell.setCellStyle(createTitleStyle(workbook));
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, REPORT_COLUMN_COUNT - 1));
    }

    private void writeHeader(XSSFSheet sheet, List<ExcelColumn> columns, XSSFWorkbook workbook) {
        CellStyle style = createHeaderStyle(workbook);
        Row row = sheet.createRow(1);
        row.setHeightInPoints(20);
        for (int index = 0; index < columns.size(); index++) {
            Cell cell = row.createCell(index);
            cell.setCellValue(columns.get(index).header());
            cell.setCellStyle(style);
        }
    }

    private long writeRows(XSSFSheet sheet, List<ExcelColumn> columns, ResultSet resultSet, XSSFWorkbook workbook) throws SQLException {
        long count = 0;
        while (resultSet.next()) {
            Row row = sheet.createRow((int) count + 2);
            row.setHeightInPoints(20);
            boolean totalRow = "TOTAL".equals(resultSet.getString(1));
            for (int index = 0; index < columns.size(); index++) {
                Cell cell = row.createCell(index);
                Object value = resultSet.getObject(columns.get(index).index());
                writeCellValue(cell, value);
                cell.setCellStyle(styleFor(value, columns.get(index), totalRow, workbook));
            }
            count++;
        }
        return count;
    }

    private CellStyle createTitleStyle(XSSFWorkbook workbook) {
        Font font = workbook.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 20);
        XSSFCellStyle style = workbook.createCellStyle();
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        return style;
    }

    private CellStyle createHeaderStyle(XSSFWorkbook workbook) {
        Font font = workbook.createFont();
        font.setBold(true);
        font.setColor(IndexedColors.WHITE.getIndex());
        XSSFCellStyle style = workbook.createCellStyle();
        style.setFont(font);
        style.setFillForegroundColor(HEADER_COLOR);
        style.setFillPattern(org.apache.poi.ss.usermodel.FillPatternType.SOLID_FOREGROUND);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setWrapText(true);
        applyThinBorder(style);
        return style;
    }

    private CellStyle styleFor(Object value, ExcelColumn column, boolean totalRow, XSSFWorkbook workbook) {
        if (value instanceof Number number) {
            int decimalPlaces = TWO_DECIMAL_HEADERS.contains(column.header()) ? 2 : decimalPlaces(number);
            return createNumberStyle(workbook, totalRow, decimalPlaces);
        }
        if (isDateValue(value)) {
            return createDateStyle(workbook, totalRow);
        }
        return createBodyStyle(workbook, totalRow, HorizontalAlignment.LEFT);
    }

    private CellStyle createBodyStyle(XSSFWorkbook workbook, boolean totalRow, HorizontalAlignment alignment) {
        Font font = workbook.createFont();
        font.setFontHeightInPoints((short) 10);
        XSSFCellStyle style = workbook.createCellStyle();
        style.setFont(font);
        style.setAlignment(alignment);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        if (totalRow) {
            style.setFillForegroundColor(TOTAL_COLOR);
            style.setFillPattern(org.apache.poi.ss.usermodel.FillPatternType.SOLID_FOREGROUND);
        }
        applyThinBorder(style);
        return style;
    }

    private CellStyle createNumberStyle(XSSFWorkbook workbook, boolean totalRow, int decimalPlaces) {
        XSSFCellStyle style = (XSSFCellStyle) createBodyStyle(workbook, totalRow, HorizontalAlignment.RIGHT);
        String format = decimalPlaces == 0 ? "#,##0" : "#,##0." + "0".repeat(decimalPlaces);
        style.setDataFormat(workbook.createDataFormat().getFormat(format));
        return style;
    }

    private CellStyle createDateStyle(XSSFWorkbook workbook, boolean totalRow) {
        XSSFCellStyle style = (XSSFCellStyle) createBodyStyle(workbook, totalRow, HorizontalAlignment.LEFT);
        style.setDataFormat(workbook.createDataFormat().getFormat("yyyy-MM-dd"));
        return style;
    }

    private void applyThinBorder(XSSFCellStyle style) {
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setTopBorderColor(BORDER_COLOR);
        style.setBottomBorderColor(BORDER_COLOR);
        style.setLeftBorderColor(BORDER_COLOR);
        style.setRightBorderColor(BORDER_COLOR);
    }

    private void writeCellValue(Cell cell, Object value) {
        if (value == null) return;
        if (value instanceof Number number) {
            cell.setCellValue(number.doubleValue());
        } else if (value instanceof LocalDate localDate) {
            cell.setCellValue(localDate);
        } else if (value instanceof LocalDateTime localDateTime) {
            cell.setCellValue(localDateTime);
        } else if (value instanceof Timestamp timestamp) {
            cell.setCellValue(timestamp);
        } else if (value instanceof Date date) {
            cell.setCellValue(date);
        } else {
            cell.setCellValue(String.valueOf(value));
        }
    }

    private boolean isDateValue(Object value) {
        return value instanceof Date || value instanceof LocalDate || value instanceof LocalDateTime;
    }

    private int decimalPlaces(Number value) {
        BigDecimal decimal = value instanceof BigDecimal bigDecimal
                ? bigDecimal.stripTrailingZeros()
                : BigDecimal.valueOf(value.doubleValue()).stripTrailingZeros();
        return Math.max(0, decimal.scale());
    }

    private void applyColumnWidths(XSSFSheet sheet) {
        for (int index = 0; index < COLUMN_WIDTHS.length; index++) {
            sheet.setColumnWidth(index, COLUMN_WIDTHS[index] * 256);
        }
    }

    private static XSSFColor color(int red, int green, int blue) {
        return new XSSFColor(new byte[]{(byte) red, (byte) green, (byte) blue}, new DefaultIndexedColorMap());
    }

    public record ExcelWriteResult(Path outputFile, long rowCount) { }
}
