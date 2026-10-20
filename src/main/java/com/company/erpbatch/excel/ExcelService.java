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
    private static final DateTimeFormatter TIMESTAMP_FORMAT = DateTimeFormatter.ofPattern("yyMMdd");
    private static final int REPORT_COLUMN_COUNT = 33;
    private static final int[] COLUMN_WIDTHS = {16, 16, 13, 18, 16, 16, 16, 16, 25, 16, 16, 8, 16, 8, 16, 8,
            16, 16, 16, 16, 16, 16, 16, 16, 16, 16, 16, 16, 16, 10, 15, 16, 10};
    private static final XSSFColor HEADER_COLOR = color(0x03, 0xB0, 0x51);
    private static final XSSFColor TOTAL_COLOR = color(0xAC, 0xFA, 0xB8);
    private static final XSSFColor BORDER_COLOR = color(0xAA, 0xAA, 0xAA);
    private static final Set<String> TWO_DECIMAL_HEADERS = Set.of("판매기준가", "판매단가", "판매금액", "부가세액", "판매금액계");

    public ExcelWriteResult write(ResultSet resultSet, Path outputDirectory, String fileNamePattern) throws SQLException {
        return write(resultSet, outputDirectory, fileNamePattern, LocalDateTime.now());
    }

    public ExcelWriteResult write(ResultSet resultSet, Path outputDirectory, String fileNamePattern, LocalDateTime timestamp) throws SQLException {
        Path outputFile = resolveOutputPath(outputDirectory, fileNamePattern, timestamp);
        log.info("Excel 생성을 시작합니다: {}", outputFile.toAbsolutePath());
        try {
            Files.createDirectories(outputDirectory);
            try (XSSFWorkbook workbook = new XSSFWorkbook()) {
                XSSFSheet sheet = workbook.createSheet("Report");
                List<ExcelColumn> columns = columns(resultSet.getMetaData());
                ExcelStyles styles = new ExcelStyles(workbook);
                writeTitle(sheet, styles.title());
                writeHeader(sheet, columns, styles.header());
                long rowCount = writeRows(sheet, columns, resultSet, styles);
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

    private void writeTitle(XSSFSheet sheet, CellStyle titleStyle) {
        Row row = sheet.createRow(0);
        row.setHeightInPoints(50);
        Cell cell = row.createCell(0);
        cell.setCellValue("일자별판매실적분석");
        cell.setCellStyle(titleStyle);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, REPORT_COLUMN_COUNT - 1));
    }

    private void writeHeader(XSSFSheet sheet, List<ExcelColumn> columns, CellStyle headerStyle) {
        Row row = sheet.createRow(1);
        row.setHeightInPoints(20);
        for (int index = 0; index < columns.size(); index++) {
            Cell cell = row.createCell(index);
            cell.setCellValue(columns.get(index).header());
            cell.setCellStyle(headerStyle);
        }
    }

    private long writeRows(XSSFSheet sheet, List<ExcelColumn> columns, ResultSet resultSet, ExcelStyles styles) throws SQLException {
        long count = 0;
        while (resultSet.next()) {
            Row row = sheet.createRow((int) count + 2);
            row.setHeightInPoints(20);
            boolean totalRow = "TOTAL".equals(resultSet.getString(1));
            for (int index = 0; index < columns.size(); index++) {
                Cell cell = row.createCell(index);
                Object value = resultSet.getObject(columns.get(index).index());
                writeCellValue(cell, value);
                cell.setCellStyle(styleFor(value, columns.get(index), totalRow, styles));
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

    private CellStyle styleFor(Object value, ExcelColumn column, boolean totalRow, ExcelStyles styles) {
        if (value instanceof Number) {
            if (TWO_DECIMAL_HEADERS.contains(column.header())) {
                return totalRow ? styles.totalMoney() : styles.bodyMoney();
            }
            if (decimalPlaces((Number) value) == 0) {
                return totalRow ? styles.totalNumber() : styles.bodyNumber();
            }
            return totalRow ? styles.totalDecimal() : styles.bodyDecimal();
        }
        if (isDateValue(value)) {
            return totalRow ? styles.totalDate() : styles.bodyDate();
        }
        return totalRow ? styles.totalText() : styles.bodyText();
    }

    private CellStyle createBodyStyle(XSSFWorkbook workbook, Font font, boolean totalRow, HorizontalAlignment alignment) {
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

    private final class ExcelStyles {
        private final CellStyle title;
        private final CellStyle header;
        private final CellStyle totalText;
        private final CellStyle totalDate;
        private final CellStyle totalNumber;
        private final CellStyle totalDecimal;
        private final CellStyle totalMoney;
        private final CellStyle bodyText;
        private final CellStyle bodyDate;
        private final CellStyle bodyNumber;
        private final CellStyle bodyDecimal;
        private final CellStyle bodyMoney;

        private ExcelStyles(XSSFWorkbook workbook) {
            Font bodyFont = workbook.createFont();
            bodyFont.setFontHeightInPoints((short) 10);
            short dateFormat = workbook.createDataFormat().getFormat("yyyy-MM-dd");
            short integerFormat = workbook.createDataFormat().getFormat("#,##0");
            short numberFormat = workbook.createDataFormat().getFormat("#,##0.###############");
            short moneyFormat = workbook.createDataFormat().getFormat("#,##0.00");

            bodyText = createBodyStyle(workbook, bodyFont, false, HorizontalAlignment.LEFT);
            bodyDate = createBodyStyle(workbook, bodyFont, false, HorizontalAlignment.LEFT);
            bodyDate.setDataFormat(dateFormat);
            bodyNumber = createBodyStyle(workbook, bodyFont, false, HorizontalAlignment.RIGHT);
            bodyNumber.setDataFormat(integerFormat);
            bodyDecimal = createBodyStyle(workbook, bodyFont, false, HorizontalAlignment.RIGHT);
            bodyDecimal.setDataFormat(numberFormat);
            bodyMoney = createBodyStyle(workbook, bodyFont, false, HorizontalAlignment.RIGHT);
            bodyMoney.setDataFormat(moneyFormat);
            totalText = createBodyStyle(workbook, bodyFont, true, HorizontalAlignment.LEFT);
            totalDate = createBodyStyle(workbook, bodyFont, true, HorizontalAlignment.LEFT);
            totalDate.setDataFormat(dateFormat);
            totalNumber = createBodyStyle(workbook, bodyFont, true, HorizontalAlignment.RIGHT);
            totalNumber.setDataFormat(integerFormat);
            totalDecimal = createBodyStyle(workbook, bodyFont, true, HorizontalAlignment.RIGHT);
            totalDecimal.setDataFormat(numberFormat);
            totalMoney = createBodyStyle(workbook, bodyFont, true, HorizontalAlignment.RIGHT);
            totalMoney.setDataFormat(moneyFormat);
            title = createTitleStyle(workbook);
            header = createHeaderStyle(workbook);
        }

        private CellStyle title() { return title; }
        private CellStyle header() { return header; }
        private CellStyle totalText() { return totalText; }
        private CellStyle totalDate() { return totalDate; }
        private CellStyle totalNumber() { return totalNumber; }
        private CellStyle totalDecimal() { return totalDecimal; }
        private CellStyle totalMoney() { return totalMoney; }
        private CellStyle bodyText() { return bodyText; }
        private CellStyle bodyDate() { return bodyDate; }
        private CellStyle bodyNumber() { return bodyNumber; }
        private CellStyle bodyDecimal() { return bodyDecimal; }
        private CellStyle bodyMoney() { return bodyMoney; }
    }

    public record ExcelWriteResult(Path outputFile, long rowCount) { }
}
