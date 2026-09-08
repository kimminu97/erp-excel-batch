package com.company.erpbatch.excel;

import com.company.erpbatch.exception.BatchException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
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
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;

/** ResultSet을 일반적인 Excel 통합 문서로 저장합니다. */
public final class ExcelService {
    private static final Logger log = LoggerFactory.getLogger(ExcelService.class);
    private static final DateTimeFormatter TIMESTAMP_FORMAT = DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss");

    public ExcelWriteResult write(ResultSet resultSet, Path outputDirectory, String fileNamePattern) throws SQLException {
        Path outputFile = resolveOutputPath(outputDirectory, fileNamePattern, LocalDateTime.now());
        log.info("Excel 생성을 시작합니다: {}", outputFile.toAbsolutePath());
        try {
            Files.createDirectories(outputDirectory);
            try (XSSFWorkbook workbook = new XSSFWorkbook()) {
                XSSFSheet sheet = workbook.createSheet("Report");
                List<ExcelColumn> columns = columns(resultSet.getMetaData());
                writeHeader(sheet, columns, workbook);
                long rowCount = writeRows(sheet, columns, resultSet);
                adjustColumnWidths(sheet, columns.size());
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

    private void writeHeader(XSSFSheet sheet, List<ExcelColumn> columns, XSSFWorkbook workbook) {
        Font font = workbook.createFont();
        font.setBold(true);
        CellStyle style = workbook.createCellStyle();
        style.setFont(font);
        Row row = sheet.createRow(0);
        for (int index = 0; index < columns.size(); index++) {
            Cell cell = row.createCell(index);
            cell.setCellValue(columns.get(index).header());
            cell.setCellStyle(style);
        }
    }

    private long writeRows(XSSFSheet sheet, List<ExcelColumn> columns, ResultSet resultSet) throws SQLException {
        long count = 0;
        while (resultSet.next()) {
            Row row = sheet.createRow((int) ++count);
            for (int index = 0; index < columns.size(); index++) {
                Object value = resultSet.getObject(columns.get(index).index());
                if (value != null) row.createCell(index).setCellValue(String.valueOf(value));
            }
        }
        return count;
    }

    private void adjustColumnWidths(XSSFSheet sheet, int columnCount) {
        for (int index = 0; index < columnCount; index++) {
            sheet.autoSizeColumn(index);
            sheet.setColumnWidth(index, Math.min(sheet.getColumnWidth(index) + 512, 12000));
        }
    }

    public record ExcelWriteResult(Path outputFile, long rowCount) { }
}
