package com.company.erpbatch.excel;

import org.junit.jupiter.api.Test;

import java.nio.file.Path;
import java.time.LocalDateTime;

import static org.junit.jupiter.api.Assertions.assertEquals;

class ExcelServiceTest {
    @Test
    void replacesTimestampAndAddsExtension() {
        ExcelService service = new ExcelService();
        Path path = service.resolveOutputPath(Path.of("C:/output"), "report_{timestamp}", LocalDateTime.of(2026, 9, 8, 1, 2, 3));
        assertEquals(Path.of("C:/output/report_20260908_010203.xlsx"), path);
    }
}
