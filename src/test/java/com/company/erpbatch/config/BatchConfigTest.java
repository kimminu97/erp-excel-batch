package com.company.erpbatch.config;

import com.company.erpbatch.exception.BatchException;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.charset.StandardCharsets;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

class BatchConfigTest {
    @TempDir Path tempDir;

    @Test
    void loadsExternalProperties() throws Exception {
        Path file = tempDir.resolve("application.properties");
        Files.writeString(file, "db.url=jdbc:test\ndb.username=user\ndb.password=secret\nstored-procedure.name=dbo.Report\noutput.directory=C:/output\noutput.file-name-pattern=GALA_일자별판매실적분석_{timestamp}.xlsx\n", StandardCharsets.UTF_8);
        BatchConfig config = BatchConfig.load(new String[]{file.toString()});
        assertEquals("jdbc:test", config.databaseUrl());
        assertEquals("dbo.Report", config.storedProcedureName());
        assertEquals("GALA_일자별판매실적분석_{timestamp}.xlsx", config.outputFileNamePattern());
    }

    @Test
    void failsWhenConfigFileDoesNotExist() {
        assertThrows(BatchException.class, () -> BatchConfig.load(new String[]{tempDir.resolve("missing.properties").toString()}));
    }
}
