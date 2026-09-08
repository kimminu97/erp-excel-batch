package com.company.erpbatch.config;

import com.company.erpbatch.exception.BatchException;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.nio.file.Files;
import java.nio.file.Path;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

class BatchConfigTest {
    @TempDir Path tempDir;

    @Test
    void loadsExternalProperties() throws Exception {
        Path file = tempDir.resolve("application.properties");
        Files.writeString(file, "db.url=jdbc:test\ndb.username=user\ndb.password=secret\nstored-procedure.name=dbo.Report\noutput.directory=C:/output\noutput.file-name-pattern=report_{timestamp}.xlsx\n");
        BatchConfig config = BatchConfig.load(new String[]{file.toString()});
        assertEquals("jdbc:test", config.databaseUrl());
        assertEquals("dbo.Report", config.storedProcedureName());
    }

    @Test
    void failsWhenConfigFileDoesNotExist() {
        assertThrows(BatchException.class, () -> BatchConfig.load(new String[]{tempDir.resolve("missing.properties").toString()}));
    }
}
