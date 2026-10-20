package com.company.erpbatch.config;

import com.company.erpbatch.exception.BatchException;

import java.io.IOException;
import java.io.BufferedReader;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.charset.StandardCharsets;
import java.util.Properties;

/** 외부 application.properties의 배치 설정을 제공합니다. */
public final class BatchConfig {
    private static final String DEFAULT_CONFIG_FILE = "application.properties";
    private final Properties properties;

    private BatchConfig(Properties properties) {
        this.properties = properties;
    }

    public static BatchConfig load(String[] args) {
        Path configPath = args.length == 0 ? Path.of(DEFAULT_CONFIG_FILE) : Path.of(args[0]);
        if (!Files.isRegularFile(configPath)) {
            throw new BatchException("설정 파일을 찾을 수 없습니다: " + configPath.toAbsolutePath());
        }
        Properties properties = new Properties();
        try (BufferedReader reader = Files.newBufferedReader(configPath, StandardCharsets.UTF_8)) {
            properties.load(reader);
        } catch (IOException exception) {
            throw new BatchException("설정 파일을 읽을 수 없습니다: " + configPath.toAbsolutePath(), exception);
        }
        return new BatchConfig(properties);
    }

    public String databaseUrl() { return required("db.url"); }
    public String databaseUsername() { return required("db.username"); }
    public String databasePassword() { return required("db.password"); }
    public String storedProcedureName() { return required("db.stored-procedure", "stored-procedure.name"); }
    public Path outputDirectory() { return Path.of(required("output.directory")); }
    public String outputFileNamePattern() { return required("excel.output-file-name-pattern", "output.file-name-pattern"); }

    private String required(String... keys) {
        for (String key : keys) {
            String value = properties.getProperty(key);
            if (value != null && !value.isBlank()) {
                return value.trim();
            }
        }
        throw new BatchException("필수 설정값이 비어 있습니다: " + keys[0]);
    }
}
