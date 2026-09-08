package com.company.erpbatch.db;

import com.company.erpbatch.config.BatchConfig;
import com.company.erpbatch.exception.BatchException;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;

/** SQL Server 연결 생성을 한 곳에서 관리합니다. */
public final class DatabaseConnection {
    private static final Logger log = LoggerFactory.getLogger(DatabaseConnection.class);
    private final BatchConfig config;

    public DatabaseConnection(BatchConfig config) {
        this.config = config;
    }

    public Connection open() {
        try {
            log.info("데이터베이스 연결을 시작합니다.");
            Connection connection = DriverManager.getConnection(
                    config.databaseUrl(), config.databaseUsername(), config.databasePassword());
            log.info("데이터베이스 연결에 성공했습니다.");
            return connection;
        } catch (SQLException exception) {
            throw new BatchException("데이터베이스에 연결할 수 없습니다.", exception);
        }
    }
}
