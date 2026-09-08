package com.company.erpbatch.db;

import com.company.erpbatch.exception.BatchException;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.sql.CallableStatement;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.List;

/** 저장 프로시저 결과를 연결이 열린 상태에서 소비하도록 실행합니다. */
public final class StoredProcedureExecutor {
    private static final Logger log = LoggerFactory.getLogger(StoredProcedureExecutor.class);
    private final DatabaseConnection databaseConnection;

    public StoredProcedureExecutor(DatabaseConnection databaseConnection) {
        this.databaseConnection = databaseConnection;
    }

    public <T> T execute(String procedureName, List<?> parameters, ResultSetHandler<T> handler) {
        String callSyntax = createCallSyntax(procedureName, parameters.size());
        log.info("저장 프로시저 실행을 시작합니다: {}", procedureName);
        try (Connection connection = databaseConnection.open();
             CallableStatement statement = connection.prepareCall(callSyntax)) {
            for (int index = 0; index < parameters.size(); index++) {
                statement.setObject(index + 1, parameters.get(index));
            }
            boolean hasResultSet = statement.execute();
            while (!hasResultSet && statement.getUpdateCount() != -1) {
                hasResultSet = statement.getMoreResults();
            }
            if (!hasResultSet) {
                throw new BatchException("저장 프로시저가 ResultSet을 반환하지 않았습니다: " + procedureName);
            }
            try (ResultSet resultSet = statement.getResultSet()) {
                T result = handler.handle(resultSet);
                log.info("저장 프로시저 실행이 완료되었습니다: {}", procedureName);
                return result;
            }
        } catch (SQLException exception) {
            throw new BatchException("저장 프로시저 실행 또는 ResultSet 처리에 실패했습니다: " + procedureName, exception);
        }
    }

    private String createCallSyntax(String procedureName, int parameterCount) {
        if (procedureName == null || procedureName.isBlank()) {
            throw new BatchException("저장 프로시저 이름이 비어 있습니다.");
        }
        String placeholders = parameterCount == 0 ? "" : "?".repeat(parameterCount).replace("?", "?,");
        if (!placeholders.isEmpty()) placeholders = placeholders.substring(0, placeholders.length() - 1);
        return "{call " + procedureName + "(" + placeholders + ")}";
    }

    @FunctionalInterface
    public interface ResultSetHandler<T> {
        T handle(ResultSet resultSet) throws SQLException;
    }
}
