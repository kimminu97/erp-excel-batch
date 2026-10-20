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

    public <T> List<T> executeAll(String procedureName, List<?> parameters, IndexedResultSetHandler<T> handler) {
        String callSyntax = createCallSyntax(procedureName, parameters.size());
        log.info("저장 프로시저 실행을 시작합니다: {}", procedureName);
        try (Connection connection = databaseConnection.open();
             CallableStatement statement = connection.prepareCall(callSyntax)) {
            for (int index = 0; index < parameters.size(); index++) {
                statement.setObject(index + 1, parameters.get(index));
            }
            List<T> results = consumeAllResultSets(statement, handler);
            log.info("저장 프로시저 실행이 완료되었습니다: {}. ResultSet 수: {}", procedureName, results.size());
            return results;
        } catch (SQLException exception) {
            throw new BatchException("저장 프로시저 실행 또는 ResultSet 처리에 실패했습니다: " + procedureName, exception);
        }
    }

    static <T> List<T> consumeAllResultSets(CallableStatement statement, IndexedResultSetHandler<T> handler) throws SQLException {
        List<T> results = new java.util.ArrayList<>();
        boolean hasResultSet = statement.execute();
        while (hasResultSet || statement.getUpdateCount() != -1) {
            if (hasResultSet) {
                try (ResultSet resultSet = statement.getResultSet()) {
                    results.add(handler.handle(results.size(), resultSet));
                }
            }
            hasResultSet = statement.getMoreResults();
        }
        return results;
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
    public interface IndexedResultSetHandler<T> {
        T handle(int resultSetIndex, ResultSet resultSet) throws SQLException;
    }
}
