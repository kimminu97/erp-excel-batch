package com.company.erpbatch.db;

import org.junit.jupiter.api.Test;

import javax.sql.rowset.CachedRowSet;
import javax.sql.rowset.RowSetMetaDataImpl;
import javax.sql.rowset.RowSetProvider;
import java.lang.reflect.Proxy;
import java.sql.CallableStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Types;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;

class StoredProcedureExecutorTest {
    @Test
    void consumesAllResultSetsInStoredProcedureOrder() throws Exception {
        CallableStatement statement = statement(resultSet("당월"), resultSet("전년동월"), resultSet("전월"));

        List<String> values = StoredProcedureExecutor.consumeAllResultSets(
                statement, (index, resultSet) -> index + ":" + firstValue(resultSet));

        assertEquals(List.of("0:당월", "1:전년동월", "2:전월"), values);
    }

    private CallableStatement statement(ResultSet... resultSets) {
        final int[] current = {0};
        return (CallableStatement) Proxy.newProxyInstance(getClass().getClassLoader(), new Class[]{CallableStatement.class}, (proxy, method, args) -> {
            return switch (method.getName()) {
                case "execute" -> true;
                case "getResultSet" -> current[0] < resultSets.length ? resultSets[current[0]] : null;
                case "getMoreResults" -> ++current[0] < resultSets.length;
                case "getUpdateCount" -> -1;
                default -> throw new UnsupportedOperationException(method.getName());
            };
        });
    }

    private CachedRowSet resultSet(String value) throws Exception {
        CachedRowSet rowSet = RowSetProvider.newFactory().createCachedRowSet();
        RowSetMetaDataImpl metadata = new RowSetMetaDataImpl();
        metadata.setColumnCount(1);
        metadata.setColumnName(1, "구분");
        metadata.setColumnType(1, Types.VARCHAR);
        rowSet.setMetaData(metadata);
        rowSet.moveToInsertRow();
        rowSet.updateString(1, value);
        rowSet.insertRow();
        rowSet.moveToCurrentRow();
        rowSet.beforeFirst();
        return rowSet;
    }

    private String firstValue(ResultSet resultSet) throws SQLException {
        resultSet.next();
        return resultSet.getString(1);
    }
}
