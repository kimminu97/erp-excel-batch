package com.company.erpbatch.batch;

import com.company.erpbatch.config.BatchConfig;
import com.company.erpbatch.db.StoredProcedureExecutor;
import com.company.erpbatch.excel.ExcelService;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.company.erpbatch.exception.BatchException;
import java.time.LocalDateTime;
import java.util.List;

/** 현재는 매개변수 없는 일반 저장 프로시저를 실행하는 기본 배치입니다. */
public final class BatchJob {
    private static final Logger log = LoggerFactory.getLogger(BatchJob.class);
    private static final List<String> RESULT_SET_NAMES = List.of("당월", "전년동월", "전월");
    private final BatchConfig config;
    private final StoredProcedureExecutor executor;
    private final ExcelService excelService;

    public BatchJob(BatchConfig config, StoredProcedureExecutor executor, ExcelService excelService) {
        this.config = config;
        this.executor = executor;
        this.excelService = excelService;
    }

    public List<ExcelService.ExcelWriteResult> run() {
        log.info("배치 작업을 시작합니다.");
        LocalDateTime executionTime = LocalDateTime.now();
        List<ExcelService.ExcelWriteResult> results = executor.executeAll(
                config.storedProcedureName(), List.of(),
                (resultSetIndex, resultSet) -> writeResultSet(resultSetIndex, resultSet, executionTime));
        if (results.size() != RESULT_SET_NAMES.size()) {
            throw new BatchException("예상 ResultSet 수: " + RESULT_SET_NAMES.size() + ", 실제 ResultSet 수: " + results.size());
        }
        log.info("총 {}개의 Excel 파일 생성이 완료되었습니다.", results.size());
        return results;
    }

    private ExcelService.ExcelWriteResult writeResultSet(int resultSetIndex, java.sql.ResultSet resultSet, LocalDateTime executionTime) throws java.sql.SQLException {
        if (resultSetIndex >= RESULT_SET_NAMES.size()) {
            return null;
        }
        String resultSetName = RESULT_SET_NAMES.get(resultSetIndex);
        ExcelService.ExcelWriteResult result = excelService.write(
                resultSet, config.outputDirectory(), config.outputFileNamePattern(), executionTime, resultSetName);
        log.info("{} Excel 생성 완료: {}", resultSetName, result.outputFile().toAbsolutePath());
        return result;
    }
}
