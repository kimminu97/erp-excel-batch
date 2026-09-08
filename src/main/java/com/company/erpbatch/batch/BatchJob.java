package com.company.erpbatch.batch;

import com.company.erpbatch.config.BatchConfig;
import com.company.erpbatch.db.StoredProcedureExecutor;
import com.company.erpbatch.excel.ExcelService;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.List;

/** 현재는 매개변수 없는 일반 저장 프로시저를 실행하는 기본 배치입니다. */
public final class BatchJob {
    private static final Logger log = LoggerFactory.getLogger(BatchJob.class);
    private final BatchConfig config;
    private final StoredProcedureExecutor executor;
    private final ExcelService excelService;

    public BatchJob(BatchConfig config, StoredProcedureExecutor executor, ExcelService excelService) {
        this.config = config;
        this.executor = executor;
        this.excelService = excelService;
    }

    public ExcelService.ExcelWriteResult run() {
        log.info("배치 작업을 시작합니다.");
        ExcelService.ExcelWriteResult result = executor.execute(
                config.storedProcedureName(), List.of(),
                resultSet -> excelService.write(resultSet, config.outputDirectory(), config.outputFileNamePattern()));
        log.info("배치 작업이 성공했습니다. 출력 파일: {}", result.outputFile().toAbsolutePath());
        return result;
    }
}
