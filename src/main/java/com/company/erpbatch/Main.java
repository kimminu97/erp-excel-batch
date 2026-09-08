package com.company.erpbatch;

import com.company.erpbatch.batch.BatchJob;
import com.company.erpbatch.config.BatchConfig;
import com.company.erpbatch.db.DatabaseConnection;
import com.company.erpbatch.db.StoredProcedureExecutor;
import com.company.erpbatch.excel.ExcelService;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/** 명령행에서 실행되는 ERP Excel 배치의 진입점입니다. */
public final class Main {
    private static final Logger log = LoggerFactory.getLogger(Main.class);

    private Main() { }

    public static void main(String[] args) {
        int exitCode = run(args);
        System.exit(exitCode);
    }

    static int run(String[] args) {
        try {
            log.info("ERP Excel 배치를 시작합니다.");
            BatchConfig config = BatchConfig.load(args);
            BatchJob job = new BatchJob(config, new StoredProcedureExecutor(new DatabaseConnection(config)), new ExcelService());
            job.run();
            log.info("ERP Excel 배치가 정상 종료되었습니다.");
            return 0;
        } catch (Exception exception) {
            log.error("ERP Excel 배치가 실패했습니다: {}", exception.getMessage(), exception);
            return 1;
        }
    }
}
