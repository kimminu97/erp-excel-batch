package com.company.erpbatch.exception;

/** 배치 처리 중 발생한 복구 불가능한 오류를 표현합니다. */
public class BatchException extends RuntimeException {
    public BatchException(String message) {
        super(message);
    }

    public BatchException(String message, Throwable cause) {
        super(message, cause);
    }
}
