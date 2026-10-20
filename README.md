# ERP Excel Batch

영림원 ERP에서 반복적으로 수행하는 Excel 다운로드 업무를 자동화하기 위한 **Java 기반 사내 배치 프로그램**입니다.

기존 ERP 소스는 수정하지 않고, MSSQL Stored Procedure로 데이터를 조회한 뒤 Java에서 ERP 형식의 Excel 파일을 생성하도록 구성했습니다.

---

## 프로젝트 배경

기존 업무에서는 ERP 화면에서 조회 조건을 입력한 뒤 Excel 파일을 반복적으로 다운로드해야 했습니다.

ERP 소스를 직접 수정하기 어려운 환경이므로 다음과 같이 자동화 구조를 설계했습니다.

```text
영림원 ERP 화면 분석
        ↓
MSSQL Stored Procedure
        ↓
Java Batch
        ↓
JDBC ResultSet
        ↓
Apache POI
        ↓
ERP 형식 Excel 생성
        ↓
Windows 공유 폴더
        ↓
Windows 작업 스케줄러
```

Java 애플리케이션은 데이터 조회 및 Excel 생성 후 종료되는 경량 Command-line Batch 방식입니다.

---

## 기술 스택

- Java 17
- Maven
- Microsoft SQL Server
- JDBC
- Apache POI
- SLF4J + Logback
- JUnit 5
- Maven Shade Plugin
- Windows

---

## 주요 구현

### MSSQL Stored Procedure 연동

JDBC `CallableStatement`를 이용해 실제 MSSQL Stored Procedure를 실행합니다.

업무 조회 조건, 데이터 가공, TOTAL 계산 등은 Stored Procedure에서 처리하고 Java는 조회 결과를 받아 Excel로 출력하는 역할에 집중하도록 구성했습니다.

```text
Main
 → BatchConfig
 → BatchJob
 → DatabaseConnection
 → StoredProcedureExecutor
 → ResultSet
 → ExcelService
```

---

### 다중 ResultSet 처리

하나의 Stored Procedure 실행 결과로 여러 ResultSet이 반환되는 경우를 처리할 수 있도록 구현했습니다.

현재 `일자별판매실적분석 3PP` 배치는 한 번의 프로시저 실행으로 다음 3개의 ResultSet을 반환합니다.

```text
ResultSet 1 → 당월
ResultSet 2 → 전년동월
ResultSet 3 → 전월
```

Java에서는 다음 JDBC 흐름으로 모든 ResultSet을 순차 처리합니다.

```text
CallableStatement.execute()
        ↓
getResultSet()
        ↓
getMoreResults()
```

배치 1회 실행 시 각 ResultSet을 별도의 Excel 파일로 생성합니다.

예:

```text
GALA_일자별판매실적분석_3PP_당월_261020.xlsx
GALA_일자별판매실적분석_3PP_전년동월_261020.xlsx
GALA_일자별판매실적분석_3PP_전월_261020.xlsx
```

예상 ResultSet 수와 실제 반환 수가 일치하지 않는 경우 배치를 실패 처리하도록 구성했습니다.

---

### 외부 설정 관리

DB 접속정보와 실행 설정은 외부 `application.properties`에서 관리합니다.

```text
DB URL
DB 계정
DB 비밀번호
Stored Procedure
Excel 출력 경로
파일명 패턴
```

실제 설정 파일은 Git에서 제외하여 운영 DB 계정과 비밀번호가 저장소에 포함되지 않도록 했습니다.

예제 설정:

```properties
db.stored-procedure=galadev_SWSLSalesListAnalysisDailyAuto_Daily_3PP

excel.output-file-name-pattern=GALA_일자별판매실적분석_3PP_{period}_{timestamp}.xlsx
```

Java 코드에는 `3PP` 또는 특정 리포트명을 하드코딩하지 않고 파일명 패턴을 설정으로 관리합니다.

따라서 향후 다른 Stored Procedure를 사용할 경우에도 Java 소스 수정 없이 설정 변경만으로 프로시저와 파일명 규칙을 변경할 수 있습니다.

파일명 치환값:

```text
{period}
 → 당월
 → 전년동월
 → 전월

{timestamp}
 → yyMMdd
```

`application.properties`는 UTF-8로 읽도록 처리하여 한글 파일명 및 설정값 깨짐 문제도 해결했습니다.

---

### ERP Excel 형식 재현

Apache POI를 사용해 실제 ERP에서 다운로드한 `일자별판매실적분석` Excel 형식을 재현했습니다.

적용 내용:

- 제목 행 병합 및 서식
- 녹색 헤더 / 흰색 글씨
- TOTAL 행 연녹색 배경
- ERP 기준 컬럼 너비 및 행 높이
- 셀 테두리
- 날짜 `yyyy-MM-dd` 표시
- 숫자 Excel Numeric Cell 처리
- 천 단위 콤마
- 금액 컬럼 소수점 2자리 표시
- ResultSetMetaData 기반 동적 컬럼 생성

컬럼명은 Stored Procedure의 한글 Alias를 `ResultSetMetaData`로 읽어 동적으로 생성합니다.

---

### Apache POI CellStyle 재사용

대용량 ResultSet 처리 중 다음 오류가 발생했습니다.

```text
The maximum number of Cell Styles was exceeded.
You can define up to 64000 style in a .xlsx Workbook
```

기존에는 셀마다 새로운 `CellStyle`을 생성하여 데이터 건수가 많아질수록 Workbook의 Style 개수가 증가했습니다.

이를 개선하여 Workbook 생성 시 필요한 스타일을 한 번만 생성하고 이후 모든 셀에서 재사용하도록 변경했습니다.

재사용하는 주요 스타일:

```text
제목
헤더

본문
- 문자
- 날짜
- 정수
- 일반 소수
- 금액 소수점 2자리

TOTAL
- 문자
- 날짜
- 정수
- 일반 소수
- 금액 소수점 2자리
```

셀 생성 반복문에서는 새로운 `CellStyle`을 생성하지 않고 미리 생성된 스타일만 적용합니다.

---

## 테스트

JUnit 5 기반으로 주요 기능을 테스트합니다.

현재 테스트 항목:

- 설정 파일 로딩
- 파일명 패턴 치환
- Excel 생성
- ERP 스타일 적용
- 날짜 / 숫자 / 금액 타입
- TOTAL 행 스타일
- 다중 ResultSet 처리
- 대용량 Excel 생성

대용량 테스트에서는 약 3,000행의 데이터를 생성하여 Excel 생성이 정상적으로 완료되는지 검증합니다.

또한 Workbook의 CellStyle 개수가 행 수에 비례해 증가하지 않도록 Style 개수를 검증합니다.

---

## 프로젝트 구조

```text
erp-excel-batch/
├─ pom.xml
├─ README.md
├─ .gitignore
├─ src/
│  ├─ main/
│  │  ├─ java/com/company/erpbatch/
│  │  │  ├─ Main.java
│  │  │  ├─ batch/
│  │  │  ├─ config/
│  │  │  ├─ db/
│  │  │  ├─ excel/
│  │  │  └─ exception/
│  │  └─ resources/
│  │     ├─ application.properties.example
│  │     └─ logback.xml
│  └─ test/
└─ logs/
```

---

## 빌드 및 실행

테스트:

```powershell
mvn clean test
```

실행 가능한 JAR 생성:

```powershell
mvn clean package
```

빌드 결과:

```text
target/erp-excel-batch.jar
```

실행:

```powershell
java -jar .\target\erp-excel-batch.jar
```

애플리케이션은 성공 시 종료 코드 `0`, 실패 시 `1`을 반환합니다.

---

## 배포 구조

스케줄러 서버의 Java 설치 환경에 영향을 받지 않도록 전용 Java 17 Runtime을 배포 폴더에 함께 구성했습니다.

```text
erp-excel-batch-deploy/
├─ runtime/
│  ├─ bin/
│  │  └─ java.exe
│  ├─ conf/
│  └─ lib/
├─ erp-excel-batch.jar
├─ application.properties
├─ run.bat
└─ logs/
```

`run.bat`은 시스템에 설치된 Java가 아니라 배포 폴더 내부 Runtime을 직접 사용합니다.

```bat
@echo off
setlocal

cd /d "%~dp0"

"%~dp0runtime\bin\java.exe" ^
  -jar "%~dp0erp-excel-batch.jar" ^
  "%~dp0application.properties"

set EXIT_CODE=%ERRORLEVEL%

exit /b %EXIT_CODE%
```

이를 통해 서버의 `JAVA_HOME`, `PATH`, 설치된 Java 버전과 관계없이 동일한 실행 환경을 유지할 수 있습니다.

---

## 트러블슈팅

### Fat JAR 서명 충돌

Maven Shade Plugin으로 실행 가능한 JAR 생성 후 다음 오류가 발생했습니다.

```text
Invalid signature file digest for Manifest main attributes
```

의존성 JAR의 전자서명 파일이 Shade 과정에서 충돌한 것이 원인이었습니다.

다음 파일을 Fat JAR에서 제외하여 해결했습니다.

```text
META-INF/*.SF
META-INF/*.DSA
META-INF/*.RSA
```

---

### 한글 설정값 및 파일명 깨짐

`application.properties`에서 한글 파일명을 설정했을 때 문자열이 깨지는 문제가 발생했습니다.

기본 `Properties.load(InputStream)` 방식 대신 UTF-8 `Reader`를 사용하도록 수정하여 해결했습니다.

```text
application.properties
        ↓ UTF-8
BatchConfig
        ↓
한글 설정값 정상 처리
```

---

### Excel CellStyle 64,000개 제한

대용량 데이터를 Excel로 생성하면서 Apache POI의 CellStyle 최대 개수를 초과하는 문제가 발생했습니다.

```text
The maximum number of Cell Styles was exceeded.
```

셀마다 새로운 Style을 생성하던 구조를 Workbook 단위 Style 재사용 방식으로 변경했습니다.

이후 대용량 테스트를 추가하여 동일 문제가 재발하지 않도록 검증하고 있습니다.

---

### 외부 설정 파일 관리

실제 DB 접속정보는 JAR 내부가 아닌 외부 `application.properties`에서 읽도록 구성했습니다.

이를 통해 환경별 DB 정보, Stored Procedure, 출력 경로, 파일명 패턴을 소스 수정 없이 변경할 수 있습니다.

---

## 현재 진행 상태

### 완료

- [x] Java 17 / Maven 배치 프로젝트 구성
- [x] MSSQL JDBC 연결
- [x] 실제 Stored Procedure 실행
- [x] CallableStatement 기반 다중 ResultSet 처리
- [x] 당월 / 전년동월 / 전월 Excel 3개 생성
- [x] ResultSet 개수 검증 및 오류 처리
- [x] ResultSet 기반 Excel 생성
- [x] ERP 원본 Excel 형식 재현
- [x] 숫자 / 날짜 타입 및 표시 형식 적용
- [x] TOTAL 행 출력 및 스타일 적용
- [x] 프로시저별 Excel 파일명 패턴 외부 설정
- [x] 한글 설정값 및 파일명 UTF-8 처리
- [x] Apache POI CellStyle 재사용 구조 적용
- [x] 3,000행 대용량 Excel 테스트
- [x] 실제 DB 연동 테스트
- [x] JUnit 테스트 및 Maven 빌드 성공
- [x] 실행 가능한 Fat JAR 생성
- [x] 전용 Java Runtime 포함 배포 패키지 구성
- [x] `run.bat` 기반 독립 실행 테스트

### 예정

- [ ] 실제 Windows 공유 폴더 적용
- [ ] 스케줄러 서버 배포
- [ ] Windows 작업 스케줄러 월~금 실행 등록
- [ ] 작업 스케줄러 계정 공유 폴더 권한 검증
- [ ] 운영 환경 자동 실행 검증