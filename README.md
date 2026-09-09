# ERP Excel Batch

영림원 ERP에서 반복적으로 수행하는 Excel 다운로드 업무를 자동화하기 위한 **Java 기반 사내 배치 프로그램**입니다.

기존 ERP 소스는 수정하지 않고, MSSQL Stored Procedure로 데이터를 조회한 뒤 Java에서 Excel 파일을 생성하도록 구성했습니다.

## 프로젝트 배경

기존 업무에서는 ERP 화면에서 조회 후 Excel 파일을 반복적으로 다운로드해야 했습니다.

ERP 소스를 직접 수정할 수 없는 환경이므로 다음과 같이 자동화 구조를 설계했습니다.

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
```

Java 애플리케이션은 데이터 조회 및 Excel 생성 후 종료되는 경량 Command-line Batch 방식입니다.

---

## 기술 스택

* Java 17
* Maven
* Microsoft SQL Server
* JDBC
* Apache POI
* SLF4J + Logback
* JUnit 5
* Maven Shade Plugin
* Windows

---

## 주요 구현

### MSSQL Stored Procedure 연동

JDBC `CallableStatement`를 이용해 실제 MSSQL Stored Procedure를 실행합니다.

업무 조회 조건과 데이터 가공, TOTAL 계산 등은 Stored Procedure에서 처리하고 Java는 결과를 받아 Excel로 출력하는 역할에 집중하도록 구성했습니다.

```text
Main
 → BatchConfig
 → BatchJob
 → DatabaseConnection
 → StoredProcedureExecutor
 → ResultSet
 → ExcelService
```

### 외부 설정 관리

DB 접속정보와 실행 설정은 `application.properties`에서 관리합니다.

```text
DB URL
DB 계정
DB 비밀번호
Stored Procedure
Excel 출력 경로
파일명 패턴
```

실제 설정 파일은 Git에서 제외하여 운영 DB 계정과 비밀번호가 저장소에 포함되지 않도록 했습니다.

### ERP Excel 형식 재현

Apache POI를 사용해 실제 ERP에서 다운로드한 `일자별판매실적분석` Excel 형식을 재현했습니다.

적용 내용:

* 제목 행 병합 및 서식
* 녹색 헤더 / 흰색 글씨
* TOTAL 행 연녹색 배경
* ERP 기준 컬럼 너비 및 행 높이
* 셀 테두리
* 날짜 `yyyy-MM-dd` 표시
* 숫자 Excel Numeric Cell 처리
* 천 단위 콤마
* 금액 컬럼 소수점 2자리 표시
* `application.properties`를 UTF-8로 읽도록 처리하여 한글 파일명 및 설정값 깨짐 문제를 해결했습니다.

컬럼명은 Stored Procedure의 한글 Alias를 `ResultSetMetaData`로 읽어 동적으로 생성합니다.


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

실행:

```powershell
java -jar target\erp-excel-batch.jar
```

빌드 결과:

```text
target/erp-excel-batch.jar
```

애플리케이션은 성공 시 종료 코드 `0`, 실패 시 `1`을 반환합니다.

---

## 트러블슈팅

### Fat JAR 서명 충돌

Maven Shade Plugin으로 실행 가능한 JAR 생성 후 다음 오류가 발생했습니다.

```text
Invalid signature file digest for Manifest main attributes
```

의존성 JAR의 전자서명 파일이 Shade 과정에서 충돌한 것이 원인이었습니다.

다음 파일을 Fat JAR에서 제외하도록 설정하여 해결했습니다.

```text
META-INF/*.SF
META-INF/*.DSA
META-INF/*.RSA
```

### 외부 설정 파일 관리

실제 DB 접속정보는 JAR 내부가 아닌 외부 `application.properties`에서 읽도록 구성했습니다.

이를 통해 환경별 DB 정보와 출력 경로를 소스 수정 없이 변경할 수 있습니다.

---

## 현재 진행 상태

### 완료

* [x] Java 17 / Maven 배치 프로젝트 구성
* [x] MSSQL JDBC 연결
* [x] 실제 Stored Procedure 실행
* [x] ResultSet 기반 Excel 생성
* [x] ERP 원본 Excel 형식 재현
* [x] 숫자 / 날짜 타입 및 표시 형식 적용
* [x] TOTAL 행 출력 및 스타일 적용
* [x] 실제 DB 연동 테스트
* [x] JUnit 테스트 및 Maven 빌드 성공
* [x] 실행 가능한 Fat JAR 생성

### 예정

* [ ] 실제 Windows 공유 폴더 적용
* [ ] Windows 작업 스케줄러 등록
* [ ] 운영 환경 자동 실행 검증
