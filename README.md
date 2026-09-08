# ERP Excel Batch

영림원 ERP에서 반복적으로 수행하는 Excel 다운로드 업무를 자동화하기 위해 개발하는 **Java 기반 사내 배치 프로그램**입니다.

기존 ERP 애플리케이션의 소스 코드는 수정하지 않고, MSSQL Stored Procedure를 통해 필요한 데이터를 조회한 뒤 Java 배치 프로그램에서 Excel 파일을 생성하는 구조로 설계했습니다.

현재 단계에서는 특정 ERP 업무에 종속되지 않는 **공통 배치 프레임워크**를 구축한 상태입니다.

---

## 1. 프로젝트 배경

기존 업무에서는 ERP 화면에서 조회 조건을 입력하고 Excel 파일을 반복적으로 다운로드해야 하는 작업이 존재합니다.

ERP 자체 소스 코드를 수정할 수 없는 환경이기 때문에 다음과 같은 방식으로 자동화를 설계했습니다.

```text
영림원 ERP 화면 분석
        ↓
MSSQL Stored Procedure 작성
        ↓
Java Batch에서 Stored Procedure 호출
        ↓
ResultSet 조회
        ↓
Apache POI로 Excel 생성
        ↓
Windows 공유 폴더 저장
        ↓
Windows 작업 스케줄러를 통한 자동 실행
```

ERP 시스템과 자동화 프로그램을 분리하여 기존 ERP에 영향을 주지 않도록 하는 것을 기본 원칙으로 했습니다.

---

## 2. 기술 스택

* Java 17
* Maven
* Microsoft SQL Server
* JDBC
* Apache POI
* SLF4J
* Logback
* JUnit 5
* Maven Shade Plugin
* Windows

Spring Boot나 별도의 WAS는 사용하지 않고, 실행 후 작업을 완료하면 종료되는 경량 Command-line Batch 형태로 구현했습니다.

---

## 3. 프로젝트 구조

```text
erp-excel-batch/
├─ pom.xml
├─ README.md
├─ .gitignore
│
├─ src/
│  ├─ main/
│  │  ├─ java/
│  │  │  └─ com/company/erpbatch/
│  │  │     ├─ Main.java
│  │  │     │
│  │  │     ├─ batch/
│  │  │     │  └─ BatchJob.java
│  │  │     │
│  │  │     ├─ config/
│  │  │     │  └─ BatchConfig.java
│  │  │     │
│  │  │     ├─ db/
│  │  │     │  ├─ DatabaseConnection.java
│  │  │     │  └─ StoredProcedureExecutor.java
│  │  │     │
│  │  │     ├─ excel/
│  │  │     │  ├─ ExcelService.java
│  │  │     │  └─ ExcelColumn.java
│  │  │     │
│  │  │     └─ exception/
│  │  │        └─ BatchException.java
│  │  │
│  │  └─ resources/
│  │     ├─ application.properties.example
│  │     └─ logback.xml
│  │
│  └─ test/
│     └─ java/
│
└─ logs/
```

---

## 4. 애플리케이션 동작 구조

전체 실행 흐름은 다음과 같습니다.

```text
Main
  ↓
BatchConfig
  ↓
BatchJob
  ↓
DatabaseConnection
  ↓
StoredProcedureExecutor
  ↓
ResultSet
  ↓
ExcelService
  ↓
.xlsx 파일 생성
```

### Main

애플리케이션의 진입점입니다.

설정 파일을 로드하고 BatchJob을 실행하며 작업 결과에 따라 프로세스 종료 코드를 반환합니다.

```text
0 : 정상 종료
1 : 오류 종료
```

### BatchConfig

외부 `application.properties` 파일에서 다음과 같은 실행 정보를 읽습니다.

* MSSQL JDBC URL
* DB 사용자
* DB 비밀번호
* Stored Procedure 이름
* Excel 출력 경로
* Excel 파일명 패턴

민감한 운영 정보가 코드 또는 Git 저장소에 포함되지 않도록 외부 설정 방식으로 구성했습니다.

### StoredProcedureExecutor

JDBC `CallableStatement`를 이용해 MSSQL Stored Procedure를 호출합니다.

실제 ERP 테이블이나 Stored Procedure 이름은 프레임워크에 하드코딩하지 않고, 이후 실제 업무 화면을 분석한 뒤 연결할 예정입니다.

### ExcelService

Apache POI의 `XSSFWorkbook`을 사용하여 조회 결과를 `.xlsx` 파일로 생성합니다.

현재 공통 프레임워크에서는 `ResultSetMetaData`를 이용해 컬럼명을 자동으로 Excel 헤더로 만들고 데이터를 순차적으로 기록합니다.

실제 ERP 업무 적용 단계에서는 필요한 경우 컬럼 순서, 숫자 및 날짜 형식, 셀 스타일 등을 전용 로직으로 확장할 예정입니다.

---

## 5. 설정 파일

예제 설정 파일을 복사합니다.

```powershell
Copy-Item src\main\resources\application.properties.example application.properties
```

`application.properties`에 실제 실행 환경 값을 입력합니다.

실제 설정 파일에는 DB 계정 및 비밀번호가 포함될 수 있기 때문에 Git에서 제외되어 있습니다.

```gitignore
application.properties
```

설정 파일을 별도의 위치에서 관리할 수도 있습니다.

```powershell
java -jar target\erp-excel-batch.jar C:\batch-config\application.properties
```

DB 비밀번호나 실제 운영 접속 정보는 소스 코드와 로그에 출력하지 않습니다.

---

## 6. 빌드

### 테스트

```powershell
mvn clean test
```

### 실행 가능한 JAR 생성

```powershell
mvn clean package
```

빌드가 완료되면 다음 실행 파일이 생성됩니다.

```text
target/erp-excel-batch.jar
```

---

## 7. 실행

```powershell
java -jar target\erp-excel-batch.jar
```

기본 실행 위치에 `application.properties`가 없는 경우 다음과 같이 오류를 출력하고 종료합니다.

```text
설정 파일을 찾을 수 없습니다.
```

이는 애플리케이션의 Main 진입 및 예외 처리 로직이 정상적으로 동작하고 있음을 의미합니다.

실제 DB 연결 테스트는 업무용 Stored Procedure와 실행 환경 설정이 준비된 이후 진행합니다.

---

## 8. 로그

SLF4J와 Logback을 사용합니다.

주요 실행 단계에 대해 다음 정보를 기록하도록 구성했습니다.

* 프로그램 시작
* DB 연결
* Stored Procedure 실행
* 조회 건수
* Excel 생성 시작 및 완료
* 결과 파일 경로
* 정상 종료
* 오류 및 Stack Trace

로그 파일은 다음 경로에 생성됩니다.

```text
logs/erp-excel-batch.log
```

DB 비밀번호 등의 민감 정보는 로그에 기록하지 않습니다.

---

## 9. 테스트

JUnit 5 기반 단위 테스트를 구성했습니다.

현재 프레임워크 단계에서는 실제 MSSQL 서버에 의존하지 않는 테스트만 작성했습니다.

주요 테스트 대상:

* 외부 설정 파일 로딩
* Excel 파일명 `{timestamp}` 처리

실제 업무 로직이 추가되면 Stored Procedure 입력값 처리와 Excel 데이터 변환 로직에 대한 테스트를 확장할 예정입니다.

---

## 10. 트러블슈팅

### Maven 실행 환경 구성

초기 개발 환경에는 Maven이 설치되어 있지 않아 프로젝트 빌드를 실행할 수 없었습니다.

Apache Maven을 설치하고 Windows 환경 변수에 다음과 같이 등록했습니다.

```text
MAVEN_HOME=C:\Tools\apache-maven-3.9.16
```

`Path`:

```text
%MAVEN_HOME%\bin
```

환경 구성 확인:

```powershell
mvn -version
```

검증 환경:

```text
Apache Maven 3.9.16
Java 17
Windows 11
```

이후 다음 명령으로 단위 테스트가 정상 통과하는 것을 확인했습니다.

```powershell
mvn clean test
```

---

### 실행 가능한 Fat JAR의 전자서명 충돌

Maven Shade Plugin을 이용해 모든 의존성을 포함하는 실행 가능한 JAR을 만든 후 다음 오류가 발생했습니다.

```text
java.lang.SecurityException:
Invalid signature file digest for Manifest main attributes
```

일부 의존성 JAR의 전자서명 파일이 Shade 과정에서 함께 병합되면서 기존 서명 검증 정보가 유효하지 않게 된 것이 원인이었습니다.

`maven-shade-plugin`에서 다음 파일을 Fat JAR에 포함하지 않도록 설정하여 해결했습니다.

```xml
<filters>
    <filter>
        <artifact>*:*</artifact>
        <excludes>
            <exclude>META-INF/*.SF</exclude>
            <exclude>META-INF/*.DSA</exclude>
            <exclude>META-INF/*.RSA</exclude>
        </excludes>
    </filter>
</filters>
```

수정 후:

```powershell
mvn clean package
```

를 다시 실행하여 기존 `target` 빌드 산출물을 제거하고 JAR을 재생성했습니다.

이후:

```powershell
java -jar target\erp-excel-batch.jar
```

실행 시 정상적으로 Main 클래스에 진입하고 외부 설정 파일을 로드하는 단계까지 동작하는 것을 확인했습니다.

---

## 11. 보안 및 Git 관리

다음 파일과 디렉터리는 Git에서 제외합니다.

```gitignore
application.properties
logs/
target/
.idea/
*.iml
```

이를 통해 다음 정보가 저장소에 포함되는 것을 방지합니다.

* DB 비밀번호
* 운영 DB 접속 정보
* 로컬 IDE 설정
* 빌드 산출물
* 실행 로그

---

## 12. 향후 개발 계획

실제 첫 번째 ERP 자동화 업무는 다음 순서로 진행합니다.

1. 영림원 ERP 대상 화면 분석
2. 화면 조회 조건 및 데이터 구조 확인
3. 동일한 결과를 반환하는 MSSQL Stored Procedure 작성
4. ERP 화면 조회 결과와 Stored Procedure 결과 비교
5. Excel 컬럼 및 표시 형식 정의
6. Java Batch에 실제 Stored Procedure 연결
7. 조회조건 파라미터 처리 구현
8. Excel 생성 결과 검증
9. Windows 공유 폴더 저장 테스트
10. Windows 작업 스케줄러 등록
11. 운영 로그 및 실패 처리 검증

---

## 13. 현재 진행 상태

### 완료

* [x] Java 17 Maven 프로젝트 구성
* [x] MSSQL JDBC 의존성 구성
* [x] Stored Procedure 실행 공통 구조 구현
* [x] Apache POI Excel 생성 공통 구조 구현
* [x] 외부 설정 파일 구성
* [x] 실제 설정 파일 Git 제외
* [x] SLF4J + Logback 로그 구성
* [x] JUnit 5 테스트 구성
* [x] Maven 테스트 성공
* [x] 실행 가능한 Fat JAR 생성
* [x] Maven Shade Plugin 서명 충돌 해결
* [x] JAR Main 클래스 실행 검증

### 예정

* [ ] 실제 ERP 화면 분석
* [ ] 업무용 Stored Procedure 작성
* [ ] 실제 MSSQL 연결
* [ ] Stored Procedure 파라미터 처리
* [ ] ERP 전용 Excel 양식 구현
* [ ] 공유 폴더 저장 검증
* [ ] Windows 작업 스케줄러 운영 적용
