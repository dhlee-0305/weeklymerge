# WeeklyMerge

주간 보고서 `.docx` 파일 여러 개를 선택해서 하나의 통합 문서로 병합하는 Java Swing 프로그램입니다.

이 프로그램은 팀별 주간 보고서를 불러와 공통 템플릿 기준으로 다음 영역을 취합합니다.

- 프로젝트 진행 현황
- 인력 운용 현황
- 거래처 영업/동향 정보
- 주요 업무 사항

사용자는 UI에서 병합할 파일을 선택하고, 결과 문서를 지정된 출력 폴더에 생성할 수 있습니다.

## 기술 스택

- Java 17
- Maven
- Java Swing
- Apache POI `poi-ooxml 5.4.1`

## 개발 환경 셋팅

### 1. 필수 프로그램 설치

아래 도구가 설치되어 있어야 합니다.

- JDK 17
- Maven 3.9 이상 권장

버전 확인 예시:

```powershell
java -version
mvn -version
```

### 2. 프로젝트 준비

프로젝트 루트에서 Maven이 동작하는지 확인합니다.

```powershell
mvn clean compile
```

### 3. 환경 설정 파일 준비

프로젝트 루트에 `app.properties` 파일을 생성하고 아래처럼 입력합니다.

```properties
report.doc.path=C:\Downloads\weekly_report\
report.doc.out.path=C:\Downloads\weekly_report\output\
```

설명:

- `report.doc.path`: 병합할 원본 보고서 `.docx` 파일들이 있는 폴더
- `report.doc.out.path`: 통합 결과 파일과 템플릿 파일이 있는 폴더

### 4. 템플릿 파일 준비

출력 폴더 안에 아래 템플릿 파일이 있어야 합니다.

```text
{report.doc.out.path}\경영전략회의_template.docx
```

생성 결과 파일 이름은 아래 형식입니다.

```text
YYYYMMDD_경영전략회의.docx
```

## 프로그램 실행 방법

프로젝트 루트에서 아래 명령으로 실행합니다.

```powershell
mvn clean compile exec:java
```

실행 후 동작 방식:

1. `report.doc.path` 폴더의 `.docx` 파일 목록을 화면에 표시합니다.
2. 병합할 파일을 하나 이상 선택합니다.
3. `Merge Reports` 버튼을 클릭합니다.
4. 통합 문서가 `report.doc.out.path` 폴더에 생성됩니다.

## 프로젝트 구조

```text
weeklymerge
├─ src
│  ├─ WeeklyMergeApp.java
│  └─ ReportMerger.java
├─ pom.xml
├─ app.properties
└─ README.md
```

## 참고 사항

- `app.properties`는 로컬 환경 설정 파일이므로 Git에 커밋하지 않는 것을 권장합니다.
- 원본 문서와 템플릿 문서의 표 구조가 크게 다르면 병합 결과가 예상과 다를 수 있습니다.
- `주요 업무 사항`은 각 문서의 해당 섹션을 읽어와 통합 문서 마지막에 추가합니다.
