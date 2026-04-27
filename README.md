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

프로그램은 실행 폴더의 `app.properties` 파일에서 입출력 경로를 읽습니다.
개발 중에는 프로젝트 루트에 `app.properties` 파일을 두고, 배포 후에는 `weeklymerge.jar`와 같은 폴더에 둡니다.

```properties
report.doc.path=C:\Downloads\weekly_report\
report.doc.out.path=C:\Downloads\weekly_report\output\
```

설명:

- `report.doc.path`: 병합할 원본 보고서 `.docx` 파일들이 있는 폴더
- `report.doc.out.path`: 통합 결과 파일과 템플릿 파일이 있는 폴더

`src/main/resources/app.properties`는 빌드 산출물에 포함되는 기본 예시 파일입니다.
실제 실행 환경의 경로는 JAR 옆에 별도 `app.properties`를 두어 관리하는 방식을 권장합니다.

### 4. 메시지 파일 준비

화면 문구와 오류 메시지는 `src/main/resources/messages.properties`에 정의합니다.
이 파일은 Maven 빌드 시 JAR 내부에 포함되므로, 기본 배포에는 별도 `messages.properties`가 필요하지 않습니다.

배포 후 문구만 바꾸고 싶다면 `weeklymerge.jar`와 같은 폴더에 `messages.properties`를 따로 둘 수 있습니다.
실행 폴더의 파일이 JAR 내부 기본 메시지보다 우선 적용됩니다.

### 5. 템플릿 파일 준비

출력 폴더 안에 아래 템플릿 파일이 있어야 합니다.

```text
{report.doc.out.path}\경영전략회의_template_java.docx
```

생성 결과 파일 이름은 아래 형식입니다.

```text
YYYYMMDD_IT서비스부문_주간보고.docx
```



## 프로그램 실행 방법

### 1. 개발 중 실행

프로젝트 루트에서 아래 명령으로 바로 실행합니다.

```powershell
mvn clean compile exec:java
```

### 2. 배포용 실행 파일 생성

아래 명령으로 모든 의존 라이브러리가 포함된 단일 JAR 파일을 생성합니다.

```powershell
mvn clean package
```

빌드가 완료되면 `target/weeklymerge.jar` 파일이 생성됩니다.

`src/main/resources` 아래의 리소스 파일은 JAR 내부에 함께 포함됩니다.
단, 실행 환경별 문서 경로를 지정해야 하므로 생성된 JAR 파일과 `app.properties`를 같은 폴더에 두고 아래 명령으로 실행합니다.

```powershell
java -jar weeklymerge.jar
```

JAR에 리소스가 포함되었는지 확인하려면 아래 명령을 사용할 수 있습니다.

```powershell
jar tf target\weeklymerge.jar | findstr "messages.properties app.properties"
```

### 3. 실행 후 동작 방식

1. `report.doc.path` 폴더의 `.docx` 파일 목록을 화면에 표시합니다.
2. 병합할 파일을 하나 이상 선택합니다. 선택한 순서대로 파일이 병합됩니다.
3. `Merge Reports` 버튼을 클릭합니다.
4. 통합 문서가 `report.doc.out.path` 폴더에 생성됩니다.



## 프로젝트 구조

```text
weeklymerge
├─ src
│  ├─ AppConfig.java
│  ├─ Messages.java
│  ├─ WeeklyMergeApp.java
│  ├─ ReportMerger.java
│  └─ main
│     └─ resources
│        ├─ app.properties
│        └─ messages.properties
├─ pom.xml
├─ app.properties        # 개발/배포 실행 폴더용 로컬 설정 파일
└─ README.md
```



## 참고 사항

- 실행 폴더의 `app.properties`는 로컬 환경 설정 파일이므로 Git에 커밋하지 않는 것을 권장합니다.
- `messages.properties`는 JAR 내부 기본 메시지로 포함되며, 실행 폴더에 같은 이름의 파일을 두면 문구를 덮어쓸 수 있습니다.
- 원본 문서와 템플릿 문서의 표 구조가 다르면 병합 결과가 예상과 다를 수 있습니다.
