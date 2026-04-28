## Project Overview
주간 보고서 `.docx` 파일 여러 개를 선택해서 하나의 통합 문서로 병합하는 Java Swing 프로그램입니다.

## Coding Style
- 에러 발생 시 에러 발생 파일과 위치 그리고 해당 데이터를 사용자에게 제공한다.
- 커밋 메시지는 Conventional Commits 규격 준수한다.
- 커밋 메시지 작성 시 커밋 타입을 추가한다. 커밋 타입을 제외한 내용은 모두 한글로 작성한다.
---
example)
 feat: 원본 서식 유지 기능 추가
 - 원본 문서의 번호 목록 정의를 결과 문서로 복사
 - 문단/표 내부 번호 참조를 새 번호 ID로 재매핑
---
- 주석 내 코드와 관련된 부분을 제외하고 모두 한글로 주석을 작성한다.
- 클래스, 메서드는 Javadoc을 지원하는 형태로 주석을 작성한다.
- 조건문에는 필요한 경우 의도를 설명하는 주석을 남긴다.
- 예외 메시지는 사용자 관점에서 이해 가능하게 작성한다.
- 코드의 예상 동작이나 출력을 설명하는 주석을 작성한다.
```
example)
/**
 * 두 스타일의 XML 정의가 같은지 비교한다.
 *
 * @param sourceStyle 원본 문서 스타일
 * @param targetStyle 결과 문서 스타일
 * @return 스타일 정의가 같으면 true
 */
private static boolean hasSameStyleDefinition(XWPFStyle sourceStyle, XWPFStyle targetStyle){...}
```
- 기존 코드를 수정하는 경우에는 파일 전체를 새로 작성하지 않는다.
- 꼭 파일 전체를 새로 작성해야 하는 경우 필요한 이유를 설명하고, 질문을 통해 확인 후 진행한다. 

## Boundaries
- Always: 커밋 전 빌드 테스트 실행, 코딩 스타일 규칙 준수
- Ask first: 데이터베이스 스키마 변경, 종속성 추가, 파일 전체 재작성
- Never: secrets 커밋, node_modules/ 수정, CI 설정 수정