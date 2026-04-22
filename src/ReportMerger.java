import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Path;
import java.text.DecimalFormat;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.temporal.TemporalAdjusters;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Pattern;

import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTrPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;

/**
 * 여러 개의 주간 보고 Word 문서를 하나의 결과 문서로 병합하는 유틸리티 클래스다.
 * 템플릿 문서를 기준으로 섹션별 데이터를 합치고, 필요한 서식은 최대한 유지한다.
 */
/**
 * 여러 개의 주간 보고 Word 문서를 하나의 결과 문서로 병합하는 유틸리티 클래스다.
 * 템플릿 문서를 기준으로 병합을 수행하며, 가능한 범위에서 섹션 구조와 서식을 유지한다.
 */
public final class ReportMerger {
    // 템플릿 문서에서 실제 보고 일자로 치환할 플레이스홀더다.
    private static final String REPORT_DATE_PLACEHOLDER = "[보고일]";
    public static final String TEMPLATE_FILE_NAME = "경영전략회의_template_java.docx";
    private static final int STAFF_CATEGORY_COLUMN_INDEX = 0;
    private static final int SALES_EXCLUDED_COLUMN_INDEX = 0;

    private static final String OUTPUT_FILE_SUFFIX = "_IT서비스부문_주간보고.docx";
    private static final SectionSpec PROJECT_STATUS_SECTION = new SectionSpec(List.of("프로젝트 진행 현황", "프로젝트 현황"), 0, 1);
    private static final SectionSpec STAFF_STATUS_SECTION = new SectionSpec(List.of("인력 운용 현황", "인력 현황"), 1, 1);
    private static final SectionSpec SALES_STATUS_SECTION = new SectionSpec(
            List.of("거래처 영업/동향 정보",
                    "거래처 영업 동향 정보",
                    "영업/동향 정보"),
            2,
            1);
    private static final List<String> SALES_FIXED_HEADER_ROW = List.of("구분", "고객사/부서", "주요 정보", "비고");
    private static final List<String> SALES_FIXED_HEADER_ROW_WITHOUT_FIRST_COLUMN = List.of("고객사/부서", "주요 정보", "비고");
    private static final SectionSpec ISSUES_SECTION = new SectionSpec(
            List.of("주요 업무 사항",
                    "주요업무사항",
                    "주요 업무",
                    "주요업무",
                    "주요 이슈",
                    "이슈 사항"),
            3,
            1);
    private static final Pattern NUMERIC_PATTERN = Pattern.compile("[-+]?\\d+(\\.\\d+)?");
    private static final DecimalFormat NUMBER_FORMAT = new DecimalFormat("0.################");
    private static final ProgressListener NO_OP_PROGRESS_LISTENER = (percent, message) -> {
    };

    private ReportMerger() {
    }

    /**
     * 병합 진행 상태를 외부 UI에 전달하기 위한 콜백 인터페이스다.
     * percent는 0~100 범위의 진행률, message는 현재 단계 설명으로 사용된다.
     */
    /**
     * 병합 진행 상태를 UI에 전달하기 위한 콜백 인터페이스다.
     * {@code percent}는 0~100 범위의 진행률이고, {@code message}는 현재 단계 설명이다.
     */
    @FunctionalInterface
    public interface ProgressListener {
        /**
         * 최신 병합 진행 상태를 전달받는다.
         *
         * @param percent 병합 진행 퍼센트
         * @param message 현재 병합 단계 설명
         */
        void onProgress(int percent, String message);
    }

    /**
     * 여러 개의 원본 보고서를 템플릿 문서에 병합해 최종 주간보고 파일을 생성한다.
     *
     * @param sourceFiles 병합 대상 원본 문서 목록
     * @param outputDirectory 템플릿 파일이 위치하고 결과 파일을 저장할 디렉터리
     * @return 생성된 결과 문서 경로
     * @throws IOException 템플릿이 없거나 문서 입출력에 실패한 경우
     */
    /**
     * 외부 진행률 보고 없이 원본 보고서를 병합해 새 결과 문서를 만든다.
     *
     * @param sourceFiles 병합할 원본 보고서 문서 목록
     * @param outputDirectory 템플릿이 있고 결과 파일이 저장될 폴더
     * @return 생성된 결과 문서 경로
     * @throws IOException 템플릿 또는 원본 문서를 처리할 수 없는 경우
     */
    public static Path mergeReports(List<Path> sourceFiles, Path outputDirectory) throws IOException {
        return mergeReports(sourceFiles, outputDirectory, NO_OP_PROGRESS_LISTENER);
    }

    /**
     * 병합 진행률을 외부에 보고하면서 결과 문서를 생성한다.
     *
     * @param sourceFiles 병합할 원본 보고서 목록
     * @param outputDirectory 템플릿과 결과 파일이 위치할 출력 폴더
     * @param progressListener 진행 상태를 전달받을 콜백, 없으면 무시된다
     * @return 생성된 결과 문서 경로
     * @throws IOException 템플릿 또는 원본 문서를 읽거나 쓰는 중 오류가 발생한 경우
     */
    /**
     * 진행률을 보고하면서 원본 보고서를 병합해 새 결과 문서를 만든다.
     *
     * @param sourceFiles 병합할 원본 보고서 문서 목록
     * @param outputDirectory 템플릿이 있고 결과 파일이 저장될 폴더
     * @param progressListener 진행률 콜백, {@code null}이면 무시한다
     * @return 생성된 결과 문서 경로
     * @throws IOException 템플릿 또는 원본 문서를 처리할 수 없는 경우
     */
    public static Path mergeReports(
            List<Path> sourceFiles,
            Path outputDirectory,
            ProgressListener progressListener) throws IOException {
        Path templatePath = outputDirectory.resolve(TEMPLATE_FILE_NAME);
        ProgressListener listener = progressListener == null ? NO_OP_PROGRESS_LISTENER : progressListener;
        int totalSteps = sourceFiles.size() + 7;
        int completedSteps = 0;

        // 병합 시작 직후 0% 상태를 먼저 전달해 UI가 즉시 반응하도록 한다.
        // 병합 시작 직후 0% 상태를 먼저 전달해 UI가 바로 반응하도록 한다.
        listener.onProgress(0, "Preparing merge...");
        // 템플릿이 없으면 이후 병합 흐름 전체가 성립하지 않으므로 즉시 예외를 발생시킨다.
        if (!Files.exists(templatePath)) {
            throw new IOException("Template file not found: " + templatePath.toAbsolutePath());
        }

        try (InputStream templateStream = Files.newInputStream(templatePath);
                XWPFDocument targetDocument = new XWPFDocument(templateStream)) {
            completedSteps = reportProgress(listener, completedSteps + 1, totalSteps, "Template loaded.");
            applyReportDate(targetDocument);
            completedSteps = reportProgress(listener, completedSteps + 1, totalSteps, "Applied report date.");

            List<SourceDocument> sourceDocuments = new ArrayList<>();
            try {
                for (Path sourceFile : sourceFiles) {
                    sourceDocuments.add(new SourceDocument(sourceFile, openDocument(sourceFile)));
                    completedSteps = reportProgress(
                            listener,
                            completedSteps + 1,
                            totalSteps,
                            "Loaded source file: " + sourceFile.getFileName());
                }

                // 결과 문서는 템플릿 섹션 순서를 유지하면서 각 섹션별 병합을 수행한다.
                // 템플릿 섹션 순서대로 병합해 최종 보고서 레이아웃이 기대한 형태를 유지하도록 한다.
                mergeAppendSection(targetDocument, sourceDocuments, PROJECT_STATUS_SECTION);
                completedSteps = reportProgress(listener, completedSteps + 1, totalSteps, "Merged project status section.");
                mergeStaffSection(targetDocument, sourceDocuments);
                completedSteps = reportProgress(listener, completedSteps + 1, totalSteps, "Merged staff status section.");
                mergeAppendSection(targetDocument, sourceDocuments, SALES_STATUS_SECTION);
                completedSteps = reportProgress(listener, completedSteps + 1, totalSteps, "Merged sales status section.");
                appendIssuesSectionToDocumentEnd(targetDocument, sourceDocuments);
                completedSteps = reportProgress(listener, completedSteps + 1, totalSteps, "Appended issues section.");

                // 출력 파일명은 실행일 기준 yyyyMMdd 형식 접두사와 고정 suffix로 구성된다.
                // 현재 날짜 접두어와 고정 suffix를 조합해 결과 파일명을 만든다.
                Path outputFile = outputDirectory.resolve(
                        LocalDate.now().format(DateTimeFormatter.BASIC_ISO_DATE) + OUTPUT_FILE_SUFFIX);
                try (OutputStream outputStream = Files.newOutputStream(outputFile)) {
                    targetDocument.write(outputStream);
                }
                reportProgress(listener, totalSteps, totalSteps, "Merge completed.");
                return outputFile;
            } finally {
                // 원본 문서를 모두 닫아 파일 잠금과 리소스 누수를 방지한다.
                // 파일 잠금과 리소스 누수를 막기 위해 연 원본 문서는 모두 닫는다.
                for (SourceDocument sourceDocument : sourceDocuments) {
                    sourceDocument.document().close();
                }
            }
        }
    }

    /**
     * 파일 경로를 기준으로 Word 문서를 열어 POI 문서 객체로 변환한다.
     *
     * @param sourceFile 열 대상 파일 경로
     * @return 메모리에 로드된 Word 문서
     * @throws IOException 파일을 읽지 못한 경우
     */
    /**
     * 디스크에서 Word 문서를 열어 메모리에 로드한다.
     *
     * @param sourceFile 원본 문서 경로
     * @return 로드된 Word 문서
     * @throws IOException 파일을 열 수 없는 경우
     */
    private static XWPFDocument openDocument(Path sourceFile) throws IOException {
        try (InputStream inputStream = Files.newInputStream(sourceFile)) {
            return new XWPFDocument(inputStream);
        }
    }

    /**
     * 완료된 단계 수를 퍼센트로 환산해 진행 콜백에 전달한다.
     *
     * @param listener 진행 상태를 수신할 리스너
     * @param completedSteps 현재까지 완료한 단계 수
     * @param totalSteps 전체 단계 수
     * @param message 사용자에게 보여 줄 현재 단계 설명
     * @return 전달한 완료 단계 수
     */
    /**
     * 완료된 단계 수를 퍼센트로 환산해 진행률 리스너에 전달한다.
     *
     * @param listener 진행률 리스너
     * @param completedSteps 현재까지 완료한 병합 단계 수
     * @param totalSteps 전체 병합 단계 수
     * @param message 현재 단계 설명
     * @return 완료 단계 수
     */
    private static int reportProgress(
            ProgressListener listener,
            int completedSteps,
            int totalSteps,
            String message) {
        int percent = totalSteps <= 0 ? 100 : Math.min(100, (completedSteps * 100) / totalSteps);
        listener.onProgress(percent, message);
        return completedSteps;
    }

    /**
     * 문서 전체에서 보고일 플레이스홀더를 찾아 다음 주 월요일 문자열로 치환한다.
     *
     * @param document 플레이스홀더를 치환할 대상 문서
     */
    private static void applyReportDate(XWPFDocument document) {
        String reportDate = buildNextMondayReportDate();
        for (XWPFParagraph paragraph : document.getParagraphs()) {
            replacePlaceholderInParagraph(paragraph, REPORT_DATE_PLACEHOLDER, reportDate);
        }
        for (XWPFTable table : document.getTables()) {
            replacePlaceholderInTable(table, REPORT_DATE_PLACEHOLDER, reportDate);
        }
    }

    /**
     * 표와 표 안의 중첩 표까지 재귀적으로 순회하며 플레이스홀더를 치환한다.
     *
     * @param table 치환할 대상 표
     * @param placeholder 찾을 문자열
     * @param replacement 대체할 문자열
     */
    private static void replacePlaceholderInTable(XWPFTable table, String placeholder, String replacement) {
        for (XWPFTableRow row : table.getRows()) {
            for (XWPFTableCell cell : row.getTableCells()) {
                for (XWPFParagraph paragraph : cell.getParagraphs()) {
                    replacePlaceholderInParagraph(paragraph, placeholder, replacement);
                }
                // 셀 안의 중첩 표도 같은 규칙으로 재귀 치환한다.
                for (XWPFTable nestedTable : cell.getTables()) {
                    replacePlaceholderInTable(nestedTable, placeholder, replacement);
                }
            }
        }
    }

    /**
     * 문단 내 플레이스홀더를 치환하고 첫 run의 서식을 최대한 유지한다.
     *
     * @param paragraph 치환 대상 문단
     * @param placeholder 찾을 문자열
     * @param replacement 대체할 문자열
     */
    private static void replacePlaceholderInParagraph(
            XWPFParagraph paragraph,
            String placeholder,
            String replacement) {
        String paragraphText = paragraph.getText();
        // 플레이스홀더가 없는 문단은 수정하지 않아 불필요한 스타일 변경을 막는다.
        if (paragraphText == null || !paragraphText.contains(placeholder)) {
            return;
        }

        String updatedText = paragraphText.replace(placeholder, replacement);
        CTRPr runProperties = null;
        if (!paragraph.getRuns().isEmpty() && paragraph.getRuns().get(0).getCTR().isSetRPr()) {
            runProperties = (CTRPr) paragraph.getRuns().get(0).getCTR().getRPr().copy();
        }

        for (int runIndex = paragraph.getRuns().size() - 1; runIndex >= 0; runIndex--) {
            paragraph.removeRun(runIndex);
        }

        XWPFRun run = paragraph.createRun();
        if (runProperties != null) {
            run.getCTR().setRPr(runProperties);
        }
        // 치환 후 문단 전체 텍스트는 하나의 run으로 다시 기록된다.
        run.setText(updatedText);
    }

    /**
     * 오늘이 월요일이면 오늘 날짜를, 그렇지 않으면 다음 주 월요일 날짜를 `yyyy-MM-dd (요일)` 형식으로 생성한다.
     *
     * @return 예: `2026-04-20 (월요일)` 형태의 보고일 문자열
     */
    private static String buildNextMondayReportDate() {
        LocalDate nextMonday = LocalDate.now().with(TemporalAdjusters.nextOrSame(DayOfWeek.MONDAY));
        return nextMonday.format(DateTimeFormatter.ISO_LOCAL_DATE) + " (" + getKoreanDayName(nextMonday.getDayOfWeek()) + ")";
    }

    /**
     * 요일 enum을 한국어 요일 문자열로 변환한다.
     *
     * @param dayOfWeek 변환할 요일
     * @return 한국어 요일명
     */
    private static String getKoreanDayName(DayOfWeek dayOfWeek) {
        return switch (dayOfWeek) {
            case MONDAY -> "월요일";
            case TUESDAY -> "화요일";
            case WEDNESDAY -> "수요일";
            case THURSDAY -> "목요일";
            case FRIDAY -> "금요일";
            case SATURDAY -> "토요일";
            case SUNDAY -> "일요일";
        };
    }

    /**
     * 일반 섹션 표를 병합한다.
     * 영업 현황 섹션은 첫 번째 열을 보존해야 하므로 별도 분기 처리한다.
     *
     * @param targetDocument 병합 결과를 기록할 템플릿 문서
     * @param sourceDocuments 병합할 원본 문서 목록
     * @param sectionSpec 섹션 제목과 표 위치 정보
     * @throws IOException 템플릿 섹션 표를 찾지 못한 경우
     */
    private static void mergeAppendSection(
            XWPFDocument targetDocument,
            List<SourceDocument> sourceDocuments,
            SectionSpec sectionSpec) throws IOException {
        XWPFTable targetTable = findSectionTable(targetDocument, sectionSpec);
        if (targetTable == null) {
            throw new IOException("Template section table not found. Expected one of: "
                    + String.join(", ", sectionSpec.titles())
                    + " or fallback table index " + (sectionSpec.tableIndex() + 1));
        }

        List<List<String>> mergedRows = new ArrayList<>();
        List<String> templateFirstColumnValues = List.of();
        if (sectionSpec == SALES_STATUS_SECTION) {
            // 영업 현황은 템플릿 첫 열을 유지하고 나머지 열만 원본 데이터로 다시 채운다.
            templateFirstColumnValues = extractTemplateColumnValues(targetTable, SALES_EXCLUDED_COLUMN_INDEX);
            List<SalesRowData> salesRows = new ArrayList<>();
            for (SourceDocument sourceDocument : sourceDocuments) {
                XWPFTable sourceTable = findSectionTable(sourceDocument.document(), sectionSpec);
                if (sourceTable == null) {
                    continue;
                }
                salesRows.addAll(extractSalesRows(sourceTable, sectionSpec.dataStartRowIndex()));
            }

            rewriteSalesTableRowsPreservingColumn(
                    targetTable,
                    salesRows,
                    templateFirstColumnValues,
                    SALES_EXCLUDED_COLUMN_INDEX);
            return;
        }

        for (SourceDocument sourceDocument : sourceDocuments) {
            XWPFTable sourceTable = findSectionTable(sourceDocument.document(), sectionSpec);
            if (sourceTable == null) {
                continue;
            }
            List<List<String>> sectionRows = extractDataRows(sourceTable, sectionSpec.dataStartRowIndex());
            mergedRows.addAll(sectionRows);
        }

        // 일반 섹션은 추출된 행을 순서대로 이어 붙여 템플릿 표를 다시 작성한다.
        rewriteTableRows(targetTable, mergedRows);
    }

    /**
     * 인력 현황 섹션을 카테고리 기준으로 병합한다.
     * 숫자 셀은 합산하고, 문자열 셀은 줄바꿈으로 이어 붙이는 동작을 기대한다.
     *
     * @param targetDocument 병합 결과를 기록할 템플릿 문서
     * @param sourceDocuments 병합할 원본 문서 목록
     * @throws IOException 템플릿 섹션 표를 찾지 못한 경우
     */
    private static void mergeStaffSection(
            XWPFDocument targetDocument,
            List<SourceDocument> sourceDocuments) throws IOException {
        XWPFTable targetTable = findSectionTable(targetDocument, STAFF_STATUS_SECTION);
        if (targetTable == null) {
            throw new IOException("Template section table not found. Expected one of: "
                    + String.join(", ", STAFF_STATUS_SECTION.titles())
                    + " or fallback table index " + (STAFF_STATUS_SECTION.tableIndex() + 1));
        }

        List<String> templateCategories = extractTemplateColumnValues(targetTable, STAFF_CATEGORY_COLUMN_INDEX);
        Map<String, Integer> templateCategoryIndexes = buildTemplateCategoryIndexes(templateCategories);
        Map<Integer, Map<Integer, CellAccumulator>> accumulators = new LinkedHashMap<>();

        for (SourceDocument sourceDocument : sourceDocuments) {
            XWPFTable sourceTable = findSectionTable(sourceDocument.document(), STAFF_STATUS_SECTION);
            if (sourceTable == null) {
                continue;
            }

            List<List<String>> rows = extractDataRows(sourceTable, STAFF_STATUS_SECTION.dataStartRowIndex());
            for (int rowIndex = 0; rowIndex < rows.size(); rowIndex++) {
                List<String> row = rows.get(rowIndex);
                // 카테고리명이 일치하면 해당 템플릿 행에 누적하고, 아니면 원래 순서를 fallback으로 사용한다.
                int targetRowIndex = resolveStaffTargetRowIndex(row, rowIndex, templateCategoryIndexes,
                        templateCategories);
                if (targetRowIndex < 0) {
                    continue;
                }

                Map<Integer, CellAccumulator> rowAccumulator = accumulators.computeIfAbsent(targetRowIndex,
                        key -> new LinkedHashMap<>());
                for (int colIndex = 0; colIndex < row.size(); colIndex++) {
                    if (colIndex == STAFF_CATEGORY_COLUMN_INDEX) {
                        // 첫 번째 열은 분류 기준이므로 누적 대상에서 제외한다.
                        continue;
                    }

                    String cellValue = row.get(colIndex);
                    // 2열~4열(인덱스 1~3)은 숫자만 집계 대상이므로 비숫자 값은 건너뛴다.
                    if (colIndex >= 1 && colIndex <= 3
                            && parseNumber(cleanCellText(cellValue)) == null) {
                        continue;
                    }

                    rowAccumulator
                            .computeIfAbsent(colIndex, key -> new CellAccumulator())
                            .add(cellValue);
                }
            }
        }

        List<List<String>> mergedRows = new ArrayList<>();
        List<Integer> accumulatorKeys = new ArrayList<>();
        for (Map.Entry<Integer, Map<Integer, CellAccumulator>> entry : accumulators.entrySet()) {
            accumulatorKeys.add(entry.getKey());
            Map<Integer, CellAccumulator> rowAccumulator = entry.getValue();
            int maxColumnIndex = rowAccumulator.keySet().stream().mapToInt(Integer::intValue).max().orElse(-1);
            List<String> mergedRow = new ArrayList<>();
            for (int colIndex = 0; colIndex <= maxColumnIndex; colIndex++) {
                CellAccumulator accumulator = rowAccumulator.get(colIndex);
                mergedRow.add(accumulator == null ? "" : accumulator.getMergedValue());
            }
            mergedRows.add(mergedRow);
        }

        applyStaffSumCalculations(mergedRows, accumulatorKeys, templateCategories);

        // 최종 출력 시 카테고리 열은 템플릿 기준 값을 유지해 표 구조를 안정적으로 맞춘다.
        rewriteStaffTableRows(targetTable, mergedRows, templateCategories);
    }

    /**
     * 주요 이슈 섹션을 원본 문서별로 추출해 결과 문서 맨 뒤에 덧붙인다.
     *
     * @param targetDocument 내용을 추가할 대상 문서
     * @param sourceDocuments 원본 문서 목록
     */
    private static void appendIssuesSectionToDocumentEnd(
            XWPFDocument targetDocument,
            List<SourceDocument> sourceDocuments) {
        boolean appendedAnyContent = false;
        for (SourceDocument sourceDocument : sourceDocuments) {
            List<IBodyElement> sectionElements = extractSectionElementsFromStartToEnd(
                    sourceDocument.document(),
                    ISSUES_SECTION);
            if (sectionElements.isEmpty()) {
                continue;
            }

            if (appendedAnyContent) {
                // 두 번째 팀부터는 빈 문단을 추가해 섹션 경계를 눈에 띄게 만든다.
                appendBlankParagraph(targetDocument);
            }

            appendPlainParagraph(
                    targetDocument,
                    "[" + extractTeamName(sourceDocument.path()) + "]");

            for (IBodyElement bodyElement : sectionElements) {
                if (bodyElement instanceof XWPFParagraph paragraph
                        && isIssuesTitleParagraph(paragraph)) {
                    // 팀명을 별도 헤더로 추가하므로 원본 섹션 제목은 중복 출력하지 않는다.
                    continue;
                }
                appendBodyElement(targetDocument, bodyElement);
            }

            appendedAnyContent = true;
        }
    }

    /**
     * 병합된 인력 현황 행에 합계 로직을 적용한다.
     * 데이터 행의 4열은 2열 + 3열로 계산하고, "합계" 행은 각 열의 데이터 합으로 재산정한다.
     *
     * @param mergedRows 병합된 행 목록 (인덱스 0: 카테고리 자리, 1: 2열, 2: 3열, 3: 4열)
     * @param accumulatorKeys mergedRows 각 행에 대응하는 templateCategories 인덱스
     * @param templateCategories 템플릿 카테고리 목록
     */
    private static void applyStaffSumCalculations(
            List<List<String>> mergedRows,
            List<Integer> accumulatorKeys,
            List<String> templateCategories) {
        int sumRowInMergedRows = -1;
        for (int i = 0; i < accumulatorKeys.size(); i++) {
            int templateIndex = accumulatorKeys.get(i);
            if (templateIndex < templateCategories.size()
                    && isSumCategory(templateCategories.get(templateIndex))) {
                sumRowInMergedRows = i;
                break;
            }
        }

        // 데이터 행: 4열(인덱스 3) = 2열(인덱스 1) + 3열(인덱스 2) 계산값으로 대체
        for (int i = 0; i < mergedRows.size(); i++) {
            if (i == sumRowInMergedRows) {
                continue;
            }
            List<String> row = mergedRows.get(i);
            BigDecimal col2 = safeParseNumber(row.size() > 1 ? row.get(1) : "");
            BigDecimal col3 = safeParseNumber(row.size() > 2 ? row.get(2) : "");
            String col4Value = formatStaffValue(col2.add(col3));
            if (row.size() > 3) {
                row.set(3, col4Value);
            } else {
                while (row.size() < 3) {
                    row.add("");
                }
                row.add(col4Value);
            }
        }

        // 합계 행: 2열과 3열은 데이터 행 합산, 4열은 두 합계의 합으로 재산정
        if (sumRowInMergedRows >= 0) {
            BigDecimal sumCol2 = BigDecimal.ZERO;
            BigDecimal sumCol3 = BigDecimal.ZERO;
            for (int i = 0; i < mergedRows.size(); i++) {
                if (i == sumRowInMergedRows) {
                    continue;
                }
                List<String> row = mergedRows.get(i);
                sumCol2 = sumCol2.add(safeParseNumber(row.size() > 1 ? row.get(1) : ""));
                sumCol3 = sumCol3.add(safeParseNumber(row.size() > 2 ? row.get(2) : ""));
            }
            List<String> sumRow = mergedRows.get(sumRowInMergedRows);
            while (sumRow.size() <= 3) {
                sumRow.add("");
            }
            sumRow.set(1, formatStaffValue(sumCol2));
            sumRow.set(2, formatStaffValue(sumCol3));
            sumRow.set(3, formatStaffValue(sumCol2.add(sumCol3)));
        }
    }

    private static boolean isSumCategory(String category) {
        return "합계".equals(normalizeText(category));
    }

    // '-' 값은 숫자로 파싱되지 않아 0으로 처리되며, 합계가 0이면 공백을 반환한다.
    private static String formatStaffValue(BigDecimal value) {
        return value.compareTo(BigDecimal.ZERO) == 0 ? "" : NUMBER_FORMAT.format(value);
    }

    private static BigDecimal safeParseNumber(String value) {
        BigDecimal result = parseNumber(value);
        return result != null ? result : BigDecimal.ZERO;
    }

    /**
     * 템플릿의 인력 카테고리명을 정규화해 행 인덱스와 매핑한다.
     *
     * @param templateCategories 템플릿 표에서 읽은 카테고리명 목록
     * @return 정규화된 카테고리명과 대상 행 번호 매핑
     */
    private static Map<String, Integer> buildTemplateCategoryIndexes(List<String> templateCategories) {
        Map<String, Integer> indexes = new LinkedHashMap<>();
        for (int rowIndex = 0; rowIndex < templateCategories.size(); rowIndex++) {
            String normalizedCategory = normalizeText(templateCategories.get(rowIndex));
            // 동일한 카테고리가 중복되면 첫 번째 위치를 기준으로 사용해 출력 구조를 안정적으로 유지한다.
            if (!normalizedCategory.isEmpty()) {
                indexes.putIfAbsent(normalizedCategory, rowIndex);
            }
        }
        return indexes;
    }

    /**
     * 인력 현황의 한 행이 템플릿 어느 행에 매핑될지 결정한다.
     *
     * @param row 현재 원본 데이터 행
     * @param fallbackRowIndex 카테고리 매칭 실패 시 사용할 기본 행 번호
     * @param templateCategoryIndexes 템플릿 카테고리와 행 번호 매핑
     * @param templateCategories 템플릿 카테고리 목록
     * @return 대상 행 번호, 매핑 불가 시 -1
     */
    private static int resolveStaffTargetRowIndex(
            List<String> row,
            int fallbackRowIndex,
            Map<String, Integer> templateCategoryIndexes,
            List<String> templateCategories) {
        if (!row.isEmpty()) {
            String normalizedCategory = normalizeText(row.get(STAFF_CATEGORY_COLUMN_INDEX));
            Integer matchedRowIndex = templateCategoryIndexes.get(normalizedCategory);
            if (matchedRowIndex != null) {
                return matchedRowIndex;
            }
        }

        // fallback 행도 템플릿 범위를 벗어나면 기록하지 않는다.
        return fallbackRowIndex < templateCategories.size() ? fallbackRowIndex : -1;
    }

    /**
     * 제목 문단 또는 표 내부 텍스트를 바탕으로 섹션의 표를 찾는다.
     *
     * @param document 탐색할 문서
     * @param sectionSpec 섹션 식별 정보
     * @return 찾은 표, 없으면 null
     */
    private static XWPFTable findSectionTable(XWPFDocument document, SectionSpec sectionSpec) {
        List<String> normalizedTitles = sectionSpec.titles().stream()
                .map(ReportMerger::normalizeText)
                .toList();
        boolean nextTableMatches = false;

        for (IBodyElement bodyElement : document.getBodyElements()) {
            if (bodyElement instanceof XWPFParagraph) {
                String paragraphText = normalizeText(((XWPFParagraph) bodyElement).getText());
                if (containsAny(paragraphText, normalizedTitles)) {
                    // 제목 문단을 찾으면 바로 다음 표를 해당 섹션 표로 간주한다.
                    nextTableMatches = true;
                } else if (!paragraphText.isEmpty()) {
                    nextTableMatches = false;
                }
                continue;
            }

            if (bodyElement instanceof XWPFTable) {
                XWPFTable table = (XWPFTable) bodyElement;
                if (nextTableMatches) {
                    return table;
                }

                if (containsAny(normalizeText(table.getText()), normalizedTitles)) {
                    // 제목 문단이 없더라도 표 내부에 섹션명이 포함되면 보조 규칙으로 매칭한다.
                    return table;
                }

                nextTableMatches = false;
            }
        }

        if (document.getTables().size() > sectionSpec.tableIndex()) {
            return document.getTables().get(sectionSpec.tableIndex());
        }

        return null;
    }

    /**
     * 특정 섹션 제목이 시작된 지점부터 문서 끝까지의 본문 요소를 추출한다.
     *
     * @param document 탐색할 문서
     * @param sectionSpec 섹션 식별 정보
     * @return 시작 지점 이후의 문단/표 목록
     */
    private static List<IBodyElement> extractSectionElementsFromStartToEnd(
            XWPFDocument document,
            SectionSpec sectionSpec) {
        List<String> normalizedTitles = sectionSpec.titles().stream()
                .map(ReportMerger::normalizeText)
                .toList();
        List<IBodyElement> elements = new ArrayList<>();
        boolean collecting = false;

        for (IBodyElement bodyElement : document.getBodyElements()) {
            if (bodyElement instanceof XWPFParagraph paragraph) {
                String text = cleanCellText(paragraph.getText());
                String normalizedText = normalizeText(text);

                if (!collecting && containsAny(normalizedText, normalizedTitles)) {
                    // 주요 이슈는 끝 경계 없이 문서 끝까지 이어진다고 보고 수집을 시작한다.
                    collecting = true;
                }

                if (collecting) {
                    elements.add(bodyElement);
                }
                continue;
            }

            if (collecting) {
                elements.add(bodyElement);
            }
        }

        return elements;
    }

    /**
     * 문단이 "주요 이슈" 섹션 제목인지 판별한다.
     *
     * @param paragraph 검사할 문단
     * @return 이슈 섹션 제목이면 true
     */
    private static boolean isIssuesTitleParagraph(XWPFParagraph paragraph) {
        String normalizedText = normalizeText(cleanCellText(paragraph.getText()));
        List<String> normalizedTitles = ISSUES_SECTION.titles().stream()
                .map(ReportMerger::normalizeText)
                .toList();
        return containsAny(normalizedText, normalizedTitles);
    }

    /**
     * 결과 문서 끝에 빈 줄 하나를 추가해 섹션 사이 간격을 만든다.
     *
     * @param document 빈 문단을 추가할 대상 문서
     */
    private static void appendBlankParagraph(XWPFDocument document) {
        XmlCursor cursor = document.getDocument().getBody().newCursor();
        cursor.toEndToken();
        XWPFParagraph paragraph = document.insertNewParagraph(cursor);
        cursor.dispose();
        paragraph.createRun().addBreak();
    }

    /**
     * 결과 문서 끝에 단순 텍스트 문단을 추가한다.
     * 현재는 팀명을 구분 헤더처럼 보여 주는 용도로 사용한다.
     *
     * @param document 문단을 추가할 대상 문서
     * @param text 추가할 텍스트
     */
    private static void appendPlainParagraph(XWPFDocument document, String text) {
        XmlCursor cursor = document.getDocument().getBody().newCursor();
        cursor.toEndToken();
        XWPFParagraph paragraph = document.insertNewParagraph(cursor);
        cursor.dispose();
        paragraph.setAlignment(ParagraphAlignment.LEFT);
        XWPFRun run = paragraph.createRun();
        run.setFontSize(12);
        run.setText(text);
    }

    /**
     * 원본 문서의 문단 또는 표를 결과 문서 끝으로 복사한다.
     *
     * @param targetDocument 복사 대상을 붙여 넣을 결과 문서
     * @param bodyElement 복사할 원본 본문 요소
     */
    private static void appendBodyElement(XWPFDocument targetDocument, IBodyElement bodyElement) {
        XmlCursor cursor = targetDocument.getDocument().getBody().newCursor();
        cursor.toEndToken();

        if (bodyElement instanceof XWPFParagraph paragraph) {
            XWPFParagraph appendedParagraph = targetDocument.insertNewParagraph(cursor);
            appendedParagraph.getCTP().set((CTP) paragraph.getCTP().copy());
        } else if (bodyElement instanceof XWPFTable table) {
            XWPFTable appendedTable = targetDocument.insertNewTbl(cursor);
            appendedTable.getCTTbl().set((CTTbl) table.getCTTbl().copy());
        }

        cursor.dispose();
    }

    /**
     * 파일명에서 팀명을 추출한다.
     * 확장자를 제거한 뒤 마지막 `_` 뒤의 문자열을 팀명으로 사용한다.
     *
     * @param sourceFile 팀명이 포함된 원본 파일 경로
     * @return 추출된 팀명
     */
    private static String extractTeamName(Path sourceFile) {
        String fileName = sourceFile.getFileName().toString();
        int extensionIndex = fileName.lastIndexOf('.');
        if (extensionIndex >= 0) {
            fileName = fileName.substring(0, extensionIndex);
        }

        int lastUnderscoreIndex = fileName.lastIndexOf('_');
        if (lastUnderscoreIndex >= 0 && lastUnderscoreIndex < fileName.length() - 1) {
            return fileName.substring(lastUnderscoreIndex + 1).trim();
        }

        return fileName.trim();
    }

    /**
     * 문자열에 후보 텍스트 중 하나라도 포함되어 있는지 확인한다.
     *
     * @param text 검사할 원본 문자열
     * @param candidates 포함 여부를 확인할 후보 문자열 목록
     * @return 하나라도 포함되면 true
     */
    private static boolean containsAny(String text, List<String> candidates) {
        return candidates.stream().anyMatch(text::contains);
    }

    /**
     * 행 전체가 영업 섹션의 고정 헤더 행인지 판별한다.
     *
     * @param row 비교할 행 값 목록
     * @param normalizedFixedRowValues 정규화된 고정 헤더 문자열 목록
     * @return 모든 고정 문자열을 포함하는 헤더 행이면 true
     */
    private static boolean matchesFixedRow(List<String> row, List<String> normalizedFixedRowValues) {
        String normalizedJoinedRow = row.stream()
                .map(ReportMerger::normalizeText)
                .reduce("", String::concat);

        for (String fixedValue : normalizedFixedRowValues) {
            // 헤더를 구성하는 값 중 하나라도 없으면 실제 데이터 행으로 간주한다.
            if (!normalizedJoinedRow.contains(fixedValue)) {
                return false;
            }
        }

        return !normalizedJoinedRow.isEmpty();
    }

    /**
     * 영업 현황 표에서 실제 데이터 행만 추출한다.
     * 반복 헤더와 빈 행을 제거한 뒤, 템플릿이 보존할 첫 번째 열은 제외한다.
     *
     * @param table 원본 영업 현황 표
     * @param dataStartRowIndex 데이터가 시작되는 행 인덱스
     * @return 정제된 영업 행 목록
     */
    private static List<SalesRowData> extractSalesRows(XWPFTable table, int dataStartRowIndex) {
        List<SalesRowData> rows = new ArrayList<>();
        for (int rowIndex = Math.max(0, dataStartRowIndex); rowIndex < table.getNumberOfRows(); rowIndex++) {
            XWPFTableRow row = table.getRow(rowIndex);
            if (row == null) {
                continue;
            }

            List<String> rowValues = extractCellValues(row);
            if (!isEmptyRow(rowValues)) {
                rows.add(new SalesRowData(rowValues, new ArrayList<>(row.getTableCells())));
            }
        }

        // 원본 문서 중간에 반복 삽입된 헤더 행이 병합 데이터로 들어오지 않게 제거한다.
        rows = filterOutFixedSalesRows(rows, SALES_FIXED_HEADER_ROW);
        rows = removeFirstSalesColumn(rows);
        rows = filterOutFixedSalesRows(rows, SALES_FIXED_HEADER_ROW_WITHOUT_FIRST_COLUMN);
        return rows;
    }

    /**
     * 영업 현황 표에서 반복 삽입된 고정 헤더 행을 제거한다.
     *
     * @param rows 후보 영업 행 목록
     * @param fixedRowValues 제거할 헤더 행의 기준 문자열
     * @return 고정 헤더가 제거된 영업 행 목록
     */
    private static List<SalesRowData> filterOutFixedSalesRows(List<SalesRowData> rows, List<String> fixedRowValues) {
        List<String> normalizedFixedRowValues = fixedRowValues.stream()
                .map(ReportMerger::normalizeText)
                .toList();
        List<SalesRowData> filteredRows = new ArrayList<>();

        for (SalesRowData row : rows) {
            if (!matchesFixedRow(row.values(), normalizedFixedRowValues)) {
                filteredRows.add(row);
            }
        }

        return filteredRows;
    }

    /**
     * 영업 현황 데이터에서 템플릿이 보존할 첫 번째 열을 제거한 복사본을 만든다.
     *
     * @param rows 원본 영업 행 목록
     * @return 첫 번째 열이 제거된 영업 행 목록
     */
    private static List<SalesRowData> removeFirstSalesColumn(List<SalesRowData> rows) {
        List<SalesRowData> adjustedRows = new ArrayList<>();
        for (SalesRowData row : rows) {
            List<String> adjustedValues = new ArrayList<>();
            List<XWPFTableCell> adjustedCells = new ArrayList<>();
            for (int index = 0; index < row.values().size(); index++) {
                if (index != SALES_EXCLUDED_COLUMN_INDEX) {
                    adjustedValues.add(row.values().get(index));
                }
            }
            for (int index = 0; index < row.sourceCells().size(); index++) {
                if (index != SALES_EXCLUDED_COLUMN_INDEX) {
                    adjustedCells.add(row.sourceCells().get(index));
                }
            }
            adjustedRows.add(new SalesRowData(adjustedValues, adjustedCells));
        }
        return adjustedRows;
    }

    /**
     * 섹션 제목 후보, 기본 표 위치, 실제 데이터 시작 행을 묶어 두는 설정 레코드다.
     */
    private record SectionSpec(List<String> titles, int tableIndex, int dataStartRowIndex) {
    }

    /**
     * 표에서 비어 있지 않은 데이터 행만 추출한다.
     *
     * @param table 원본 표
     * @param dataStartRowIndex 데이터 시작 행 인덱스
     * @return 셀 문자열 목록으로 구성된 행들
     */
    private static List<List<String>> extractDataRows(XWPFTable table, int dataStartRowIndex) {
        List<List<String>> rows = new ArrayList<>();
        for (int rowIndex = Math.max(0, dataStartRowIndex); rowIndex < table.getNumberOfRows(); rowIndex++) {
            XWPFTableRow row = table.getRow(rowIndex);
            if (row == null) {
                continue;
            }
            List<String> cellValues = extractCellValues(row);
            if (!isEmptyRow(cellValues)) {
                rows.add(cellValues);
            }
        }
        return rows;
    }

    /**
     * 표 행의 모든 셀 텍스트를 정리된 문자열 목록으로 추출한다.
     *
     * @param row 추출할 대상 행
     * @return 셀 값 목록
     */
    private static List<String> extractCellValues(XWPFTableRow row) {
        List<String> cellValues = new ArrayList<>();
        for (XWPFTableCell cell : row.getTableCells()) {
            cellValues.add(cleanCellText(cell.getText()));
        }
        return cellValues;
    }

    /**
     * 템플릿 표의 특정 열 값을 읽어 기준 데이터로 사용한다.
     *
     * @param table 템플릿 표
     * @param columnIndex 추출할 열 인덱스
     * @return 추출된 열 값 목록
     */
    private static List<String> extractTemplateColumnValues(XWPFTable table, int columnIndex) {
        List<String> values = new ArrayList<>();
        for (int rowIndex = 1; rowIndex < table.getNumberOfRows(); rowIndex++) {
            XWPFTableRow row = table.getRow(rowIndex);
            if (row == null || row.getTableCells().size() <= columnIndex) {
                values.add("");
                continue;
            }
            values.add(cleanCellText(row.getCell(columnIndex).getText()));
        }
        return values;
    }

    /**
     * 템플릿 표의 데이터 영역을 새 행 목록으로 다시 구성한다.
     *
     * @param table 다시 작성할 표
     * @param rows 기록할 데이터 행 목록
     * @throws IOException 템플릿 표 구조가 비정상인 경우
     */
    private static void rewriteTableRows(XWPFTable table, List<List<String>> rows) throws IOException {
        if (table == null || table.getNumberOfRows() == 0) {
            throw new IOException("Template table is empty.");
        }

        int templateRowIndex = table.getNumberOfRows() > 1 ? 1 : 0;
        XWPFTableRow templateRow = table.getRow(templateRowIndex);
        if (templateRow == null) {
            throw new IOException("Template table does not contain a usable row.");
        }

        for (int rowIndex = table.getNumberOfRows() - 1; rowIndex > templateRowIndex; rowIndex--) {
            table.removeRow(rowIndex);
        }

        if (rows.isEmpty()) {
            // 병합 대상이 없으면 템플릿의 첫 데이터 행만 비워 두는 결과를 만든다.
            clearRow(templateRow);
            return;
        }

        clearRow(templateRow);
        setRowValues(templateRow, rows.get(0));

        for (int rowIndex = 1; rowIndex < rows.size(); rowIndex++) {
            List<String> rowValues = rows.get(rowIndex);
            XWPFTableRow row = createStyledRow(table, templateRow);
            setRowValues(row, rowValues);
        }
    }

    /**
     * 인력 현황 표를 다시 작성한 뒤 카테고리 열을 템플릿 값으로 복원한다.
     *
     * @param table 대상 표
     * @param rows 병합된 데이터 행
     * @param templateCategories 템플릿 카테고리 목록
     * @throws IOException 템플릿 표 구조가 비정상인 경우
     */
    private static void rewriteStaffTableRows(
            XWPFTable table,
            List<List<String>> rows,
            List<String> templateCategories) throws IOException {
        rewriteTableRows(table, rows);

        int templateRowIndex = table.getNumberOfRows() > 1 ? 1 : 0;
        for (int rowOffset = 0; rowOffset < templateCategories.size(); rowOffset++) {
            int rowIndex = templateRowIndex + rowOffset;
            if (rowIndex >= table.getNumberOfRows()) {
                break;
            }
            XWPFTableRow row = table.getRow(rowIndex);
            if (row == null || row.getTableCells().size() <= STAFF_CATEGORY_COLUMN_INDEX) {
                continue;
            }
            // 첫 열은 누적 결과가 아니라 분류 기준을 보여주는 열이므로 템플릿 값을 유지한다.
            setCellText(row.getCell(STAFF_CATEGORY_COLUMN_INDEX), templateCategories.get(rowOffset));
        }
        applyStaffSummaryRowBoldStyle(table, templateCategories.size());
    }

    /**
     * 영업 현황 표를 다시 작성하되 특정 열은 템플릿 값을 유지한다.
     *
     * @param table 대상 표
     * @param rows 병합된 영업 데이터
     * @param preservedColumnValues 유지할 템플릿 열 값
     * @param preservedColumnIndex 유지할 열 인덱스
     * @throws IOException 템플릿 표 구조가 비정상인 경우
     */
    private static void rewriteSalesTableRowsPreservingColumn(
            XWPFTable table,
            List<SalesRowData> rows,
            List<String> preservedColumnValues,
            int preservedColumnIndex) throws IOException {
        if (table == null || table.getNumberOfRows() == 0) {
            throw new IOException("Template table is empty.");
        }

        int templateRowIndex = table.getNumberOfRows() > 1 ? 1 : 0;
        XWPFTableRow templateRow = table.getRow(templateRowIndex);
        if (templateRow == null) {
            throw new IOException("Template table does not contain a usable row.");
        }

        for (int rowIndex = table.getNumberOfRows() - 1; rowIndex > templateRowIndex; rowIndex--) {
            table.removeRow(rowIndex);
        }

        if (rows.isEmpty()) {
            clearRow(templateRow);
            // 영업 데이터가 없어도 보존 열은 남겨 템플릿 구성을 유지한다.
            if (templateRow.getTableCells().size() > preservedColumnIndex && !preservedColumnValues.isEmpty()) {
                setCellText(templateRow.getCell(preservedColumnIndex), preservedColumnValues.get(0));
            }
            return;
        }

        clearRow(templateRow);
        copySalesRow(templateRow, rows.get(0), preservedColumnIndex + 1);
        if (templateRow.getTableCells().size() > preservedColumnIndex && !preservedColumnValues.isEmpty()) {
            setCellText(templateRow.getCell(preservedColumnIndex), preservedColumnValues.get(0));
        }

        for (int rowIndex = 1; rowIndex < rows.size(); rowIndex++) {
            XWPFTableRow row = createStyledRow(table, templateRow);
            clearRow(row);
            copySalesRow(row, rows.get(rowIndex), preservedColumnIndex + 1);
            if (row.getTableCells().size() > preservedColumnIndex && rowIndex < preservedColumnValues.size()) {
                setCellText(row.getCell(preservedColumnIndex), preservedColumnValues.get(rowIndex));
            }
        }
    }

    /**
     * 영업 행의 셀 내용과 서식을 대상 행의 지정 열부터 복사한다.
     *
     * @param targetRow 복사 결과를 기록할 행
     * @param sourceRow 복사할 원본 영업 행
     * @param startColumnIndex 붙여 넣기를 시작할 대상 열 인덱스
     */
    private static void copySalesRow(XWPFTableRow targetRow, SalesRowData sourceRow, int startColumnIndex) {
        int targetColumnIndex = startColumnIndex;
        for (XWPFTableCell sourceCell : sourceRow.sourceCells()) {
            XWPFTableCell targetCell = getOrCreateCell(targetRow, targetColumnIndex++);
            copyCellContentPreservingStyle(sourceCell, targetCell);
        }
    }

    /**
     * 셀의 문단 구조와 내용을 복사해 서식 손실을 최소화한다.
     *
     * @param sourceCell 원본 셀
     * @param targetCell 대상 셀
     */
    private static void copyCellContentPreservingStyle(XWPFTableCell sourceCell, XWPFTableCell targetCell) {
        clearCellText(targetCell);

        while (targetCell.getParagraphs().size() > 1) {
            targetCell.removeParagraph(targetCell.getParagraphs().size() - 1);
        }

        for (int paragraphIndex = targetCell.getParagraphs().size() - 1; paragraphIndex >= 0; paragraphIndex--) {
            targetCell.removeParagraph(paragraphIndex);
        }

        for (XWPFParagraph sourceParagraph : sourceCell.getParagraphs()) {
            XWPFParagraph targetParagraph = targetCell.addParagraph();
            targetParagraph.getCTP().set((CTP) sourceParagraph.getCTP().copy());
        }

        if (targetCell.getParagraphs().isEmpty()) {
            targetCell.addParagraph();
        }
    }

    /**
     * 템플릿 행의 스타일을 복사한 새 행을 생성한다.
     *
     * @param table 행을 추가할 표
     * @param templateRow 스타일 기준 행
     * @return 스타일이 적용된 새 행
     */
    private static XWPFTableRow createStyledRow(XWPFTable table, XWPFTableRow templateRow) {
        XWPFTableRow newRow = table.createRow();
        copyRowStyle(templateRow, newRow);

        int targetCellCount = Math.max(templateRow.getTableCells().size(), newRow.getTableCells().size());
        while (newRow.getTableCells().size() < targetCellCount) {
            newRow.addNewTableCell();
        }

        for (int cellIndex = 0; cellIndex < targetCellCount; cellIndex++) {
            XWPFTableCell templateCell = getOrCreateCell(templateRow, cellIndex);
            XWPFTableCell targetCell = getOrCreateCell(newRow, cellIndex);
            copyCellStyle(templateCell, targetCell);
            clearCellText(targetCell);
        }

        return newRow;
    }

    /**
     * 요청한 열 인덱스의 셀을 반환하고, 부족하면 새 셀을 추가해 맞춘다.
     *
     * @param row 셀을 가져올 대상 행
     * @param cellIndex 필요한 셀 인덱스
     * @return 지정 위치의 셀
     */
    private static XWPFTableCell getOrCreateCell(XWPFTableRow row, int cellIndex) {
        while (row.getTableCells().size() <= cellIndex) {
            row.addNewTableCell();
        }
        return row.getCell(cellIndex);
    }

    /**
     * 템플릿 행의 높이와 행 속성을 대상 행에 복사한다.
     *
     * @param templateRow 서식 기준이 되는 템플릿 행
     * @param targetRow 서식을 적용할 대상 행
     */
    private static void copyRowStyle(XWPFTableRow templateRow, XWPFTableRow targetRow) {
        if (templateRow == null || targetRow == null) {
            return;
        }

        CTRow templateCtRow = templateRow.getCtRow();
        CTRow targetCtRow = targetRow.getCtRow();

        if (templateCtRow.isSetTrPr()) {
            CTTrPr copiedProperties = (CTTrPr) templateCtRow.getTrPr().copy();
            targetCtRow.setTrPr(copiedProperties);
        }

        targetRow.setHeight(templateRow.getHeight());
    }

    /**
     * 템플릿 셀의 테두리, 정렬 등 셀/문단 속성을 대상 셀에 복사한다.
     *
     * @param templateCell 서식 기준이 되는 셀
     * @param targetCell 서식을 적용할 대상 셀
     */
    private static void copyCellStyle(XWPFTableCell templateCell, XWPFTableCell targetCell) {
        if (templateCell == null || targetCell == null) {
            return;
        }

        if (templateCell.getCTTc().isSetTcPr()) {
            CTTcPr copiedProperties = (CTTcPr) templateCell.getCTTc().getTcPr().copy();
            targetCell.getCTTc().setTcPr(copiedProperties);
        }

        while (targetCell.getParagraphs().size() > 1) {
            targetCell.removeParagraph(targetCell.getParagraphs().size() - 1);
        }

        XWPFParagraph sourceParagraph = templateCell.getParagraphs().isEmpty()
                ? null
                : templateCell.getParagraphs().get(0);
        XWPFParagraph targetParagraph = targetCell.getParagraphs().isEmpty()
                ? targetCell.addParagraph()
                : targetCell.getParagraphs().get(0);

        if (sourceParagraph != null && sourceParagraph.getCTP().isSetPPr()) {
            CTPPr copiedProperties = (CTPPr) sourceParagraph.getCTP().getPPr().copy();
            targetParagraph.getCTP().setPPr(copiedProperties);
        }
    }

    /**
     * 행 안의 모든 셀 텍스트를 제거한다.
     *
     * @param row 비울 대상 행
     */
    private static void clearRow(XWPFTableRow row) {
        for (XWPFTableCell cell : row.getTableCells()) {
            clearCellText(cell);
        }
    }

    /**
     * 행의 각 셀에 문자열 값을 순서대로 기록한다.
     * 필요한 경우 부족한 셀을 자동으로 생성한다.
     *
     * @param row 값을 기록할 대상 행
     * @param values 기록할 셀 값 목록
     */
    private static void setRowValues(XWPFTableRow row, List<String> values) {
        int limit = Math.max(row.getTableCells().size(), values.size());
        for (int cellIndex = 0; cellIndex < limit; cellIndex++) {
            XWPFTableCell cell = getOrCreateCell(row, cellIndex);
            String value = cellIndex < values.size() ? values.get(cellIndex) : "";
            setCellText(cell, value);
        }
    }

    private static void applyStaffSummaryRowBoldStyle(XWPFTable table, int dataRowCount) {
        if (table == null || dataRowCount <= 0) {
            return;
        }

        int templateRowIndex = table.getNumberOfRows() > 1 ? 1 : 0;
        int summaryRowIndex = templateRowIndex + dataRowCount - 1;
        if (summaryRowIndex >= table.getNumberOfRows()) {
            return;
        }

        XWPFTableRow summaryRow = table.getRow(summaryRowIndex);
        if (summaryRow == null) {
            return;
        }

        for (int cellIndex = 1; cellIndex <= 3; cellIndex++) {
            if (cellIndex >= summaryRow.getTableCells().size()) {
                break;
            }
            setCellBold(summaryRow.getCell(cellIndex), true);
        }
    }

    /**
     * 셀 안의 run과 추가 문단을 제거해 빈 셀 상태로 만든다.
     *
     * @param cell 비울 대상 셀
     */
    private static void clearCellText(XWPFTableCell cell) {
        for (int paragraphIndex = cell.getParagraphs().size() - 1; paragraphIndex >= 0; paragraphIndex--) {
            XWPFParagraph paragraph = cell.getParagraphs().get(paragraphIndex);
            for (int runIndex = paragraph.getRuns().size() - 1; runIndex >= 0; runIndex--) {
                paragraph.removeRun(runIndex);
            }
            if (paragraphIndex > 0) {
                cell.removeParagraph(paragraphIndex);
            }
        }

        if (cell.getParagraphs().isEmpty()) {
            cell.addParagraph();
        }
    }

    /**
     * 셀에 텍스트를 기록한다.
     * 줄바꿈은 Word 줄바꿈으로 변환되므로 결과 문서에서도 여러 줄 출력이 유지된다.
     *
     * @param cell 값을 기록할 셀
     * @param value 기록할 문자열
     */
    private static void setCellText(XWPFTableCell cell, String value) {
        CTRPr runProperties = null;
        if (!cell.getParagraphs().isEmpty()) {
            XWPFParagraph sourceParagraph = cell.getParagraphs().get(0);
            if (!sourceParagraph.getRuns().isEmpty() && sourceParagraph.getRuns().get(0).getCTR().isSetRPr()) {
                runProperties = (CTRPr) sourceParagraph.getRuns().get(0).getCTR().getRPr().copy();
            }
        }

        clearCellText(cell);

        String[] lines = cleanCellText(value).split("\\R", -1);
        XWPFParagraph paragraph = cell.getParagraphs().get(0);
        XWPFRun run = paragraph.createRun();
        if (runProperties != null) {
            run.getCTR().setRPr(runProperties);
        }

        for (int lineIndex = 0; lineIndex < lines.length; lineIndex++) {
            if (lineIndex > 0) {
                // 여러 줄 문자열은 Word 내부 줄바꿈으로 순서대로 출력된다.
                run.addBreak();
            }
            run.setText(lines[lineIndex]);
        }
    }

    /**
     * 셀 텍스트 비교와 출력에 쓰기 좋도록 개행/공백을 정리한다.
     *
     * @param value 정리할 원본 문자열
     * @return 앞뒤 공백과 carriage return이 제거된 문자열
     */
    private static void setCellBold(XWPFTableCell cell, boolean bold) {
        if (cell == null) {
            return;
        }

        for (XWPFParagraph paragraph : cell.getParagraphs()) {
            if (paragraph.getRuns().isEmpty()) {
                paragraph.createRun().setBold(bold);
                continue;
            }

            for (XWPFRun run : paragraph.getRuns()) {
                run.setBold(bold);
            }
        }
    }

    private static String cleanCellText(String value) {
        return value == null ? "" : value.replace("\r", "").trim();
    }

    /**
     * 행의 모든 셀이 비어 있는지 확인한다.
     *
     * @param row 검사할 행 값 목록
     * @return 비어 있지 않은 값이 하나도 없으면 true
     */
    private static boolean isEmptyRow(List<String> row) {
        for (String value : row) {
            if (!value.isBlank()) {
                return false;
            }
        }
        return true;
    }

    /**
     * 공백 차이를 무시하고 비교할 수 있도록 문자열을 정규화한다.
     *
     * @param value 정규화할 원본 문자열
     * @return 모든 공백이 제거된 문자열
     */
    /**
     * 레이아웃 차이를 무시하고 비교할 수 있도록 모든 공백을 제거한다.
     *
     * @param value 정규화할 원본 문자열
     * @return 모든 공백이 제거된 문자열
     */
    private static String normalizeText(String value) {
        return value == null ? "" : value.replaceAll("\\s+", "");
    }

    /**
     * 문자열이 숫자 형식이면 BigDecimal로 변환한다.
     *
     * @param value 변환할 문자열
     * @return 숫자면 BigDecimal, 아니면 null
     */
    /**
     * 숫자 문자열을 {@link BigDecimal}로 변환한다.
     *
     * @param value 변환할 문자열
     * @return 변환된 숫자, 숫자가 아니면 {@code null}
     */
    private static BigDecimal parseNumber(String value) {
        String normalized = value.replace(",", "").trim();
        // 숫자가 아닌 값은 병합 시 텍스트로 처리되므로 여기서는 null을 반환한다.
        // 숫자가 아니면 합산 대신 문자열 병합 대상으로 처리하기 위해 null을 반환한다.
        if (!NUMERIC_PATTERN.matcher(normalized).matches()) {
            return null;
        }
        return new BigDecimal(normalized);
    }

    /**
     * 인력 현황 셀 값을 누적하기 위한 보조 클래스다.
     * 모든 값이 숫자면 합계를 반환하고, 하나라도 문자열이면 줄바꿈 연결 결과를 반환한다.
     */
    private static final class CellAccumulator {
        private final List<String> values = new ArrayList<>();
        private BigDecimal numericSum = BigDecimal.ZERO;
        private boolean hasValue;
        private boolean allNumeric = true;

        CellAccumulator() {
        }

        /**
         * 셀 값을 누적한다.
         * 숫자만 들어오면 합계를 계산하고, 문자가 섞이면 줄바꿈 연결 방식으로 전환한다.
         *
         * @param value 누적할 셀 문자열
         */
        /**
         * 셀 값을 누적한다.
         * 숫자만 들어오면 합계를 계산하고, 문자가 섞이면 줄 단위 텍스트를 유지한다.
         *
         * @param value 누적할 셀 텍스트
         */
        void add(String value) {
            String cleanedValue = cleanCellText(value);
            if (cleanedValue.isEmpty()) {
                // 빈 값은 병합 결과에 영향을 주지 않으므로 무시한다.
                return;
            }

            hasValue = true;
            values.add(cleanedValue);

            BigDecimal parsedNumber = parseNumber(cleanedValue);
            if (parsedNumber == null) {
                // 하나라도 숫자가 아니면 이 셀의 최종 결과는 문자열 병합으로 전환된다.
                allNumeric = false;
                return;
            }

            numericSum = numericSum.add(parsedNumber);
        }

        /**
         * 누적된 값을 최종 병합 문자열로 반환한다.
         *
         * @return 숫자 합계 또는 줄바꿈으로 연결된 문자열
         */
        /**
         * 누적된 입력값으로 최종 병합 결과 문자열을 만든다.
         *
         * @return 숫자 합계 또는 줄바꿈으로 연결한 텍스트
         */
        String getMergedValue() {
            if (!hasValue) {
                return "";
            }

            if (allNumeric) {
                // 숫자만 누적된 셀은 합산 결과가 최종 출력값이 된다.
                return NUMBER_FORMAT.format(numericSum);
            }

            // 문자열이 포함된 셀은 입력 순서를 유지하며 줄바꿈으로 이어 붙인다.
            return String.join("\n", values);
        }
    }

    private record SourceDocument(Path path, XWPFDocument document) {
    }

    private record SalesRowData(List<String> values, List<XWPFTableCell> sourceCells) {
    }
}
