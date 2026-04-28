import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.nio.file.Files;
import java.nio.file.Path;
import java.text.DecimalFormat;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.temporal.TemporalAdjusters;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.regex.Pattern;

import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFAbstractNum;
import org.apache.poi.xwpf.usermodel.XWPFNum;
import org.apache.poi.xwpf.usermodel.XWPFNumbering;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFStyle;
import org.apache.poi.xwpf.usermodel.XWPFStyles;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTAbstractNum;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTNum;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTNumLvl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTStyle;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTrPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;

/**
 * 여러 개의 주간 보고 Word 문서를 하나의 결과 문서로 병합하는 유틸리티 클래스다.
 * 템플릿 문서를 기준으로 병합을 수행하며, 가능한 범위에서 섹션 구조와 서식을 유지한다.
 */
public final class ReportMerger {
    // 템플릿 문서에서 실제 보고 일자로 치환할 플레이스홀더다.
    private static final String REPORT_DATE_PLACEHOLDER = "[취합일]";
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
    private static final String WORDPROCESSINGML_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static final javax.xml.namespace.QName WORD_VAL_ATTRIBUTE = new javax.xml.namespace.QName(
            WORDPROCESSINGML_NAMESPACE, "val");
    private static final Set<String> STYLE_REFERENCE_ELEMENTS = Set.of(
            "basedOn", "link", "next", "pStyle", "rStyle", "tblStyle");

    /**
     * 유틸리티 클래스이므로 외부에서 인스턴스를 생성하지 못하게 막는다.
     */
    private ReportMerger() {
    }

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
     * 외부 진행률 보고 없이 원본 보고서를 병합해 새 결과 문서를 만든다.
     *
     * @param sourceFiles     병합할 원본 보고서 문서 목록
     * @param outputDirectory 템플릿이 있고 결과 파일이 저장될 폴더
     * @return 생성된 결과 문서 경로
     * @throws IOException 템플릿 또는 원본 문서를 처리할 수 없는 경우
     */
    public static Path mergeReports(List<Path> sourceFiles, Path outputDirectory) throws IOException {
        return mergeReports(sourceFiles, outputDirectory, NO_OP_PROGRESS_LISTENER);
    }

    /**
     * 진행률을 보고하면서 원본 보고서를 병합해 새 결과 문서를 만든다.
     *
     * @param sourceFiles      병합할 원본 보고서 문서 목록
     * @param outputDirectory  템플릿이 있고 결과 파일이 저장될 폴더
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

        // 병합 시작 직후 0% 상태를 먼저 전달해 UI가 바로 반응하도록 한다.
        listener.onProgress(0, Messages.get("progress.preparing"));
        // 템플릿이 없으면 이후 병합 흐름 전체가 성립하지 않으므로 즉시 예외를 발생시킨다.
        if (!Files.exists(templatePath)) {
            throw new IOException(Messages.format("error.template.not.found", templatePath.toAbsolutePath()));
        }

        try (InputStream templateStream = Files.newInputStream(templatePath);
                XWPFDocument targetDocument = new XWPFDocument(templateStream)) {
            completedSteps = reportProgress(listener, completedSteps + 1, totalSteps,
                    Messages.get("progress.template.loaded"));
            applyReportDate(targetDocument);
            completedSteps = reportProgress(listener, completedSteps + 1, totalSteps,
                    Messages.get("progress.date.applied"));

            List<SourceDocument> sourceDocuments = new ArrayList<>();
            try {
                for (Path sourceFile : sourceFiles) {
                    sourceDocuments.add(new SourceDocument(sourceFile, openDocument(sourceFile)));
                    completedSteps = reportProgress(
                            listener,
                            completedSteps + 1,
                            totalSteps,
                            Messages.format("progress.file.loaded", sourceFile.getFileName()));
                }

                // 템플릿 섹션 순서대로 병합해 최종 보고서 레이아웃이 기대한 형태를 유지하도록 한다.
                mergeAppendSection(targetDocument, sourceDocuments, PROJECT_STATUS_SECTION);
                completedSteps = reportProgress(listener, completedSteps + 1, totalSteps,
                        Messages.get("progress.section.project"));
                mergeStaffSection(targetDocument, sourceDocuments);
                completedSteps = reportProgress(listener, completedSteps + 1, totalSteps,
                        Messages.get("progress.section.staff"));
                mergeAppendSection(targetDocument, sourceDocuments, SALES_STATUS_SECTION);
                completedSteps = reportProgress(listener, completedSteps + 1, totalSteps,
                        Messages.get("progress.section.sales"));
                appendIssuesSectionToDocumentEnd(targetDocument, sourceDocuments);
                completedSteps = reportProgress(listener, completedSteps + 1, totalSteps,
                        Messages.get("progress.section.issues"));

                // 현재 날짜 접두어와 고정 suffix를 조합해 결과 파일명을 만든다.
                Path outputFile = outputDirectory.resolve(
                        LocalDate.now().format(DateTimeFormatter.BASIC_ISO_DATE) + OUTPUT_FILE_SUFFIX);
                try (OutputStream outputStream = Files.newOutputStream(outputFile)) {
                    targetDocument.write(outputStream);
                }
                reportProgress(listener, totalSteps, totalSteps, Messages.get("progress.merge.completed"));
                return outputFile;
            } finally {
                // 파일 잠금과 리소스 누수를 막기 위해 연 원본 문서는 모두 닫는다.
                for (SourceDocument sourceDocument : sourceDocuments) {
                    sourceDocument.document().close();
                }
            }
        }
    }

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
     * 완료된 단계 수를 퍼센트로 환산해 진행률 리스너에 전달한다.
     *
     * @param listener       진행률 리스너
     * @param completedSteps 현재까지 완료한 병합 단계 수
     * @param totalSteps     전체 병합 단계 수
     * @param message        현재 단계 설명
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
        String reportDate = buildThisFridayDate();
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
     * @param table       치환할 대상 표
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
     * @param paragraph   치환 대상 문단
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
     * 오늘이 금요일이면 오늘 날짜를, 그렇지 않으면 금주 금요일 날짜를 `yyyy-MM-dd (요일)` 형식으로 생성한다.
     *
     * @return 예: `2026-04-24 (금)` 형태의 취합일 문자열
     */
    private static String buildThisFridayDate() {
        LocalDate friday = LocalDate.now().with(TemporalAdjusters.nextOrSame(DayOfWeek.FRIDAY));
        return friday.format(DateTimeFormatter.ISO_LOCAL_DATE) + " (" + getKoreanDayName(friday.getDayOfWeek())
                + ")";
    }

    /**
     * 요일 enum을 한국어 요일 문자열로 변환한다.
     *
     * @param dayOfWeek 변환할 요일
     * @return 한국어 요일명
     */
    private static String getKoreanDayName(DayOfWeek dayOfWeek) {
        return switch (dayOfWeek) {
            case MONDAY -> "월";
            case TUESDAY -> "화";
            case WEDNESDAY -> "수";
            case THURSDAY -> "목";
            case FRIDAY -> "금";
            case SATURDAY -> "토";
            case SUNDAY -> "일";
        };
    }

    /**
     * 일반 섹션 표를 병합한다.
     * 영업 현황 섹션은 첫 번째 열을 보존해야 하므로 별도 분기 처리한다.
     *
     * @param targetDocument  병합 결과를 기록할 템플릿 문서
     * @param sourceDocuments 병합할 원본 문서 목록
     * @param sectionSpec     섹션 제목과 표 위치 정보
     * @throws IOException 템플릿 섹션 표를 찾지 못한 경우
     */
    private static void mergeAppendSection(
            XWPFDocument targetDocument,
            List<SourceDocument> sourceDocuments,
            SectionSpec sectionSpec) throws IOException {
        XWPFTable targetTable = findSectionTable(targetDocument, sectionSpec);
        if (targetTable == null) {
            throw new IOException(Messages.format("error.template.section.not.found",
                    String.join(", ", sectionSpec.titles()), sectionSpec.tableIndex() + 1));
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
     * @param targetDocument  병합 결과를 기록할 템플릿 문서
     * @param sourceDocuments 병합할 원본 문서 목록
     * @throws IOException 템플릿 섹션 표를 찾지 못한 경우
     */
    private static void mergeStaffSection(
            XWPFDocument targetDocument,
            List<SourceDocument> sourceDocuments) throws IOException {
        XWPFTable targetTable = findSectionTable(targetDocument, STAFF_STATUS_SECTION);
        if (targetTable == null) {
            throw new IOException(Messages.format("error.template.section.not.found",
                    String.join(", ", STAFF_STATUS_SECTION.titles()), STAFF_STATUS_SECTION.tableIndex() + 1));
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
     * @param targetDocument  내용을 추가할 대상 문서
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

            Map<BigInteger, BigInteger> numberingIdMap = new LinkedHashMap<>();
            Map<String, String> styleIdMap = ensureSourceStylesAvailable(
                    targetDocument,
                    sourceDocument.document(),
                    numberingIdMap);
            for (IBodyElement bodyElement : sectionElements) {
                if (bodyElement instanceof XWPFParagraph paragraph
                        && isIssuesTitleParagraph(paragraph)) {
                    // 팀명을 별도 헤더로 추가하므로 원본 섹션 제목은 중복 출력하지 않는다.
                    continue;
                }
                appendBodyElement(targetDocument, sourceDocument.document(), bodyElement, numberingIdMap, styleIdMap);
            }

            appendedAnyContent = true;
        }
    }

    /**
     * 병합된 인력 현황 행에 합계 로직을 적용한다.
     * 데이터 행의 4열은 2열 + 3열로 계산하고, "합계" 행은 각 열의 데이터 합으로 재산정한다.
     *
     * @param mergedRows         병합된 행 목록 (인덱스 0: 카테고리 자리, 1: 2열, 2: 3열, 3: 4열)
     * @param accumulatorKeys    mergedRows 각 행에 대응하는 templateCategories 인덱스
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

    /**
     * 인력 현황 카테고리가 합계 행인지 확인한다.
     *
     * @param category 검사할 카테고리명
     * @return 정규화한 카테고리명이 합계이면 true
     */
    private static boolean isSumCategory(String category) {
        return "합계".equals(normalizeText(category));
    }

    /**
     * 인력 현황 숫자 값을 출력용 문자열로 변환한다.
     * 합계가 0이면 빈칸으로 보이도록 공백을 반환한다.
     *
     * @param value 출력할 숫자 값
     * @return 0이 아닌 경우 서식이 적용된 숫자 문자열, 0이면 빈 문자열
     */
    private static String formatStaffValue(BigDecimal value) {
        return value.compareTo(BigDecimal.ZERO) == 0 ? "" : NUMBER_FORMAT.format(value);
    }

    /**
     * 문자열을 숫자로 변환하고, 변환할 수 없는 값은 0으로 처리한다.
     * '-' 같은 표시 값도 합산 계산에서는 0으로 간주한다.
     *
     * @param value 변환할 문자열
     * @return 변환된 숫자 또는 0
     */
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
     * @param row                     현재 원본 데이터 행
     * @param fallbackRowIndex        카테고리 매칭 실패 시 사용할 기본 행 번호
     * @param templateCategoryIndexes 템플릿 카테고리와 행 번호 매핑
     * @param templateCategories      템플릿 카테고리 목록
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
     * @param document    탐색할 문서
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
     * @param document    탐색할 문서
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
        cursor.close();
        paragraph.createRun().addBreak();
    }

    /**
     * 결과 문서 끝에 단순 텍스트 문단을 추가한다.
     * 현재는 팀명을 구분 헤더처럼 보여 주는 용도로 사용한다.
     *
     * @param document 문단을 추가할 대상 문서
     * @param text     추가할 텍스트
     */
    private static void appendPlainParagraph(XWPFDocument document, String text) {
        XmlCursor cursor = document.getDocument().getBody().newCursor();
        cursor.toEndToken();
        XWPFParagraph paragraph = document.insertNewParagraph(cursor);
        cursor.close();
        paragraph.setAlignment(ParagraphAlignment.LEFT);
        XWPFRun run = paragraph.createRun();
        run.setFontSize(12);
        run.setText(text);
    }

    /**
     * 원본 문서의 문단 또는 표를 결과 문서 끝으로 복사한다.
     *
     * @param targetDocument 복사 대상을 붙여 넣을 결과 문서
     * @param bodyElement    복사할 원본 본문 요소
     */
    private static void appendBodyElement(
            XWPFDocument targetDocument,
            XWPFDocument sourceDocument,
            IBodyElement bodyElement,
            Map<BigInteger, BigInteger> numberingIdMap,
            Map<String, String> styleIdMap) {
        XmlCursor cursor = targetDocument.getDocument().getBody().newCursor();
        cursor.toEndToken();

        if (bodyElement instanceof XWPFParagraph paragraph) {
            XWPFParagraph appendedParagraph = targetDocument.insertNewParagraph(cursor);
            appendedParagraph.getCTP().set((CTP) paragraph.getCTP().copy());
            remapStyleReferences(appendedParagraph.getCTP(), styleIdMap);
        } else if (bodyElement instanceof XWPFTable table) {
            XWPFTable appendedTable = targetDocument.insertNewTbl(cursor);
            appendedTable.getCTTbl().set((CTTbl) table.getCTTbl().copy());
            remapStyleReferences(appendedTable.getCTTbl(), styleIdMap);
        }

        cursor.close();
    }

    /**
     * 원본 문서에서 사용하는 스타일을 결과 문서에 사용할 수 있도록 준비한다.
     * 동일 ID의 스타일 정의가 다르면 새 ID로 복사하고 참조 매핑을 반환한다.
     *
     * @param targetDocument 결과 문서
     * @param sourceDocument 원본 문서
     * @param numberingIdMap 원본 번호 ID와 결과 번호 ID 매핑
     * @return 원본 스타일 ID와 결과 스타일 ID 매핑
     */
    private static Map<String, String> ensureSourceStylesAvailable(
            XWPFDocument targetDocument,
            XWPFDocument sourceDocument,
            Map<BigInteger, BigInteger> numberingIdMap) {
        Map<String, String> styleIdMap = new LinkedHashMap<>();
        XWPFStyles sourceStyles = sourceDocument.getStyles();
        if (sourceStyles == null) {
            return styleIdMap;
        }

        XWPFStyles targetStyles = targetDocument.getStyles();
        if (targetStyles == null) {
            targetStyles = targetDocument.createStyles();
        }

        Set<String> reservedStyleIds = new HashSet<>();
        for (XWPFStyle sourceStyle : sourceStyles.getStyles()) {
            String styleId = sourceStyle.getStyleId();
            if (styleId == null || styleId.isBlank()) {
                continue;
            }

            XWPFStyle targetStyle = targetStyles.getStyle(styleId);
            if (targetStyle == null || hasSameStyleDefinition(sourceStyle, targetStyle)) {
                styleIdMap.put(styleId, styleId);
                reservedStyleIds.add(styleId);
                continue;
            }

            String copiedStyleId = buildCopiedStyleId(styleId, targetStyles, reservedStyleIds);
            styleIdMap.put(styleId, copiedStyleId);
            reservedStyleIds.add(copiedStyleId);
        }

        for (XWPFStyle sourceStyle : sourceStyles.getStyles()) {
            String styleId = sourceStyle.getStyleId();
            String targetStyleId = styleIdMap.get(styleId);
            if (targetStyleId == null || targetStyles.styleExist(targetStyleId)) {
                continue;
            }

            CTStyle copiedStyle = (CTStyle) sourceStyle.getCTStyle().copy();
            copiedStyle.setStyleId(targetStyleId);
            remapStyleReferences(copiedStyle, styleIdMap);
            remapNumberingReferences(targetDocument, sourceDocument, copiedStyle, numberingIdMap, styleIdMap);
            targetStyles.addStyle(new XWPFStyle(copiedStyle));
        }

        return styleIdMap;
    }

    /**
     * 두 스타일의 XML 정의가 같은지 비교한다.
     *
     * @param sourceStyle 원본 문서 스타일
     * @param targetStyle 결과 문서 스타일
     * @return 스타일 정의가 같으면 true
     */
    private static boolean hasSameStyleDefinition(XWPFStyle sourceStyle, XWPFStyle targetStyle) {
        return sourceStyle.getCTStyle().xmlText().equals(targetStyle.getCTStyle().xmlText());
    }

    /**
     * 결과 문서에서 충돌하지 않는 복사 스타일 ID를 만든다.
     *
     * @param sourceStyleId    원본 스타일 ID
     * @param targetStyles     결과 문서의 스타일 모음
     * @param reservedStyleIds 이번 복사 과정에서 이미 예약된 스타일 ID 목록
     * @return 충돌 없이 사용할 수 있는 스타일 ID
     */
    private static String buildCopiedStyleId(
            String sourceStyleId,
            XWPFStyles targetStyles,
            Set<String> reservedStyleIds) {
        String baseStyleId = "wm_" + sanitizeStyleId(sourceStyleId);
        int suffix = 1;
        String candidate = baseStyleId;
        while (targetStyles.styleExist(candidate) || reservedStyleIds.contains(candidate)) {
            suffix++;
            candidate = baseStyleId + "_" + suffix;
        }
        return candidate;
    }

    /**
     * Word 스타일 ID로 쓰기 어렵거나 너무 긴 문자를 안전한 형태로 정리한다.
     *
     * @param styleId 정리할 원본 스타일 ID
     * @return 영문, 숫자, 밑줄만 포함하는 최대 32자 스타일 ID
     */
    private static String sanitizeStyleId(String styleId) {
        String sanitized = styleId.replaceAll("[^A-Za-z0-9_]", "_");
        if (sanitized.isBlank()) {
            return "style";
        }
        return sanitized.length() > 32 ? sanitized.substring(0, 32) : sanitized;
    }

    /**
     * XML 객체 안의 스타일 참조 값을 결과 문서 스타일 ID로 재매핑한다.
     *
     * @param xmlObject  스타일 참조를 포함할 수 있는 XML 객체
     * @param styleIdMap 원본 스타일 ID와 결과 스타일 ID 매핑
     */
    private static void remapStyleReferences(
            org.apache.xmlbeans.XmlObject xmlObject,
            Map<String, String> styleIdMap) {
        if (styleIdMap.isEmpty()) {
            return;
        }

        try (XmlCursor cursor = xmlObject.newCursor()) {
            while (cursor.hasNextToken()) {
                cursor.toNextToken();
                javax.xml.namespace.QName name = cursor.getName();
                if (name == null
                        || !STYLE_REFERENCE_ELEMENTS.contains(name.getLocalPart())
                        || !WORDPROCESSINGML_NAMESPACE.equals(name.getNamespaceURI())) {
                    continue;
                }

                String sourceStyleId = cursor.getAttributeText(WORD_VAL_ATTRIBUTE);
                String targetStyleId = styleIdMap.get(sourceStyleId);
                if (targetStyleId != null) {
                    cursor.setAttributeText(WORD_VAL_ATTRIBUTE, targetStyleId);
                }
            }
        }
    }

    /**
     * XML 객체 안의 번호 매김 참조 값을 결과 문서 번호 ID로 재매핑한다.
     *
     * @param targetDocument 결과 문서
     * @param sourceDocument 원본 문서
     * @param xmlObject      번호 참조를 포함할 수 있는 XML 객체
     * @param numberingIdMap 원본 번호 ID와 결과 번호 ID 매핑
     * @param styleIdMap     원본 스타일 ID와 결과 스타일 ID 매핑
     */
    private static void remapNumberingReferences(
            XWPFDocument targetDocument,
            XWPFDocument sourceDocument,
            org.apache.xmlbeans.XmlObject xmlObject,
            Map<BigInteger, BigInteger> numberingIdMap,
            Map<String, String> styleIdMap) {
        try (XmlCursor cursor = xmlObject.newCursor()) {
            while (cursor.hasNextToken()) {
                cursor.toNextToken();
                javax.xml.namespace.QName name = cursor.getName();
                if (name == null
                        || !"numId".equals(name.getLocalPart())
                        || !WORDPROCESSINGML_NAMESPACE.equals(name.getNamespaceURI())) {
                    continue;
                }

                String sourceNumIdText = cursor.getAttributeText(WORD_VAL_ATTRIBUTE);
                if (sourceNumIdText == null || sourceNumIdText.isBlank()) {
                    continue;
                }

                BigInteger sourceNumId;
                try {
                    sourceNumId = new BigInteger(sourceNumIdText);
                } catch (NumberFormatException exception) {
                    continue;
                }

                BigInteger targetNumId = resolveTargetNumberingId(
                        targetDocument,
                        sourceDocument,
                        sourceNumId,
                        numberingIdMap,
                        styleIdMap);
                if (targetNumId != null) {
                    cursor.setAttributeText(WORD_VAL_ATTRIBUTE, targetNumId.toString());
                }
            }
        }
    }

    /**
     * 원본 문서의 번호 ID에 대응하는 결과 문서 번호 ID를 찾거나 새로 복사한다.
     * 새 번호 정의는 문서별로 독립되도록 시작 번호와 목록 식별자를 분리한다.
     *
     * @param targetDocument 결과 문서
     * @param sourceDocument 원본 문서
     * @param sourceNumId    원본 번호 ID
     * @param numberingIdMap 원본 번호 ID와 결과 번호 ID 매핑
     * @param styleIdMap     원본 스타일 ID와 결과 스타일 ID 매핑
     * @return 결과 문서에서 사용할 번호 ID
     */
    private static BigInteger resolveTargetNumberingId(
            XWPFDocument targetDocument,
            XWPFDocument sourceDocument,
            BigInteger sourceNumId,
            Map<BigInteger, BigInteger> numberingIdMap,
            Map<String, String> styleIdMap) {
        if (sourceNumId == null) {
            return null;
        }

        BigInteger mappedNumId = numberingIdMap.get(sourceNumId);
        if (mappedNumId != null) {
            return mappedNumId;
        }

        XWPFNumbering sourceNumbering = sourceDocument.getNumbering();
        XWPFNum sourceNum = sourceNumbering == null ? null : sourceNumbering.getNum(sourceNumId);
        if (sourceNum == null) {
            return sourceNumId;
        }

        BigInteger sourceAbstractNumId = sourceNum.getCTNum().getAbstractNumId().getVal();
        if (sourceAbstractNumId == null) {
            return sourceNumId;
        }

        XWPFAbstractNum sourceAbstractNum = sourceNumbering.getAbstractNum(sourceAbstractNumId);
        if (sourceAbstractNum == null) {
            return sourceNumId;
        }

        XWPFNumbering targetNumbering = targetDocument.getNumbering();
        if (targetNumbering == null) {
            targetNumbering = targetDocument.createNumbering();
        }

        CTAbstractNum copiedAbstractNum = (CTAbstractNum) sourceAbstractNum.getCTAbstractNum().copy();
        remapStyleReferences(copiedAbstractNum, styleIdMap);
        BigInteger newAbstractNumId = nextAvailableAbstractNumId(targetNumbering);
        copiedAbstractNum.setAbstractNumId(newAbstractNumId);
        // nsid/tmpl이 같으면 Word가 동일 목록으로 인식해 번호를 연속 할당하므로 고유값으로 교체한다.
        isolateAbstractNumListIds(copiedAbstractNum, newAbstractNumId);
        BigInteger targetAbstractNumId = targetNumbering.addAbstractNum(new XWPFAbstractNum(copiedAbstractNum));
        BigInteger targetNumId = nextAvailableNumId(targetNumbering);
        CTNum copiedNum = (CTNum) sourceNum.getCTNum().copy();
        copiedNum.setNumId(targetNumId);
        copiedNum.getAbstractNumId().setVal(targetAbstractNumId);
        // 기존 lvlOverride를 제거하고 모든 레벨에 startOverride=1을 추가해 팀 문서마다 번호가 1부터 시작되도록 강제한다.
        while (copiedNum.sizeOfLvlOverrideArray() > 0) {
            copiedNum.removeLvlOverride(0);
        }
        for (int level = 0; level <= 8; level++) {
            CTNumLvl lvlOverride = copiedNum.addNewLvlOverride();
            lvlOverride.setIlvl(BigInteger.valueOf(level));
            lvlOverride.addNewStartOverride().setVal(BigInteger.ONE);
        }
        targetNumbering.addNum(new XWPFNum(copiedNum, targetNumbering));
        numberingIdMap.put(sourceNumId, targetNumId);
        return targetNumId;
    }

    /**
     * Word가 서로 다른 원본 문서의 번호 목록을 같은 목록으로 묶지 않도록 목록 식별자를 고유하게 바꾼다.
     *
     * @param abstractNum   식별자를 바꿀 추상 번호 정의
     * @param abstractNumId 결과 문서에서 새로 할당한 추상 번호 ID
     */
    private static void isolateAbstractNumListIds(CTAbstractNum abstractNum, BigInteger abstractNumId) {
        String uniqueHex = String.format("%08X", abstractNumId.intValue() | 0x70000000);
        try (XmlCursor cursor = abstractNum.newCursor()) {
            while (cursor.hasNextToken()) {
                cursor.toNextToken();
                javax.xml.namespace.QName name = cursor.getName();
                if (name == null || !WORDPROCESSINGML_NAMESPACE.equals(name.getNamespaceURI())) {
                    continue;
                }
                String localPart = name.getLocalPart();
                if ("nsid".equals(localPart) || "tmpl".equals(localPart)) {
                    cursor.setAttributeText(WORD_VAL_ATTRIBUTE, uniqueHex);
                }
            }
        }
    }

    /**
     * 결과 문서 번호 모음에서 사용 가능한 다음 추상 번호 ID를 찾는다.
     *
     * @param numbering 결과 문서의 번호 모음
     * @return 아직 사용되지 않은 추상 번호 ID
     */
    private static BigInteger nextAvailableAbstractNumId(XWPFNumbering numbering) {
        BigInteger abstractNumId = BigInteger.ZERO;
        while (numbering.getAbstractNum(abstractNumId) != null) {
            abstractNumId = abstractNumId.add(BigInteger.ONE);
        }
        return abstractNumId;
    }

    /**
     * 결과 문서 번호 모음에서 사용 가능한 다음 번호 ID를 찾는다.
     *
     * @param numbering 결과 문서의 번호 모음
     * @return 아직 사용되지 않은 번호 ID
     */
    private static BigInteger nextAvailableNumId(XWPFNumbering numbering) {
        BigInteger numId = BigInteger.ONE;
        while (numbering.getNum(numId) != null) {
            numId = numId.add(BigInteger.ONE);
        }
        return numId;
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
     * @param text       검사할 원본 문자열
     * @param candidates 포함 여부를 확인할 후보 문자열 목록
     * @return 하나라도 포함되면 true
     */
    private static boolean containsAny(String text, List<String> candidates) {
        return candidates.stream().anyMatch(text::contains);
    }

    /**
     * 행 전체가 영업 섹션의 고정 헤더 행인지 판별한다.
     *
     * @param row                      비교할 행 값 목록
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
     * @param table             원본 영업 현황 표
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
     * @param rows           후보 영업 행 목록
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
     * @param table             원본 표
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
     * @param table       템플릿 표
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
     * @param rows  기록할 데이터 행 목록
     * @throws IOException 템플릿 표 구조가 비정상인 경우
     */
    private static void rewriteTableRows(XWPFTable table, List<List<String>> rows) throws IOException {
        if (table == null || table.getNumberOfRows() == 0) {
            throw new IOException(Messages.get("error.template.table.empty"));
        }

        int templateRowIndex = table.getNumberOfRows() > 1 ? 1 : 0;
        XWPFTableRow templateRow = table.getRow(templateRowIndex);
        if (templateRow == null) {
            throw new IOException(Messages.get("error.template.table.no.row"));
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
     * @param table              대상 표
     * @param rows               병합된 데이터 행
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
     * @param table                 대상 표
     * @param rows                  병합된 영업 데이터
     * @param preservedColumnValues 유지할 템플릿 열 값
     * @param preservedColumnIndex  유지할 열 인덱스
     * @throws IOException 템플릿 표 구조가 비정상인 경우
     */
    private static void rewriteSalesTableRowsPreservingColumn(
            XWPFTable table,
            List<SalesRowData> rows,
            List<String> preservedColumnValues,
            int preservedColumnIndex) throws IOException {
        if (table == null || table.getNumberOfRows() == 0) {
            throw new IOException(Messages.get("error.template.table.empty"));
        }

        int templateRowIndex = table.getNumberOfRows() > 1 ? 1 : 0;
        XWPFTableRow templateRow = table.getRow(templateRowIndex);
        if (templateRow == null) {
            throw new IOException(Messages.get("error.template.table.no.row"));
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
     * @param targetRow        복사 결과를 기록할 행
     * @param sourceRow        복사할 원본 영업 행
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
     * @param table       행을 추가할 표
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
     * @param row       셀을 가져올 대상 행
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
     * @param targetRow   서식을 적용할 대상 행
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
     * @param targetCell   서식을 적용할 대상 셀
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
     * @param row    값을 기록할 대상 행
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

    /**
     * 인력 현황 합계 행의 숫자 셀을 굵게 표시한다.
     *
     * @param table        인력 현황 표
     * @param dataRowCount 템플릿 기준 데이터 행 수
     */
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
     * @param cell  값을 기록할 셀
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
     * 셀 안의 모든 문단 run에 굵게 스타일을 적용하거나 해제한다.
     *
     * @param cell 스타일을 바꿀 셀
     * @param bold {@code true}이면 굵게, {@code false}이면 굵게 해제
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

    /**
     * 셀 텍스트 비교와 출력에 쓰기 좋도록 개행/공백을 정리한다.
     *
     * @param value 정리할 원본 문자열
     * @return 앞뒤 공백과 carriage return이 제거된 문자열
     */
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
     * 레이아웃 차이를 무시하고 비교할 수 있도록 모든 공백을 제거한다.
     *
     * @param value 정규화할 원본 문자열
     * @return 모든 공백이 제거된 문자열
     */
    private static String normalizeText(String value) {
        return value == null ? "" : value.replaceAll("\\s+", "");
    }

    /**
     * 숫자 문자열을 {@link BigDecimal}로 변환한다.
     *
     * @param value 변환할 문자열
     * @return 변환된 숫자, 숫자가 아니면 {@code null}
     */
    private static BigDecimal parseNumber(String value) {
        String normalized = value.replace(",", "").trim();
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

    /**
     * 원본 문서 경로와 열린 Word 문서 객체를 함께 보관한다.
     *
     * @param path     원본 문서 경로
     * @param document 열린 Word 문서
     */
    private record SourceDocument(Path path, XWPFDocument document) {
    }

    /**
     * 영업 현황 행의 정리된 셀 값과 원본 셀 서식을 함께 보관한다.
     *
     * @param values      정리된 셀 값 목록
     * @param sourceCells 원본 셀 목록
     */
    private record SalesRowData(List<String> values, List<XWPFTableCell> sourceCells) {
    }
}
