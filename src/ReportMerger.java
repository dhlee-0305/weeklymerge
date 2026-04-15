import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Path;
import java.text.DecimalFormat;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
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

public final class ReportMerger {
    public static final String TEMPLATE_FILE_NAME = "\uacbd\uc601\uc804\ub7b5\ud68c\uc758_template.docx";
    private static final int STAFF_CATEGORY_COLUMN_INDEX = 0;
    private static final int SALES_EXCLUDED_COLUMN_INDEX = 0;

    private static final String OUTPUT_FILE_SUFFIX = "_\uacbd\uc601\uc804\ub7b5\ud68c\uc758.docx";
    private static final SectionSpec PROJECT_STATUS_SECTION =
            new SectionSpec(List.of("\ud504\ub85c\uc81d\ud2b8 \uc9c4\ud589 \ud604\ud669", "\ud504\ub85c\uc81d\ud2b8 \ud604\ud669"), 0, 1);
    private static final SectionSpec STAFF_STATUS_SECTION =
            new SectionSpec(List.of("\uc778\ub825 \uc6b4\uc6a9 \ud604\ud669", "\uc778\ub825 \ud604\ud669"), 1, 1);
    private static final SectionSpec SALES_STATUS_SECTION =
            new SectionSpec(
                    List.of("\uac70\ub798\ucc98 \uc601\uc5c5/\ub3d9\ud5a5 \uc815\ubcf4",
                            "\uac70\ub798\ucc98 \uc601\uc5c5 \ub3d9\ud5a5 \uc815\ubcf4",
                            "\uc601\uc5c5/\ub3d9\ud5a5 \uc815\ubcf4"),
                    2,
                    1);
    private static final List<String> SALES_FIXED_HEADER_ROW =
            List.of("\uad6c\ubd84", "\uace0\uac1d\uc0ac/\ubd80\uc11c", "\uc8fc\uc694 \uc815\ubcf4", "\ube44\uace0");
    private static final List<String> SALES_FIXED_HEADER_ROW_WITHOUT_FIRST_COLUMN =
            List.of("\uace0\uac1d\uc0ac/\ubd80\uc11c", "\uc8fc\uc694 \uc815\ubcf4", "\ube44\uace0");
    private static final SectionSpec ISSUES_SECTION =
            new SectionSpec(
                    List.of("\uc8fc\uc694 \uc5c5\ubb34 \uc0ac\ud56d",
                            "\uc8fc\uc694\uc5c5\ubb34\uc0ac\ud56d",
                            "\uc8fc\uc694 \uc5c5\ubb34",
                            "\uc8fc\uc694\uc5c5\ubb34",
                            "\uc8fc\uc694 \uc774\uc288",
                            "\uc774\uc288 \uc0ac\ud56d"),
                    3,
                    1);
    private static final Pattern NUMERIC_PATTERN = Pattern.compile("[-+]?\\d+(\\.\\d+)?");
    private static final DecimalFormat NUMBER_FORMAT = new DecimalFormat("0.################");

    private ReportMerger() {
    }

    public static Path mergeReports(List<Path> sourceFiles, Path outputDirectory) throws IOException {
        Path templatePath = outputDirectory.resolve(TEMPLATE_FILE_NAME);
        if (!Files.exists(templatePath)) {
            throw new IOException("Template file not found: " + templatePath.toAbsolutePath());
        }

        try (InputStream templateStream = Files.newInputStream(templatePath);
                XWPFDocument targetDocument = new XWPFDocument(templateStream)) {

            List<SourceDocument> sourceDocuments = new ArrayList<>();
            try {
                for (Path sourceFile : sourceFiles) {
                    sourceDocuments.add(new SourceDocument(sourceFile, openDocument(sourceFile)));
                }

                mergeAppendSection(targetDocument, sourceDocuments, PROJECT_STATUS_SECTION);
                mergeStaffSection(targetDocument, sourceDocuments);
                mergeAppendSection(targetDocument, sourceDocuments, SALES_STATUS_SECTION);
                appendIssuesSectionToDocumentEnd(targetDocument, sourceDocuments);

                Path outputFile = outputDirectory.resolve(
                        LocalDate.now().format(DateTimeFormatter.BASIC_ISO_DATE) + OUTPUT_FILE_SUFFIX);
                try (OutputStream outputStream = Files.newOutputStream(outputFile)) {
                    targetDocument.write(outputStream);
                }
                return outputFile;
            } finally {
                for (SourceDocument sourceDocument : sourceDocuments) {
                    sourceDocument.document().close();
                }
            }
        }
    }

    private static XWPFDocument openDocument(Path sourceFile) throws IOException {
        try (InputStream inputStream = Files.newInputStream(sourceFile)) {
            return new XWPFDocument(inputStream);
        }
    }

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
            templateFirstColumnValues = extractTemplateColumnValues(targetTable, SALES_EXCLUDED_COLUMN_INDEX);
        }

        for (SourceDocument sourceDocument : sourceDocuments) {
            XWPFTable sourceTable = findSectionTable(sourceDocument.document(), sectionSpec);
            if (sourceTable == null) {
                continue;
            }
            List<List<String>> sectionRows = extractDataRows(sourceTable, sectionSpec.dataStartRowIndex());
            if (sectionSpec == SALES_STATUS_SECTION) {
                sectionRows = sanitizeSalesRows(sectionRows);
            }
            mergedRows.addAll(sectionRows);
        }

        if (sectionSpec == SALES_STATUS_SECTION) {
            rewriteTableRowsPreservingColumn(
                    targetTable,
                    mergedRows,
                    templateFirstColumnValues,
                    SALES_EXCLUDED_COLUMN_INDEX);
            return;
        }

        rewriteTableRows(targetTable, mergedRows);
    }

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
                int targetRowIndex = resolveStaffTargetRowIndex(row, rowIndex, templateCategoryIndexes, templateCategories);
                if (targetRowIndex < 0) {
                    continue;
                }

                Map<Integer, CellAccumulator> rowAccumulator = accumulators.computeIfAbsent(targetRowIndex,
                        key -> new LinkedHashMap<>());
                for (int colIndex = 0; colIndex < row.size(); colIndex++) {
                    if (colIndex == STAFF_CATEGORY_COLUMN_INDEX) {
                        continue;
                    }

                    rowAccumulator
                            .computeIfAbsent(colIndex, key -> new CellAccumulator())
                            .add(row.get(colIndex));
                }
            }
        }

        List<List<String>> mergedRows = new ArrayList<>();
        for (Map<Integer, CellAccumulator> rowAccumulator : accumulators.values()) {
            int maxColumnIndex = rowAccumulator.keySet().stream().mapToInt(Integer::intValue).max().orElse(-1);
            List<String> mergedRow = new ArrayList<>();
            for (int colIndex = 0; colIndex <= maxColumnIndex; colIndex++) {
                CellAccumulator accumulator = rowAccumulator.get(colIndex);
                mergedRow.add(accumulator == null ? "" : accumulator.getMergedValue());
            }
            mergedRows.add(mergedRow);
        }

        rewriteStaffTableRows(targetTable, mergedRows, templateCategories);
    }

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
                appendBlankParagraph(targetDocument);
            }

            appendPlainParagraph(
                    targetDocument,
                    "[" + extractTeamName(sourceDocument.path()) + "]");

            for (IBodyElement bodyElement : sectionElements) {
                if (bodyElement instanceof XWPFParagraph paragraph
                        && isIssuesTitleParagraph(paragraph)) {
                    continue;
                }
                appendBodyElement(targetDocument, bodyElement);
            }

            appendedAnyContent = true;
        }
    }

    private static Map<String, Integer> buildTemplateCategoryIndexes(List<String> templateCategories) {
        Map<String, Integer> indexes = new LinkedHashMap<>();
        for (int rowIndex = 0; rowIndex < templateCategories.size(); rowIndex++) {
            String normalizedCategory = normalizeText(templateCategories.get(rowIndex));
            if (!normalizedCategory.isEmpty()) {
                indexes.putIfAbsent(normalizedCategory, rowIndex);
            }
        }
        return indexes;
    }

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

        return fallbackRowIndex < templateCategories.size() ? fallbackRowIndex : -1;
    }

    private static XWPFTable findSectionTable(XWPFDocument document, SectionSpec sectionSpec) {
        List<String> normalizedTitles = sectionSpec.titles().stream()
                .map(ReportMerger::normalizeText)
                .toList();
        boolean nextTableMatches = false;

        for (IBodyElement bodyElement : document.getBodyElements()) {
            if (bodyElement instanceof XWPFParagraph) {
                String paragraphText = normalizeText(((XWPFParagraph) bodyElement).getText());
                if (containsAny(paragraphText, normalizedTitles)) {
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

    private static boolean isIssuesTitleParagraph(XWPFParagraph paragraph) {
        String normalizedText = normalizeText(cleanCellText(paragraph.getText()));
        List<String> normalizedTitles = ISSUES_SECTION.titles().stream()
                .map(ReportMerger::normalizeText)
                .toList();
        return containsAny(normalizedText, normalizedTitles);
    }

    private static void appendBlankParagraph(XWPFDocument document) {
        XmlCursor cursor = document.getDocument().getBody().newCursor();
        cursor.toEndToken();
        XWPFParagraph paragraph = document.insertNewParagraph(cursor);
        cursor.dispose();
        paragraph.createRun().addBreak();
    }

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


    private static boolean containsAny(String text, List<String> candidates) {
        return candidates.stream().anyMatch(text::contains);
    }

    private static List<List<String>> filterOutFixedRows(List<List<String>> rows, List<String> fixedRowValues) {
        List<String> normalizedFixedRowValues = fixedRowValues.stream()
                .map(ReportMerger::normalizeText)
                .toList();
        List<List<String>> filteredRows = new ArrayList<>();

        for (List<String> row : rows) {
            if (!matchesFixedRow(row, normalizedFixedRowValues)) {
                filteredRows.add(row);
            }
        }

        return filteredRows;
    }

    private static boolean matchesFixedRow(List<String> row, List<String> normalizedFixedRowValues) {
        String normalizedJoinedRow = row.stream()
                .map(ReportMerger::normalizeText)
                .reduce("", String::concat);

        for (String fixedValue : normalizedFixedRowValues) {
            if (!normalizedJoinedRow.contains(fixedValue)) {
                return false;
            }
        }

        return !normalizedJoinedRow.isEmpty();
    }

    private static List<List<String>> removeColumn(List<List<String>> rows, int columnIndex) {
        List<List<String>> adjustedRows = new ArrayList<>();
        for (List<String> row : rows) {
            List<String> adjustedRow = new ArrayList<>();
            for (int index = 0; index < row.size(); index++) {
                if (index != columnIndex) {
                    adjustedRow.add(row.get(index));
                }
            }
            adjustedRows.add(adjustedRow);
        }
        return adjustedRows;
    }

    private static List<List<String>> sanitizeSalesRows(List<List<String>> rows) {
        List<List<String>> sanitizedRows = dropFirstRow(rows);
        sanitizedRows = filterOutFixedRows(sanitizedRows, SALES_FIXED_HEADER_ROW);
        sanitizedRows = removeColumn(sanitizedRows, SALES_EXCLUDED_COLUMN_INDEX);
        sanitizedRows = filterOutFixedRows(sanitizedRows, SALES_FIXED_HEADER_ROW_WITHOUT_FIRST_COLUMN);
        return sanitizedRows;
    }

    private static List<List<String>> dropFirstRow(List<List<String>> rows) {
        if (rows.isEmpty()) {
            return rows;
        }
        return new ArrayList<>(rows.subList(1, rows.size()));
    }

    private record SectionSpec(List<String> titles, int tableIndex, int dataStartRowIndex) {
    }

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

    private static List<String> extractCellValues(XWPFTableRow row) {
        List<String> cellValues = new ArrayList<>();
        for (XWPFTableCell cell : row.getTableCells()) {
            cellValues.add(cleanCellText(cell.getText()));
        }
        return cellValues;
    }

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
            setCellText(row.getCell(STAFF_CATEGORY_COLUMN_INDEX), templateCategories.get(rowOffset));
        }
    }

    private static void rewriteTableRowsPreservingColumn(
            XWPFTable table,
            List<List<String>> rows,
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
            if (templateRow.getTableCells().size() > preservedColumnIndex && !preservedColumnValues.isEmpty()) {
                setCellText(templateRow.getCell(preservedColumnIndex), preservedColumnValues.get(0));
            }
            return;
        }

        clearRow(templateRow);
        setRowValuesWithOffset(templateRow, rows.get(0), preservedColumnIndex + 1);
        if (templateRow.getTableCells().size() > preservedColumnIndex && !preservedColumnValues.isEmpty()) {
            setCellText(templateRow.getCell(preservedColumnIndex), preservedColumnValues.get(0));
        }

        for (int rowIndex = 1; rowIndex < rows.size(); rowIndex++) {
            XWPFTableRow row = createStyledRow(table, templateRow);
            clearRow(row);
            setRowValuesWithOffset(row, rows.get(rowIndex), preservedColumnIndex + 1);
            if (row.getTableCells().size() > preservedColumnIndex && rowIndex < preservedColumnValues.size()) {
                setCellText(row.getCell(preservedColumnIndex), preservedColumnValues.get(rowIndex));
            }
        }
    }

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

    private static XWPFTableCell getOrCreateCell(XWPFTableRow row, int cellIndex) {
        while (row.getTableCells().size() <= cellIndex) {
            row.addNewTableCell();
        }
        return row.getCell(cellIndex);
    }

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

    private static void clearRow(XWPFTableRow row) {
        for (XWPFTableCell cell : row.getTableCells()) {
            clearCellText(cell);
        }
    }

    private static void setRowValues(XWPFTableRow row, List<String> values) {
        int limit = Math.max(row.getTableCells().size(), values.size());
        for (int cellIndex = 0; cellIndex < limit; cellIndex++) {
            XWPFTableCell cell = getOrCreateCell(row, cellIndex);
            String value = cellIndex < values.size() ? values.get(cellIndex) : "";
            setCellText(cell, value);
        }
    }

    private static void setRowValuesWithOffset(XWPFTableRow row, List<String> values, int startColumnIndex) {
        int limit = Math.max(row.getTableCells().size(), values.size() + startColumnIndex);
        for (int cellIndex = startColumnIndex; cellIndex < limit; cellIndex++) {
            XWPFTableCell cell = getOrCreateCell(row, cellIndex);
            int valueIndex = cellIndex - startColumnIndex;
            String value = valueIndex < values.size() ? values.get(valueIndex) : "";
            setCellText(cell, value);
        }
    }

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
                run.addBreak();
            }
            run.setText(lines[lineIndex]);
        }
    }

    private static String cleanCellText(String value) {
        return value == null ? "" : value.replace("\r", "").trim();
    }

    private static boolean isEmptyRow(List<String> row) {
        for (String value : row) {
            if (!value.isBlank()) {
                return false;
            }
        }
        return true;
    }

    private static String normalizeText(String value) {
        return value == null ? "" : value.replaceAll("\\s+", "");
    }

    private static BigDecimal parseNumber(String value) {
        String normalized = value.replace(",", "").trim();
        if (!NUMERIC_PATTERN.matcher(normalized).matches()) {
            return null;
        }
        return new BigDecimal(normalized);
    }

    private static final class CellAccumulator {
        private final List<String> values = new ArrayList<>();
        private BigDecimal numericSum = BigDecimal.ZERO;
        private boolean hasValue;
        private boolean allNumeric = true;
        CellAccumulator() {
        }

        void add(String value) {
            String cleanedValue = cleanCellText(value);
            if (cleanedValue.isEmpty()) {
                return;
            }

            hasValue = true;
            values.add(cleanedValue);

            BigDecimal parsedNumber = parseNumber(cleanedValue);
            if (parsedNumber == null) {
                allNumeric = false;
                return;
            }

            numericSum = numericSum.add(parsedNumber);
        }

        String getMergedValue() {
            if (!hasValue) {
                return "";
            }

            if (allNumeric) {
                return NUMBER_FORMAT.format(numericSum);
            }

            return String.join("\n", values);
        }
    }

    private record SourceDocument(Path path, XWPFDocument document) {
    }
}
