import java.awt.BorderLayout;
import java.awt.Dimension;
import java.awt.GridLayout;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.Stream;

import javax.swing.DefaultListModel;
import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JList;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.ListSelectionModel;
import javax.swing.SwingUtilities;

public class WeeklyMergeApp extends JFrame {
    // 설정 파일이 없거나 값이 잘못된 경우 현재 폴더를 기본 경로로 사용한다.
    private static final String CONFIG_FILE_NAME = "app.properties";
    private static final String REPORT_DOC_PATH_KEY = "report.doc.path";
    private static final String REPORT_DOC_OUT_PATH_KEY = "report.doc.out.path";
    private static final Path DEFAULT_REPORT_DOC_PATH = Paths.get(System.getProperty("user.dir"));
    private static final Path REPORT_DOC_PATH = loadConfiguredAbsolutePath(REPORT_DOC_PATH_KEY,
            DEFAULT_REPORT_DOC_PATH);
    private static final Path REPORT_DOC_OUT_PATH = loadConfiguredAbsolutePath(REPORT_DOC_OUT_PATH_KEY,
            DEFAULT_REPORT_DOC_PATH);

    private final DefaultListModel<String> listModel = new DefaultListModel<>();
    private final JList<String> fileList = new JList<>(listModel);

    /**
     * 주간 보고서 병합용 Swing 메인 창을 생성한다.
     * UI를 초기화한 뒤 바로 원본 문서 목록을 읽어 화면에 표시한다.
     */
    public WeeklyMergeApp() {
        super("Weekly Merge Report Viewer");
        initializeUi();
        loadDocFiles();
    }

    /**
     * 설정 파일에서 절대 경로 값을 읽어 오고, 유효하지 않으면 기본 경로를 반환한다.
     *
     * @param key         읽어올 설정 키
     * @param defaultPath 설정이 없거나 잘못된 경우 사용할 기본 경로
     * @return 정규화된 절대 경로 또는 기본 경로
     */
    private static Path loadConfiguredAbsolutePath(String key, Path defaultPath) {
        Path configPath = Paths.get(CONFIG_FILE_NAME);

        // 설정 파일이 없으면 사용자가 별도 설정을 하지 않은 것으로 보고 기본 경로를 사용한다.
        if (!Files.exists(configPath)) {
            return defaultPath;
        }

        try {
            String configuredPath = readConfigValue(configPath, key);
            if (configuredPath.isEmpty()) {
                return defaultPath;
            }

            Path resolvedPath = Paths.get(configuredPath);
            // 상대 경로는 실행 위치에 따라 달라질 수 있어 허용하지 않고 기본 경로로 되돌린다.
            if (!resolvedPath.isAbsolute()) {
                return defaultPath;
            }

            return resolvedPath.normalize();
        } catch (IOException exception) {
            // 설정 파일을 읽지 못해도 앱은 계속 실행될 수 있도록 기본 경로를 사용한다.
            return defaultPath;
        }
    }

    /**
     * 설정 파일에서 특정 키의 값을 직접 찾아 반환한다.
     *
     * @param configPath 설정 파일 경로
     * @param key        찾을 설정 키
     * @return 키에 대응하는 값, 없으면 빈 문자열
     * @throws IOException 설정 파일을 읽는 중 오류가 발생한 경우
     */
    private static String readConfigValue(Path configPath, String key) throws IOException {
        List<String> lines = Files.readAllLines(configPath);

        for (String line : lines) {
            String trimmedLine = line.trim();
            // 공백 줄과 주석 줄은 실제 설정 데이터가 아니므로 건너뛴다.
            if (trimmedLine.isEmpty() || trimmedLine.startsWith("#") || trimmedLine.startsWith("!")) {
                continue;
            }

            int separatorIndex = trimmedLine.indexOf('=');
            // key=value 형식이 아닌 줄은 무시한다.
            if (separatorIndex < 0) {
                continue;
            }

            String currentKey = trimmedLine.substring(0, separatorIndex).trim();
            // 요청한 키와 일치하는 항목만 반환한다.
            if (!key.equals(currentKey)) {
                continue;
            }

            return trimmedLine.substring(separatorIndex + 1).trim();
        }

        return "";
    }

    /**
     * 화면 레이아웃과 버튼, 목록 같은 Swing 컴포넌트를 초기화한다.
     * 상단에는 경로 정보, 중앙에는 파일 목록, 하단에는 병합 버튼이 배치된다.
     */
    private void initializeUi() {
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setLayout(new BorderLayout(12, 12));

        JPanel infoPanel = new JPanel(new GridLayout(2, 1));
        infoPanel.add(new JLabel("Source Path: " + REPORT_DOC_PATH.toAbsolutePath()));
        infoPanel.add(new JLabel("Output Path: " + REPORT_DOC_OUT_PATH.toAbsolutePath()));

        fileList.setSelectionMode(ListSelectionModel.MULTIPLE_INTERVAL_SELECTION);

        JScrollPane scrollPane = new JScrollPane(fileList);
        scrollPane.setPreferredSize(new Dimension(720, 360));

        JButton refreshButton = new JButton("폴더 다시 읽기");
        refreshButton.addActionListener(event -> loadDocFiles());

        JButton mergeReportsButton = new JButton("주간보고 병합");
        mergeReportsButton.addActionListener(event -> mergeSelectedFiles());

        JPanel bottomPanel = new JPanel();
        bottomPanel.add(refreshButton);
        bottomPanel.add(mergeReportsButton);

        add(infoPanel, BorderLayout.NORTH);
        add(scrollPane, BorderLayout.CENTER);
        add(bottomPanel, BorderLayout.SOUTH);

        pack();
        setLocationRelativeTo(null);
    }

    /**
     * 원본 폴더에서 병합 가능한 `.docx` 파일 목록을 읽어 리스트에 표시한다.
     * 템플릿 파일과 이미 생성된 결과 파일은 목록에서 제외한다.
     */
    private void loadDocFiles() {
        listModel.clear();

        // 원본 폴더가 없으면 파일 선택 자체가 불가능하므로 경고 후 중단한다.
        if (!Files.isDirectory(REPORT_DOC_PATH)) {
            JOptionPane.showMessageDialog(
                    this,
                    "Source folder was not found.\nPath: " + REPORT_DOC_PATH.toAbsolutePath(),
                    "Folder Not Found",
                    JOptionPane.WARNING_MESSAGE);
            return;
        }

        try (Stream<Path> paths = Files.list(REPORT_DOC_PATH)) {
            List<String> docFiles = paths
                    .filter(Files::isRegularFile)
                    .map(Path::getFileName)
                    .map(Path::toString)
                    // 병합 대상은 Word 문서만 허용한다.
                    .filter(name -> name.toLowerCase().endsWith(".docx"))
                    // 템플릿 문서는 입력 문서가 아니라 기준 문서이므로 제외한다.
                    .filter(name -> !name.equalsIgnoreCase(ReportMerger.TEMPLATE_FILE_NAME))
                    // 이미 생성된 결과 파일은 다시 병합 대상이 되지 않도록 제외한다.
                    .filter(name -> !name.matches("\\d{8}_.+\\.docx"))
                    .sorted(String.CASE_INSENSITIVE_ORDER)
                    .collect(Collectors.toList());

            if (docFiles.isEmpty()) {
                // 선택 가능한 파일이 없을 때는 안내 문구를 보여 주고 목록 선택을 비활성화한다.
                listModel.addElement("No source .docx files found.");
                fileList.setEnabled(false);
                return;
            }

            // 정상적으로 파일을 찾은 경우 목록 선택을 활성화하고 파일명을 표시한다.
            fileList.setEnabled(true);
            docFiles.forEach(listModel::addElement);
        } catch (Exception exception) {
            JOptionPane.showMessageDialog(
                    this,
                    "An error occurred while reading the file list.\n" + exception.getMessage(),
                    "Error",
                    JOptionPane.ERROR_MESSAGE);
        }
    }

    /**
     * 사용자가 선택한 파일을 실제 병합 로직에 전달하고 결과를 메시지로 안내한다.
     * 성공 시 생성된 파일 경로를 보여 주고, 실패 시 원인을 요약해 표시한다.
     */
    private void mergeSelectedFiles() {
        List<String> selectedFiles = fileList.getSelectedValuesList();

        // 선택된 파일이 없으면 병합할 대상이 없으므로 안내 메시지만 표시한다.
        if (selectedFiles.isEmpty()) {
            JOptionPane.showMessageDialog(
                    this,
                    "Select one or more report files to merge.",
                    "Notice",
                    JOptionPane.INFORMATION_MESSAGE);
            return;
        }

        // 출력 폴더가 없으면 결과 파일을 저장할 수 없으므로 병합을 시작하지 않는다.
        if (!Files.isDirectory(REPORT_DOC_OUT_PATH)) {
            JOptionPane.showMessageDialog(
                    this,
                    "Output folder was not found.\nPath: " + REPORT_DOC_OUT_PATH.toAbsolutePath(),
                    "Folder Not Found",
                    JOptionPane.WARNING_MESSAGE);
            return;
        }

        List<Path> sourceFiles = selectedFiles.stream()
                .map(name -> REPORT_DOC_PATH.resolve(name))
                .collect(Collectors.toList());

        try {
            // 선택한 파일 목록을 절대 경로로 변환해 병합 엔진에 전달한다.
            Path outputFile = ReportMerger.mergeReports(sourceFiles, REPORT_DOC_OUT_PATH);
            JOptionPane.showMessageDialog(
                    this,
                    "Merged report created:\n" + outputFile.toAbsolutePath(),
                    "Merge Complete",
                    JOptionPane.INFORMATION_MESSAGE);
        } catch (Exception exception) {
            JOptionPane.showMessageDialog(
                    this,
                    "Failed to merge reports.\n" + buildErrorMessage(exception),
                    "Merge Failed",
                    JOptionPane.ERROR_MESSAGE);
        }
    }

    /**
     * 예외 객체에서 사용자에게 보여 줄 수 있는 오류 메시지를 재귀적으로 구성한다.
     *
     * @param throwable 메시지를 추출할 예외
     * @return 사용자 안내용 오류 메시지
     */
    private static String buildErrorMessage(Throwable throwable) {
        if (throwable == null) {
            return "Unknown error";
        }

        String message = throwable.getMessage();
        // 현재 예외에 의미 있는 메시지가 있으면 가장 먼저 그 값을 사용한다.
        if (message != null && !message.isBlank()) {
            return message;
        }

        Throwable cause = throwable.getCause();
        // 현재 예외 메시지가 비어 있으면 내부 원인 예외를 따라가며 설명을 찾는다.
        if (cause != null && cause != throwable) {
            String causeMessage = buildErrorMessage(cause);
            if (causeMessage != null && !causeMessage.isBlank()) {
                return throwable.getClass().getSimpleName() + ": " + causeMessage;
            }
        }

        return throwable.getClass().getSimpleName();
    }

    /**
     * Swing 이벤트 디스패치 스레드에서 애플리케이션 창을 실행한다.
     *
     * @param args 실행 인수, 현재는 사용하지 않는다
     */
    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            // Swing UI는 EDT에서 생성해야 하므로 invokeLater 내부에서 창을 띄운다.
            WeeklyMergeApp app = new WeeklyMergeApp();
            app.setVisible(true);
        });
    }
}
