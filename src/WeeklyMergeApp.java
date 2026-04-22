import java.awt.BorderLayout;
import java.awt.Dimension;
import java.awt.FlowLayout;
import java.awt.GridLayout;
import java.beans.PropertyChangeEvent;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.Stream;

import javax.swing.DefaultListModel;
import javax.swing.JButton;
import javax.swing.JDialog;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JList;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JProgressBar;
import javax.swing.JScrollPane;
import javax.swing.ListSelectionModel;
import javax.swing.SwingUtilities;
import javax.swing.SwingWorker;

public class WeeklyMergeApp extends JFrame {
    // 설정 파일이 없거나 값이 유효하지 않으면 현재 작업 폴더를 기본 입력 경로로 사용한다.
    private static final String CONFIG_FILE_NAME = "app.properties";
    private static final String REPORT_DOC_PATH_KEY = "report.doc.path";
    private static final String REPORT_DOC_OUT_PATH_KEY = "report.doc.out.path";
    private static final Path DEFAULT_REPORT_DOC_PATH = Paths.get(System.getProperty("user.dir"));
    private static final Path REPORT_DOC_PATH = loadConfiguredAbsolutePath(REPORT_DOC_PATH_KEY,
            DEFAULT_REPORT_DOC_PATH);
    // 결과 문서는 기본적으로 실행 폴더 아래 `output` 폴더에 생성된다.
    private static final Path DEFAULT_REPORT_DOC_OUT_PATH = Paths.get(System.getProperty("user.dir"), "output");
    private static final Path REPORT_DOC_OUT_PATH = loadConfiguredAbsolutePath(
            REPORT_DOC_OUT_PATH_KEY,
            DEFAULT_REPORT_DOC_OUT_PATH);

    private final DefaultListModel<String> listModel = new DefaultListModel<>();
    private final JList<String> fileList = new JList<>(listModel);
    private final JButton refreshButton = new JButton("Refresh Files");
    private final JButton mergeReportsButton = new JButton("Merge Reports");

    /**
     * 메인 애플리케이션 창을 생성하고 기본 UI와 파일 목록을 초기화한다.
     * 창이 표시되면 사용자는 즉시 병합 가능한 보고서 파일 목록을 확인할 수 있다.
     */
    public WeeklyMergeApp() {
        super("Weekly Report Merge");
        initializeUi();
        loadDocFiles();
    }

    /**
     * 설정 파일에서 절대 경로를 읽어 오고, 사용할 수 없으면 기본 경로를 반환한다.
     *
     * @param key 설정 파일에서 읽을 키 이름
     * @param defaultPath 설정값이 없거나 잘못되었을 때 사용할 기본 경로
     * @return 정규화된 절대 경로 또는 기본 경로
     */
    private static Path loadConfiguredAbsolutePath(String key, Path defaultPath) {
        Path configPath = Paths.get(CONFIG_FILE_NAME);
        // 설정 파일이 없으면 별도 구성 없이도 앱이 실행되도록 기본 경로를 그대로 사용한다.
        if (!Files.exists(configPath)) {
            return defaultPath;
        }

        try {
            String configuredPath = readConfigValue(configPath, key);
            // 값이 비어 있으면 유효한 경로를 지정하지 않은 것으로 보고 기본값으로 되돌린다.
            if (configuredPath.isEmpty()) {
                return defaultPath;
            }

            Path resolvedPath = Paths.get(configuredPath);
            // 상대 경로는 실행 위치에 따라 해석이 달라질 수 있어, 예측 가능한 동작을 위해 허용하지 않는다.
            if (!resolvedPath.isAbsolute()) {
                return defaultPath;
            }

            return resolvedPath.normalize();
        } catch (IOException exception) {
            // 설정 파일을 읽지 못해도 앱의 핵심 기능은 계속 사용할 수 있어야 하므로 기본 경로를 반환한다.
            return defaultPath;
        }
    }

    /**
     * 단순 `key=value` 형식의 설정 파일에서 특정 키의 값을 찾아 반환한다.
     *
     * @param configPath 설정 파일 경로
     * @param key 읽어 올 설정 키
     * @return 키에 해당하는 값, 없으면 빈 문자열
     * @throws IOException 설정 파일을 읽는 중 오류가 발생한 경우
     */
    private static String readConfigValue(Path configPath, String key) throws IOException {
        List<String> lines = Files.readAllLines(configPath);

        for (String line : lines) {
            String trimmedLine = line.trim();
            // 빈 줄과 주석 줄은 실제 설정이 아니므로 건너뛴다.
            if (trimmedLine.isEmpty() || trimmedLine.startsWith("#") || trimmedLine.startsWith("!")) {
                continue;
            }

            int separatorIndex = trimmedLine.indexOf('=');
            // `=`가 없으면 기대한 설정 형식이 아니므로 무시한다.
            if (separatorIndex < 0) {
                continue;
            }

            String currentKey = trimmedLine.substring(0, separatorIndex).trim();
            // 요청한 키와 일치하는 항목만 반환하고, 나머지 줄은 계속 탐색한다.
            if (!key.equals(currentKey)) {
                continue;
            }

            return trimmedLine.substring(separatorIndex + 1).trim();
        }

        return "";
    }

    /**
     * 메인 창의 레이아웃과 버튼, 파일 목록 컴포넌트를 구성한다.
     * 상단에는 입력/출력 경로를, 중앙에는 파일 목록을, 하단에는 실행 버튼을 배치한다.
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

        refreshButton.addActionListener(event -> loadDocFiles());
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
     * 입력 폴더에서 병합 가능한 `.docx` 파일 목록을 읽어 리스트에 표시한다.
     * 템플릿 파일과 이미 생성된 결과 파일은 다시 병합하지 않도록 제외한다.
     */
    private void loadDocFiles() {
        listModel.clear();

        // 입력 폴더가 없으면 사용자가 파일을 선택할 수 없으므로 즉시 안내하고 종료한다.
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
                    // 병합 대상은 워드 문서만 허용한다.
                    .filter(name -> name.toLowerCase().endsWith(".docx"))
                    // 템플릿 자체는 입력 보고서가 아니므로 목록에서 숨긴다.
                    .filter(name -> !name.equalsIgnoreCase(ReportMerger.TEMPLATE_FILE_NAME))
                    // 기존 병합 결과물을 다시 입력으로 삼지 않도록 파일명 패턴으로 제외한다.
                    .filter(name -> !name.matches("\\d{8}_.+\\.docx"))
                    .sorted(String.CASE_INSENSITIVE_ORDER)
                    .collect(Collectors.toList());

            // 선택 가능한 문서가 하나도 없으면 안내 문구만 보여 주고 리스트 선택을 비활성화한다.
            if (docFiles.isEmpty()) {
                listModel.addElement("No source .docx files found.");
                fileList.setEnabled(false);
                return;
            }

            // 정상적으로 문서를 찾은 경우 선택 기능을 다시 활성화하고 목록을 채운다.
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
     * 사용자가 선택한 보고서 파일을 백그라운드에서 병합하고 진행 팝업을 표시한다.
     * 병합이 끝나면 같은 팝업에서 성공 또는 실패 상태를 보여 준다.
     */
    private void mergeSelectedFiles() {
        List<String> selectedFiles = fileList.getSelectedValuesList();
        // 병합할 파일을 선택하지 않은 경우에는 실제 작업을 시작하지 않고 안내만 보여 준다.
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

        // 병합 중에는 동일 작업을 중복 실행하지 못하도록 메인 창의 입력을 잠시 잠근다.
        setMergeControlsEnabled(false);
        MergeProgressDialog progressDialog = new MergeProgressDialog(this);
        progressDialog.updateProgress(0, "Preparing merge...");

        // 실제 병합은 EDT를 막지 않도록 SwingWorker의 백그라운드 스레드에서 수행한다.
        SwingWorker<Path, String> worker = new SwingWorker<>() {
            @Override
            protected Path doInBackground() throws Exception {
                return ReportMerger.mergeReports(sourceFiles, REPORT_DOC_OUT_PATH, (percent, message) -> {
                    // 진행률 값은 progress 프로퍼티로, 상세 단계 메시지는 publish로 UI 스레드에 전달한다.
                    setProgress(percent);
                    publish(message);
                });
            }

            @Override
            protected void process(List<String> chunks) {
                // 여러 메시지가 한 번에 들어오면 가장 최근 단계만 화면에 표시한다.
                if (chunks.isEmpty()) {
                    return;
                }

                String latestMessage = chunks.get(chunks.size() - 1);
                progressDialog.updateMessage(latestMessage);
            }

            @Override
            protected void done() {
                try {
                    // get()은 백그라운드 작업의 최종 결과 파일 경로를 반환한다.
                    Path outputFile = get();
                    progressDialog.completeSuccessfully(outputFile);
                } catch (Exception exception) {
                    // 병합 실패 시 원인 메시지를 팝업 내부에 그대로 보여 준다.
                    progressDialog.completeWithError(buildErrorMessage(exception));
                } finally {
                    // 성공/실패와 관계없이 메인 창 조작은 다시 가능해야 하므로 항상 실행한다.
                    setMergeControlsEnabled(true);
                }
            }
        };

        worker.addPropertyChangeListener((PropertyChangeEvent event) -> {
            // 다양한 속성 변경 중 진행률 변화만 프로그래스 바로 반영한다.
            if (!"progress".equals(event.getPropertyName())) {
                return;
            }

            Object value = event.getNewValue();
            // progress 값은 정수 퍼센트로 전달되므로 안전하게 형을 확인한 뒤 사용한다.
            if (value instanceof Integer progressValue) {
                progressDialog.updateProgress(progressValue, null);
            }
        });

        // 작업을 시작한 뒤 모달 팝업을 열어 사용자가 진행 상태를 바로 확인하도록 한다.
        worker.execute();
        progressDialog.setVisible(true);
    }

    /**
     * 병합 작업 중 메인 창의 조작 가능 여부를 한 번에 제어한다.
     *
     * @param enabled `true`이면 컨트롤을 활성화하고, `false`이면 비활성화한다
     */
    private void setMergeControlsEnabled(boolean enabled) {
        refreshButton.setEnabled(enabled);
        mergeReportsButton.setEnabled(enabled);
        // 파일 목록은 전체 활성화 상태뿐 아니라 실제 선택 가능한 항목이 있는지도 함께 확인한다.
        fileList.setEnabled(
                enabled && !listModel.isEmpty() && !"No source .docx files found.".equals(listModel.get(0)));
    }

    /**
     * 병합 진행 상황과 최종 결과를 보여 주는 모달 대화상자다.
     * 작업 중에는 닫을 수 없고, 완료 후에는 닫기 버튼이 활성화된다.
     */
    private static final class MergeProgressDialog extends JDialog {
        private final JLabel statusLabel = new JLabel("Preparing merge...");
        private final JLabel detailLabel = new JLabel(" ");
        private final JProgressBar progressBar = new JProgressBar(0, 100);
        private final JButton closeButton = new JButton("Close");

        /**
         * 진행률 팝업의 기본 UI를 구성한다.
         *
         * @param owner 팝업의 부모 프레임
         */
        MergeProgressDialog(JFrame owner) {
            super(owner, "Merging Reports", true);
            setLayout(new BorderLayout(12, 12));
            setDefaultCloseOperation(JDialog.DO_NOTHING_ON_CLOSE);

            progressBar.setStringPainted(true);
            progressBar.setValue(0);
            progressBar.setString("0%");
            progressBar.setPreferredSize(new Dimension(280, progressBar.getPreferredSize().height));

            closeButton.setEnabled(false);
            closeButton.addActionListener(event -> dispose());

            JPanel contentPanel = new JPanel(new GridLayout(0, 1, 0, 8));
            // GridLayout에서 프로그래스 바가 과도하게 늘어나는 것을 막기 위해 별도 패널에 감싼다.
            JPanel progressPanel = new JPanel(new FlowLayout(FlowLayout.CENTER, 0, 0));
            progressPanel.add(progressBar);
            contentPanel.add(statusLabel);
            contentPanel.add(detailLabel);
            contentPanel.add(progressPanel);

            JPanel buttonPanel = new JPanel();
            buttonPanel.add(closeButton);

            add(contentPanel, BorderLayout.CENTER);
            add(buttonPanel, BorderLayout.SOUTH);

            getRootPane().setDefaultButton(closeButton);
            setPreferredSize(new Dimension(420, 170));
            pack();
            setLocationRelativeTo(owner);
        }

        /**
         * 진행률 값과 현재 작업 메시지를 팝업에 반영한다.
         *
         * @param percent 0~100 범위의 진행 퍼센트
         * @param message 표시할 상세 단계 메시지, 없으면 기존 메시지를 유지한다
         */
        void updateProgress(int percent, String message) {
            progressBar.setValue(percent);
            progressBar.setString(percent + "%");
            // 메시지가 제공된 경우에만 라벨을 갱신해 불필요한 깜빡임을 줄인다.
            if (message != null && !message.isBlank()) {
                detailLabel.setText(message);
            }
        }

        /**
         * 상세 단계 메시지만 갱신한다.
         *
         * @param message 팝업에 표시할 메시지
         */
        void updateMessage(String message) {
            if (message != null && !message.isBlank()) {
                detailLabel.setText(message);
            }
        }

        /**
         * 병합 성공 상태를 표시하고 결과 파일 경로를 사용자에게 보여 준다.
         *
         * @param outputFile 생성된 결과 문서 경로
         */
        void completeSuccessfully(Path outputFile) {
            statusLabel.setText("Merge complete.");
            detailLabel.setText("<html>Merged report created:<br>" + outputFile.toAbsolutePath() + "</html>");
            progressBar.setValue(100);
            progressBar.setString("100%");
            finish();
        }

        /**
         * 병합 실패 상태를 표시하고 오류 메시지를 보여 준다.
         *
         * @param errorMessage 사용자에게 노출할 오류 메시지
         */
        void completeWithError(String errorMessage) {
            statusLabel.setText("Merge failed.");
            detailLabel.setText("<html>Failed to merge reports.<br>" + escapeHtml(errorMessage) + "</html>");
            progressBar.setString("Failed");
            finish();
        }

        /**
         * 진행 팝업을 종료 가능한 상태로 전환한다.
         * 완료 후에는 닫기 버튼과 일반적인 닫기 동작이 활성화된다.
         */
        private void finish() {
            closeButton.setEnabled(true);
            setDefaultCloseOperation(JDialog.DISPOSE_ON_CLOSE);
        }

        /**
         * 오류 메시지를 HTML 라벨에 안전하게 표시할 수 있도록 특수 문자를 이스케이프한다.
         *
         * @param text 변환할 원본 문자열
         * @return HTML에 삽입 가능한 문자열
         */
        private static String escapeHtml(String text) {
            if (text == null) {
                return "";
            }

            return text
                    .replace("&", "&amp;")
                    .replace("<", "&lt;")
                    .replace(">", "&gt;");
        }
    }

    /**
     * 예외 체인에서 사용자에게 보여 줄 수 있는 오류 메시지를 추출한다.
     *
     * @param throwable 분석할 예외 객체
     * @return 사용자 안내용 오류 메시지
     */
    private static String buildErrorMessage(Throwable throwable) {
        if (throwable == null) {
            return "Unknown error";
        }

        // SwingWorker의 get()은 실제 원인 예외를 ExecutionException으로 감쌀 수 있어 내부 원인을 우선 사용한다.
        if (throwable instanceof java.util.concurrent.ExecutionException && throwable.getCause() != null) {
            return buildErrorMessage(throwable.getCause());
        }

        String message = throwable.getMessage();
        // 예외 자체가 충분한 메시지를 가지고 있으면 가장 간단한 형태로 그대로 보여 준다.
        if (message != null && !message.isBlank()) {
            return message;
        }

        Throwable cause = throwable.getCause();
        // 현재 예외 메시지가 비어 있으면 원인 예외를 따라가며 더 의미 있는 설명을 찾는다.
        if (cause != null && cause != throwable) {
            String causeMessage = buildErrorMessage(cause);
            if (causeMessage != null && !causeMessage.isBlank()) {
                return throwable.getClass().getSimpleName() + ": " + causeMessage;
            }
        }

        return throwable.getClass().getSimpleName();
    }

    /**
     * Swing 이벤트 디스패치 스레드에서 애플리케이션 메인 창을 실행한다.
     *
     * @param args 실행 인자, 현재는 사용하지 않음
     */
    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            // Swing 컴포넌트는 EDT에서 생성해야 화면 갱신과 이벤트 처리가 안전하다.
            WeeklyMergeApp app = new WeeklyMergeApp();
            app.setVisible(true);
        });
    }
}
