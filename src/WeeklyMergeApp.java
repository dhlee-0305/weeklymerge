import java.awt.BorderLayout;
import java.awt.Dimension;
import java.awt.FlowLayout;
import java.awt.GridLayout;
import java.beans.PropertyChangeEvent;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
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
    private final DefaultListModel<String> listModel = new DefaultListModel<>();
    private final JList<String> fileList = new JList<>(listModel);
    private final JButton refreshButton = new JButton(Messages.get("button.refresh"));
    private final JButton mergeReportsButton = new JButton(Messages.get("button.merge"));
    private final List<String> selectionOrder = new ArrayList<>();

    /**
     * 메인 애플리케이션 창을 생성하고 UI와 원본 파일 목록을 초기화한다.
     * 창이 표시되면 사용자는 바로 병합 가능한 보고서 목록을 확인할 수 있다.
     */
    public WeeklyMergeApp() {
        super(Messages.get("window.title"));
        initializeUi();
        loadDocFiles();
    }

    /**
     * 경로 표시, 파일 목록, 실행 버튼을 포함한 메인 창 레이아웃을 구성한다.
     */
    private void initializeUi() {
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setLayout(new BorderLayout(12, 12));

        JPanel infoPanel = new JPanel(new GridLayout(2, 1));
        infoPanel.add(new JLabel(Messages.format("label.source.path", AppConfig.REPORT_DOC_PATH.toAbsolutePath())));
        infoPanel.add(new JLabel(Messages.format("label.output.path", AppConfig.REPORT_DOC_OUT_PATH.toAbsolutePath())));

        fileList.setSelectionMode(ListSelectionModel.MULTIPLE_INTERVAL_SELECTION);
        fileList.addListSelectionListener(e -> {
            if (e.getValueIsAdjusting()) {
                return;
            }
            List<String> currentSelection = fileList.getSelectedValuesList();
            Set<String> currentSet = new HashSet<>(currentSelection);
            selectionOrder.removeIf(item -> !currentSet.contains(item));
            Set<String> alreadyTracked = new HashSet<>(selectionOrder);
            for (String item : currentSelection) {
                if (!alreadyTracked.contains(item)) {
                    selectionOrder.add(item);
                }
            }
        });

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
     * 원본 폴더에서 병합 가능한 {@code .docx} 파일을 읽어 목록에 표시한다.
     * 템플릿 파일과 이전에 생성된 결과 파일은 제외한다.
     */
    private void loadDocFiles() {
        listModel.clear();
        selectionOrder.clear();

        // 원본 폴더가 없으면 사용자가 파일을 선택할 수 없으므로 경고 후 종료한다.
        if (!Files.isDirectory(AppConfig.REPORT_DOC_PATH)) {
            JOptionPane.showMessageDialog(
                    this,
                    Messages.format("error.source.folder.not.found", AppConfig.REPORT_DOC_PATH.toAbsolutePath()),
                    Messages.get("dialog.title.folder.not.found"),
                    JOptionPane.WARNING_MESSAGE);
            return;
        }

        try (Stream<Path> paths = Files.list(AppConfig.REPORT_DOC_PATH)) {
            List<String> docFiles = paths
                    .filter(Files::isRegularFile)
                    .map(Path::getFileName)
                    .map(Path::toString)
                    // 병합 입력 대상으로는 Word 문서만 허용한다.
                    .filter(name -> name.toLowerCase().endsWith(".docx"))
                    // 템플릿 문서는 병합 대상 보고서가 아니므로 목록에서 제외한다.
                    .filter(name -> !name.equalsIgnoreCase(ReportMerger.TEMPLATE_FILE_NAME))
                    // 이전 병합 결과가 다시 입력으로 선택되지 않도록 결과 파일은 제외한다.
                    .filter(name -> !name.matches("\\d{8}_.+\\.docx"))
                    .sorted(String.CASE_INSENSITIVE_ORDER)
                    .collect(Collectors.toList());

            // 병합 가능한 문서가 없으면 안내 문구만 보여 주고 목록 선택을 비활성화한다.
            if (docFiles.isEmpty()) {
                listModel.addElement(Messages.get("filelist.empty"));
                fileList.setEnabled(false);
                return;
            }

            // 유효한 문서를 찾은 경우 목록 선택을 다시 활성화한다.
            fileList.setEnabled(true);
            docFiles.forEach(listModel::addElement);
        } catch (Exception exception) {
            JOptionPane.showMessageDialog(
                    this,
                    Messages.format("error.file.list.read", exception.getMessage()),
                    Messages.get("dialog.title.error"),
                    JOptionPane.ERROR_MESSAGE);
        }
    }

    /**
     * 선택한 보고서 파일을 백그라운드에서 병합하고 진행 팝업을 표시한다.
     * 같은 팝업 안에서 최종 성공 또는 실패 결과까지 이어서 보여 준다.
     */
    private void mergeSelectedFiles() {
        // 선택된 파일이 없으면 병합을 시작하지 않는다.
        if (selectionOrder.isEmpty()) {
            JOptionPane.showMessageDialog(
                    this,
                    Messages.get("notice.select.files"),
                    Messages.get("dialog.title.notice"),
                    JOptionPane.INFORMATION_MESSAGE);
            return;
        }

        // 출력 폴더가 없으면 결과 파일을 저장할 수 없으므로 병합을 시작하지 않는다.
        if (!Files.isDirectory(AppConfig.REPORT_DOC_OUT_PATH)) {
            JOptionPane.showMessageDialog(
                    this,
                    Messages.format("error.output.folder.not.found", AppConfig.REPORT_DOC_OUT_PATH.toAbsolutePath()),
                    Messages.get("dialog.title.folder.not.found"),
                    JOptionPane.WARNING_MESSAGE);
            return;
        }

        List<Path> sourceFiles = selectionOrder.stream()
                .map(name -> AppConfig.REPORT_DOC_PATH.resolve(name))
                .collect(Collectors.toList());

        // 병합 중에는 중복 실행을 막기 위해 메인 화면 조작을 잠시 비활성화한다.
        setMergeControlsEnabled(false);
        MergeProgressDialog progressDialog = new MergeProgressDialog(this);
        progressDialog.updateProgress(0, Messages.get("progress.preparing"));

        // 병합은 EDT 밖에서 실행해 UI가 멈추지 않도록 한다.
        SwingWorker<Path, String> worker = new SwingWorker<>() {
            @Override
            protected Path doInBackground() throws Exception {
                return ReportMerger.mergeReports(sourceFiles, AppConfig.REPORT_DOC_OUT_PATH, (percent, message) -> {
                    // 진행률과 단계 메시지를 SwingWorker를 통해 UI 스레드로 전달한다.
                    setProgress(percent);
                    publish(message);
                });
            }

            @Override
            protected void process(List<String> chunks) {
                // 여러 메시지가 한 번에 전달되면 가장 최근 메시지만 표시한다.
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
                    // 실패 사유도 같은 팝업 안에 보여 주어 사용자 흐름을 일관되게 유지한다.
                    progressDialog.completeWithError(buildErrorMessage(exception));
                } finally {
                    // 성공 여부와 상관없이 메인 화면 조작은 다시 가능해야 한다.
                    setMergeControlsEnabled(true);
                }
            }
        };

        worker.addPropertyChangeListener((PropertyChangeEvent event) -> {
            // 다양한 속성 변경 중 진행률 변경만 프로그래스 바로 반영한다.
            if (!"progress".equals(event.getPropertyName())) {
                return;
            }

            Object value = event.getNewValue();
            // progress 값은 정수 퍼센트로 전달되므로 형을 확인한 뒤 사용한다.
            if (value instanceof Integer progressValue) {
                progressDialog.updateProgress(progressValue, null);
            }
        });

        // 작업 시작 직후 모달 팝업을 열어 사용자가 진행 상황을 바로 확인할 수 있게 한다.
        worker.execute();
        progressDialog.setVisible(true);
    }

    /**
     * 병합 작업 중 사용하는 메인 창 컨트롤의 활성화 여부를 설정한다.
     *
     * @param enabled {@code true}면 활성화, {@code false}면 비활성화
     */
    private void setMergeControlsEnabled(boolean enabled) {
        refreshButton.setEnabled(enabled);
        mergeReportsButton.setEnabled(enabled);
        // 목록은 화면이 활성 상태이고 실제 선택 가능한 항목이 있을 때만 켠다.
        fileList.setEnabled(
                enabled && !listModel.isEmpty() && !Messages.get("filelist.empty").equals(listModel.get(0)));
    }

    /**
     * 병합 진행 상태와 최종 결과를 보여 주는 모달 대화상자다.
     * 작업 중에는 닫을 수 없고, 완료 후에는 닫을 수 있다.
     */
    private static final class MergeProgressDialog extends JDialog {
        private final JLabel statusLabel = new JLabel(Messages.get("progress.preparing"));
        private final JLabel detailLabel = new JLabel(" ");
        private final JProgressBar progressBar = new JProgressBar(0, 100);
        private final JButton closeButton = new JButton(Messages.get("button.close"));

        /**
         * 진행 팝업의 기본 UI를 구성한다.
         *
         * @param owner 부모 프레임
         */
        MergeProgressDialog(JFrame owner) {
            super(owner, Messages.get("dialog.merge.title"), true);
            setLayout(new BorderLayout(12, 12));
            setDefaultCloseOperation(JDialog.DO_NOTHING_ON_CLOSE);

            progressBar.setStringPainted(true);
            progressBar.setValue(0);
            progressBar.setString("0%");
            progressBar.setPreferredSize(new Dimension(280, progressBar.getPreferredSize().height));

            closeButton.setEnabled(false);
            closeButton.addActionListener(event -> dispose());

            JPanel contentPanel = new JPanel(new GridLayout(0, 1, 0, 8));
            // 별도 패널로 감싸 GridLayout이 프로그래스 바를 전체 너비로 늘리지 않게 한다.
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
         * 진행 바 값과 상세 메시지를 갱신한다.
         *
         * @param percent 0~100 범위의 진행률
         * @param message 표시할 상세 메시지, {@code null}이면 기존 메시지를 유지한다
         */
        void updateProgress(int percent, String message) {
            progressBar.setValue(percent);
            progressBar.setString(percent + "%");
            // 실제 메시지가 전달된 경우에만 상세 라벨을 갱신한다.
            if (message != null && !message.isBlank()) {
                detailLabel.setText(message);
            }
        }

        /**
         * 상세 메시지만 갱신한다.
         *
         * @param message 상세 영역에 표시할 텍스트
         */
        void updateMessage(String message) {
            if (message != null && !message.isBlank()) {
                detailLabel.setText(message);
            }
        }

        /**
         * 팝업을 성공 상태로 전환하고 결과 파일 경로를 보여 준다.
         *
         * @param outputFile 생성된 병합 결과 문서
         */
        void completeSuccessfully(Path outputFile) {
            statusLabel.setText(Messages.get("progress.complete"));
            detailLabel.setText(Messages.format("progress.complete.detail", outputFile.toAbsolutePath()));
            progressBar.setValue(100);
            progressBar.setString("100%");
            finish();
        }

        /**
         * 팝업을 실패 상태로 전환하고 오류 메시지를 보여 준다.
         *
         * @param errorMessage 사용자에게 보여 줄 오류 설명
         */
        void completeWithError(String errorMessage) {
            statusLabel.setText(Messages.get("progress.failed"));
            detailLabel.setText(Messages.format("progress.failed.detail", escapeHtml(errorMessage)));
            progressBar.setString(Messages.get("progress.failed.label"));
            finish();
        }

        /**
         * 병합 결과를 보여 준 뒤 팝업을 닫을 수 있는 상태로 전환한다.
         */
        private void finish() {
            closeButton.setEnabled(true);
            setDefaultCloseOperation(JDialog.DISPOSE_ON_CLOSE);
        }

        /**
         * 오류 메시지를 HTML 라벨에 안전하게 표시할 수 있도록 특수 문자를 이스케이프한다.
         *
         * @param text 원본 텍스트
         * @return HTML에 안전한 문자열
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
     * @return 사용자가 이해할 수 있는 오류 메시지
     */
    private static String buildErrorMessage(Throwable throwable) {
        if (throwable == null) {
            return Messages.get("error.unknown");
        }

        // SwingWorker#get은 원인 예외를 감쌀 수 있으므로 내부 예외를 우선 사용한다.
        if (throwable instanceof java.util.concurrent.ExecutionException && throwable.getCause() != null) {
            return buildErrorMessage(throwable.getCause());
        }

        String message = throwable.getMessage();
        // 예외 자체에 메시지가 있으면 가장 간단하고 읽기 쉬운 설명으로 그대로 사용한다.
        if (message != null && !message.isBlank()) {
            return message;
        }

        Throwable cause = throwable.getCause();
        // 현재 예외 메시지가 없으면 원인 예외를 따라가며 더 나은 설명을 찾는다.
        if (cause != null && cause != throwable) {
            String causeMessage = buildErrorMessage(cause);
            if (causeMessage != null && !causeMessage.isBlank()) {
                return throwable.getClass().getSimpleName() + ": " + causeMessage;
            }
        }

        return throwable.getClass().getSimpleName();
    }

    /**
     * Swing 이벤트 디스패치 스레드에서 메인 창을 실행한다.
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
