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
    private static final String CONFIG_FILE_NAME = "app.properties";
    private static final String REPORT_DOC_PATH_KEY = "report.doc.path";
    private static final String REPORT_DOC_OUT_PATH_KEY = "report.doc.out.path";
    private static final Path DEFAULT_REPORT_DOC_PATH = Paths.get(System.getProperty("user.dir"));
    private static final Path REPORT_DOC_PATH = loadConfiguredAbsolutePath(REPORT_DOC_PATH_KEY,
            DEFAULT_REPORT_DOC_PATH);
    private static final Path DEFAULT_REPORT_DOC_OUT_PATH = Paths.get(System.getProperty("user.dir"), "output");
    private static final Path REPORT_DOC_OUT_PATH = loadConfiguredAbsolutePath(
            REPORT_DOC_OUT_PATH_KEY,
            DEFAULT_REPORT_DOC_OUT_PATH);

    private final DefaultListModel<String> listModel = new DefaultListModel<>();
    private final JList<String> fileList = new JList<>(listModel);
    private final JButton refreshButton = new JButton("Refresh Files");
    private final JButton mergeReportsButton = new JButton("Merge Reports");

    public WeeklyMergeApp() {
        super("Weekly Report Merge");
        initializeUi();
        loadDocFiles();
    }

    private static Path loadConfiguredAbsolutePath(String key, Path defaultPath) {
        Path configPath = Paths.get(CONFIG_FILE_NAME);
        if (!Files.exists(configPath)) {
            return defaultPath;
        }

        try {
            String configuredPath = readConfigValue(configPath, key);
            if (configuredPath.isEmpty()) {
                return defaultPath;
            }

            Path resolvedPath = Paths.get(configuredPath);
            if (!resolvedPath.isAbsolute()) {
                return defaultPath;
            }

            return resolvedPath.normalize();
        } catch (IOException exception) {
            return defaultPath;
        }
    }

    private static String readConfigValue(Path configPath, String key) throws IOException {
        List<String> lines = Files.readAllLines(configPath);

        for (String line : lines) {
            String trimmedLine = line.trim();
            if (trimmedLine.isEmpty() || trimmedLine.startsWith("#") || trimmedLine.startsWith("!")) {
                continue;
            }

            int separatorIndex = trimmedLine.indexOf('=');
            if (separatorIndex < 0) {
                continue;
            }

            String currentKey = trimmedLine.substring(0, separatorIndex).trim();
            if (!key.equals(currentKey)) {
                continue;
            }

            return trimmedLine.substring(separatorIndex + 1).trim();
        }

        return "";
    }

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

    private void loadDocFiles() {
        listModel.clear();

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
                    .filter(name -> name.toLowerCase().endsWith(".docx"))
                    .filter(name -> !name.equalsIgnoreCase(ReportMerger.TEMPLATE_FILE_NAME))
                    .filter(name -> !name.matches("\\d{8}_.+\\.docx"))
                    .sorted(String.CASE_INSENSITIVE_ORDER)
                    .collect(Collectors.toList());

            if (docFiles.isEmpty()) {
                listModel.addElement("No source .docx files found.");
                fileList.setEnabled(false);
                return;
            }

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

    private void mergeSelectedFiles() {
        List<String> selectedFiles = fileList.getSelectedValuesList();
        if (selectedFiles.isEmpty()) {
            JOptionPane.showMessageDialog(
                    this,
                    "Select one or more report files to merge.",
                    "Notice",
                    JOptionPane.INFORMATION_MESSAGE);
            return;
        }

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

        setMergeControlsEnabled(false);
        MergeProgressDialog progressDialog = new MergeProgressDialog(this);
        progressDialog.updateProgress(0, "Preparing merge...");

        SwingWorker<Path, String> worker = new SwingWorker<>() {
            @Override
            protected Path doInBackground() throws Exception {
                return ReportMerger.mergeReports(sourceFiles, REPORT_DOC_OUT_PATH, (percent, message) -> {
                    setProgress(percent);
                    publish(message);
                });
            }

            @Override
            protected void process(List<String> chunks) {
                if (chunks.isEmpty()) {
                    return;
                }

                String latestMessage = chunks.get(chunks.size() - 1);
                progressDialog.updateMessage(latestMessage);
            }

            @Override
            protected void done() {
                try {
                    Path outputFile = get();
                    progressDialog.completeSuccessfully(outputFile);
                } catch (Exception exception) {
                    progressDialog.completeWithError(buildErrorMessage(exception));
                } finally {
                    setMergeControlsEnabled(true);
                }
            }
        };

        worker.addPropertyChangeListener((PropertyChangeEvent event) -> {
            if (!"progress".equals(event.getPropertyName())) {
                return;
            }

            Object value = event.getNewValue();
            if (value instanceof Integer progressValue) {
                progressDialog.updateProgress(progressValue, null);
            }
        });

        worker.execute();
        progressDialog.setVisible(true);
    }

    private void setMergeControlsEnabled(boolean enabled) {
        refreshButton.setEnabled(enabled);
        mergeReportsButton.setEnabled(enabled);
        fileList.setEnabled(
                enabled && !listModel.isEmpty() && !"No source .docx files found.".equals(listModel.get(0)));
    }

    private static final class MergeProgressDialog extends JDialog {
        private final JLabel statusLabel = new JLabel("Preparing merge...");
        private final JLabel detailLabel = new JLabel(" ");
        private final JProgressBar progressBar = new JProgressBar(0, 100);
        private final JButton closeButton = new JButton("Close");

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

        void updateProgress(int percent, String message) {
            progressBar.setValue(percent);
            progressBar.setString(percent + "%");
            if (message != null && !message.isBlank()) {
                detailLabel.setText(message);
            }
        }

        void updateMessage(String message) {
            if (message != null && !message.isBlank()) {
                detailLabel.setText(message);
            }
        }

        void completeSuccessfully(Path outputFile) {
            statusLabel.setText("Merge complete.");
            detailLabel.setText("<html>Merged report created:<br>" + outputFile.toAbsolutePath() + "</html>");
            progressBar.setValue(100);
            progressBar.setString("100%");
            finish();
        }

        void completeWithError(String errorMessage) {
            statusLabel.setText("Merge failed.");
            detailLabel.setText("<html>Failed to merge reports.<br>" + escapeHtml(errorMessage) + "</html>");
            progressBar.setString("Failed");
            finish();
        }

        private void finish() {
            closeButton.setEnabled(true);
            setDefaultCloseOperation(JDialog.DISPOSE_ON_CLOSE);
        }

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

    private static String buildErrorMessage(Throwable throwable) {
        if (throwable == null) {
            return "Unknown error";
        }

        if (throwable instanceof java.util.concurrent.ExecutionException && throwable.getCause() != null) {
            return buildErrorMessage(throwable.getCause());
        }

        String message = throwable.getMessage();
        if (message != null && !message.isBlank()) {
            return message;
        }

        Throwable cause = throwable.getCause();
        if (cause != null && cause != throwable) {
            String causeMessage = buildErrorMessage(cause);
            if (causeMessage != null && !causeMessage.isBlank()) {
                return throwable.getClass().getSimpleName() + ": " + causeMessage;
            }
        }

        return throwable.getClass().getSimpleName();
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            WeeklyMergeApp app = new WeeklyMergeApp();
            app.setVisible(true);
        });
    }
}
