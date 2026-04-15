import java.awt.BorderLayout;
import java.awt.Dimension;
import java.awt.GridLayout;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
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
    private static final String CONFIG_FILE_NAME = "app.properties";
    private static final String REPORT_DOC_PATH_KEY = "report.doc.path";
    private static final String REPORT_DOC_OUT_PATH_KEY = "report.doc.out.path";
    private static final Path DEFAULT_REPORT_DOC_PATH =
            Paths.get(System.getProperty("user.home"), "Documents");
    private static final Path REPORT_DOC_PATH =
            loadConfiguredAbsolutePath(REPORT_DOC_PATH_KEY, DEFAULT_REPORT_DOC_PATH);
    private static final Path REPORT_DOC_OUT_PATH =
            loadConfiguredAbsolutePath(REPORT_DOC_OUT_PATH_KEY, DEFAULT_REPORT_DOC_PATH);

    private final DefaultListModel<String> listModel = new DefaultListModel<>();
    private final JList<String> fileList = new JList<>(listModel);

    public WeeklyMergeApp() {
        super("Weekly Merge Report Viewer");
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

        JButton showSelectedButton = new JButton("Show Selected Files");
        showSelectedButton.addActionListener(event -> showSelectedFiles());

        JButton mergeReportsButton = new JButton("Merge Reports");
        mergeReportsButton.addActionListener(event -> mergeSelectedFiles());

        JPanel bottomPanel = new JPanel();
        bottomPanel.add(showSelectedButton);
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

    private void showSelectedFiles() {
        List<String> selectedFiles = fileList.getSelectedValuesList();

        if (selectedFiles.isEmpty()) {
            JOptionPane.showMessageDialog(
                    this,
                    "No files are selected.",
                    "Notice",
                    JOptionPane.INFORMATION_MESSAGE);
            return;
        }

        JOptionPane.showMessageDialog(
                this,
                selectedFiles.stream().collect(Collectors.joining("\n")),
                "Selected Files",
                JOptionPane.INFORMATION_MESSAGE);
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

        try {
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

    private static String buildErrorMessage(Throwable throwable) {
        if (throwable == null) {
            return "Unknown error";
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
