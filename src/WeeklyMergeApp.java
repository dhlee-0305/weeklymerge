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
    // Falls back to the current working directory when the configured source path is missing or invalid.
    private static final String CONFIG_FILE_NAME = "app.properties";
    private static final String REPORT_DOC_PATH_KEY = "report.doc.path";
    private static final String REPORT_DOC_OUT_PATH_KEY = "report.doc.out.path";
    private static final Path DEFAULT_REPORT_DOC_PATH = Paths.get(System.getProperty("user.dir"));
    private static final Path REPORT_DOC_PATH = loadConfiguredAbsolutePath(REPORT_DOC_PATH_KEY, DEFAULT_REPORT_DOC_PATH);
    // Writes merged reports to the default `output` folder under the current working directory.
    private static final Path DEFAULT_REPORT_DOC_OUT_PATH = Paths.get(System.getProperty("user.dir"), "output");
    private static final Path REPORT_DOC_OUT_PATH = loadConfiguredAbsolutePath(
            REPORT_DOC_OUT_PATH_KEY,
            DEFAULT_REPORT_DOC_OUT_PATH);

    private final DefaultListModel<String> listModel = new DefaultListModel<>();
    private final JList<String> fileList = new JList<>(listModel);
    private final JButton refreshButton = new JButton("Refresh Files");
    private final JButton mergeReportsButton = new JButton("Merge Reports");

    /**
     * Creates the main application window and initializes the UI and source file list.
     * The user can review mergeable report files as soon as the window is shown.
     */
    public WeeklyMergeApp() {
        super("Weekly Report Merge");
        initializeUi();
        loadDocFiles();
    }

    /**
     * Reads an absolute path from the config file and falls back to a default path when needed.
     *
     * @param key config key to read
     * @param defaultPath path to use when the config value is missing or invalid
     * @return normalized absolute path or the default path
     */
    private static Path loadConfiguredAbsolutePath(String key, Path defaultPath) {
        Path configPath = Paths.get(CONFIG_FILE_NAME);
        // Missing config should not block the app, so the default path is used as-is.
        if (!Files.exists(configPath)) {
            return defaultPath;
        }

        try {
            String configuredPath = readConfigValue(configPath, key);
            // An empty config value is treated as "not configured" and falls back to the default path.
            if (configuredPath.isEmpty()) {
                return defaultPath;
            }

            Path resolvedPath = Paths.get(configuredPath);
            // Relative paths are rejected to keep behavior predictable across launch locations.
            if (!resolvedPath.isAbsolute()) {
                return defaultPath;
            }

            return resolvedPath.normalize();
        } catch (IOException exception) {
            // Config read failures should not stop the app from working, so the default path is returned.
            return defaultPath;
        }
    }

    /**
     * Finds the value for a key in a simple {@code key=value} config file.
     *
     * @param configPath config file path
     * @param key config key to read
     * @return matching value, or an empty string when the key is absent
     * @throws IOException if the config file cannot be read
     */
    private static String readConfigValue(Path configPath, String key) throws IOException {
        List<String> lines = Files.readAllLines(configPath);

        for (String line : lines) {
            String trimmedLine = line.trim();
            // Blank lines and comments are skipped because they are not real config entries.
            if (trimmedLine.isEmpty() || trimmedLine.startsWith("#") || trimmedLine.startsWith("!")) {
                continue;
            }

            int separatorIndex = trimmedLine.indexOf('=');
            // Lines without '=' do not match the expected config format.
            if (separatorIndex < 0) {
                continue;
            }

            String currentKey = trimmedLine.substring(0, separatorIndex).trim();
            // Only the requested key is returned; all other entries are ignored.
            if (!key.equals(currentKey)) {
                continue;
            }

            return trimmedLine.substring(separatorIndex + 1).trim();
        }

        return "";
    }

    /**
     * Builds the main window layout, including the path labels, file list, and action buttons.
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
     * Loads mergeable {@code .docx} files from the source folder and shows them in the list.
     * Template files and previously generated output files are excluded.
     */
    private void loadDocFiles() {
        listModel.clear();

        // Without a source folder the user cannot pick files, so the method stops after showing a warning.
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
                    // Only Word documents are valid merge inputs.
                    .filter(name -> name.toLowerCase().endsWith(".docx"))
                    // The template document is not a source report and must stay out of the selection list.
                    .filter(name -> !name.equalsIgnoreCase(ReportMerger.TEMPLATE_FILE_NAME))
                    // Generated output files are excluded to prevent accidental re-merge of previous results.
                    .filter(name -> !name.matches("\\d{8}_.+\\.docx"))
                    .sorted(String.CASE_INSENSITIVE_ORDER)
                    .collect(Collectors.toList());

            // When nothing can be merged, show a placeholder message and disable list selection.
            if (docFiles.isEmpty()) {
                listModel.addElement("No source .docx files found.");
                fileList.setEnabled(false);
                return;
            }

            // Re-enable selection once valid files have been discovered.
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
     * Merges the selected report files in the background and shows a progress dialog.
     * The same dialog is updated with the final success or failure result.
     */
    private void mergeSelectedFiles() {
        List<String> selectedFiles = fileList.getSelectedValuesList();
        // A merge should not start without selected files.
        if (selectedFiles.isEmpty()) {
            JOptionPane.showMessageDialog(
                    this,
                    "Select one or more report files to merge.",
                    "Notice",
                    JOptionPane.INFORMATION_MESSAGE);
            return;
        }

        // A merge should not start if there is no destination folder for the output document.
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

        // Disable the main controls while a merge is in progress to prevent duplicate runs.
        setMergeControlsEnabled(false);
        MergeProgressDialog progressDialog = new MergeProgressDialog(this);
        progressDialog.updateProgress(0, "Preparing merge...");

        // The merge runs off the EDT so the UI stays responsive.
        SwingWorker<Path, String> worker = new SwingWorker<>() {
            @Override
            protected Path doInBackground() throws Exception {
                return ReportMerger.mergeReports(sourceFiles, REPORT_DOC_OUT_PATH, (percent, message) -> {
                    // Progress percent and step messages are forwarded to the UI through SwingWorker hooks.
                    setProgress(percent);
                    publish(message);
                });
            }

            @Override
            protected void process(List<String> chunks) {
                // When several messages arrive together, the latest one best represents the current step.
                if (chunks.isEmpty()) {
                    return;
                }

                String latestMessage = chunks.get(chunks.size() - 1);
                progressDialog.updateMessage(latestMessage);
            }

            @Override
            protected void done() {
                try {
                    // get() returns the final output file path from the background task.
                    Path outputFile = get();
                    progressDialog.completeSuccessfully(outputFile);
                } catch (Exception exception) {
                    // Failure details are shown inside the same dialog for a consistent user flow.
                    progressDialog.completeWithError(buildErrorMessage(exception));
                } finally {
                    // Re-enable the main controls regardless of success or failure.
                    setMergeControlsEnabled(true);
                }
            }
        };

        worker.addPropertyChangeListener((PropertyChangeEvent event) -> {
            // Only progress property updates should change the progress bar.
            if (!"progress".equals(event.getPropertyName())) {
                return;
            }

            Object value = event.getNewValue();
            // The progress property is expected as an integer percentage.
            if (value instanceof Integer progressValue) {
                progressDialog.updateProgress(progressValue, null);
            }
        });

        // The modal dialog is shown after the worker starts so the user can track progress immediately.
        worker.execute();
        progressDialog.setVisible(true);
    }

    /**
     * Enables or disables the main window controls used during merge operations.
     *
     * @param enabled {@code true} to enable controls, {@code false} to disable them
     */
    private void setMergeControlsEnabled(boolean enabled) {
        refreshButton.setEnabled(enabled);
        mergeReportsButton.setEnabled(enabled);
        // The list is enabled only when the screen is enabled and there are real selectable file entries.
        fileList.setEnabled(
                enabled && !listModel.isEmpty() && !"No source .docx files found.".equals(listModel.get(0)));
    }

    /**
     * Modal dialog that shows merge progress and the final merge result.
     * It stays non-closable during work and becomes dismissible after completion.
     */
    private static final class MergeProgressDialog extends JDialog {
        private final JLabel statusLabel = new JLabel("Preparing merge...");
        private final JLabel detailLabel = new JLabel(" ");
        private final JProgressBar progressBar = new JProgressBar(0, 100);
        private final JButton closeButton = new JButton("Close");

        /**
         * Builds the default UI for the progress dialog.
         *
         * @param owner parent frame for the dialog
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
            // Wrapping the progress bar prevents GridLayout from stretching it to the full dialog width.
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
         * Updates the progress bar and, optionally, the detail message.
         *
         * @param percent progress percentage in the range 0..100
         * @param message detail message to show, or {@code null} to keep the current text
         */
        void updateProgress(int percent, String message) {
            progressBar.setValue(percent);
            progressBar.setString(percent + "%");
            // The detail label is updated only when a real message is provided.
            if (message != null && !message.isBlank()) {
                detailLabel.setText(message);
            }
        }

        /**
         * Updates only the detail message.
         *
         * @param message text to show in the detail area
         */
        void updateMessage(String message) {
            if (message != null && !message.isBlank()) {
                detailLabel.setText(message);
            }
        }

        /**
         * Switches the dialog to the success state and shows the output file path.
         *
         * @param outputFile generated merged document
         */
        void completeSuccessfully(Path outputFile) {
            statusLabel.setText("Merge complete.");
            detailLabel.setText("<html>Merged report created:<br>" + outputFile.toAbsolutePath() + "</html>");
            progressBar.setValue(100);
            progressBar.setString("100%");
            finish();
        }

        /**
         * Switches the dialog to the failure state and shows the error message.
         *
         * @param errorMessage user-facing error description
         */
        void completeWithError(String errorMessage) {
            statusLabel.setText("Merge failed.");
            detailLabel.setText("<html>Failed to merge reports.<br>" + escapeHtml(errorMessage) + "</html>");
            progressBar.setString("Failed");
            finish();
        }

        /**
         * Makes the dialog closable after the merge result has been shown.
         */
        private void finish() {
            closeButton.setEnabled(true);
            setDefaultCloseOperation(JDialog.DISPOSE_ON_CLOSE);
        }

        /**
         * Escapes special characters so error messages can be displayed safely inside HTML labels.
         *
         * @param text raw input text
         * @return HTML-safe text
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
     * Extracts a user-facing error message from an exception chain.
     *
     * @param throwable exception to inspect
     * @return readable error message for the user
     */
    private static String buildErrorMessage(Throwable throwable) {
        if (throwable == null) {
            return "Unknown error";
        }

        // SwingWorker#get may wrap the original cause, so the inner exception is preferred.
        if (throwable instanceof java.util.concurrent.ExecutionException && throwable.getCause() != null) {
            return buildErrorMessage(throwable.getCause());
        }

        String message = throwable.getMessage();
        // A direct exception message is the simplest readable output for the user.
        if (message != null && !message.isBlank()) {
            return message;
        }

        Throwable cause = throwable.getCause();
        // If the current exception has no message, walk the cause chain to find a better explanation.
        if (cause != null && cause != throwable) {
            String causeMessage = buildErrorMessage(cause);
            if (causeMessage != null && !causeMessage.isBlank()) {
                return throwable.getClass().getSimpleName() + ": " + causeMessage;
            }
        }

        return throwable.getClass().getSimpleName();
    }

    /**
     * Launches the main Swing window on the Event Dispatch Thread.
     *
     * @param args command-line arguments, currently unused
     */
    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            // Swing components must be created on the EDT for safe rendering and event handling.
            WeeklyMergeApp app = new WeeklyMergeApp();
            app.setVisible(true);
        });
    }
}
