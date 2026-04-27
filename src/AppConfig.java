import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;

/**
 * app.properties에서 경로 설정을 읽어 앱 전반에서 사용할 경로를 제공한다.
 * 설정 파일이 없거나 값이 유효하지 않으면 현재 작업 폴더 기반 기본 경로를 사용한다.
 */
public final class AppConfig {
    private static final String CONFIG_FILE_NAME = "app.properties";
    private static final String REPORT_DOC_PATH_KEY = "report.doc.path";
    private static final String REPORT_DOC_OUT_PATH_KEY = "report.doc.out.path";

    // 설정된 원본 경로가 없거나 유효하지 않으면 현재 작업 폴더를 기본 경로로 사용한다.
    public static final Path REPORT_DOC_PATH = loadConfiguredAbsolutePath(
            REPORT_DOC_PATH_KEY,
            Paths.get(System.getProperty("user.dir")));
    // 병합 결과 문서는 기본적으로 현재 작업 폴더 아래의 `output` 폴더에 저장한다.
    public static final Path REPORT_DOC_OUT_PATH = loadConfiguredAbsolutePath(
            REPORT_DOC_OUT_PATH_KEY,
            Paths.get(System.getProperty("user.dir"), "output"));

    /**
     * 설정 제공 유틸리티 클래스이므로 외부에서 인스턴스를 생성하지 못하게 막는다.
     */
    private AppConfig() {
    }

    /**
     * 설정 파일에서 절대 경로를 읽고, 사용할 수 없으면 기본 경로를 반환한다.
     *
     * @param key         읽어 올 설정 키
     * @param defaultPath 설정값이 없거나 잘못된 경우 사용할 기본 경로
     * @return 정규화된 절대 경로 또는 기본 경로
     */
    private static Path loadConfiguredAbsolutePath(String key, Path defaultPath) {
        Path configPath = Paths.get(CONFIG_FILE_NAME);
        // 설정 파일이 없어도 앱은 동작해야 하므로 기본 경로를 그대로 사용한다.
        if (!Files.exists(configPath)) {
            return defaultPath;
        }

        try {
            String configuredPath = readConfigValue(configPath, key);
            // 값이 비어 있으면 경로가 지정되지 않은 것으로 보고 기본 경로로 되돌린다.
            if (configuredPath.isEmpty()) {
                return defaultPath;
            }

            Path resolvedPath = Paths.get(configuredPath);
            // 상대 경로는 실행 위치에 따라 해석이 달라질 수 있어 허용하지 않는다.
            if (!resolvedPath.isAbsolute()) {
                return defaultPath;
            }

            return resolvedPath.normalize();
        } catch (IOException exception) {
            // 설정 파일을 읽지 못해도 앱 기능은 계속 사용할 수 있어야 하므로 기본 경로를 반환한다.
            return defaultPath;
        }
    }

    /**
     * 단순한 {@code key=value} 형식의 설정 파일에서 특정 키의 값을 찾는다.
     *
     * @param configPath 설정 파일 경로
     * @param key        읽어 올 설정 키
     * @return 일치하는 값, 없으면 빈 문자열
     * @throws IOException 설정 파일을 읽을 수 없는 경우
     */
    private static String readConfigValue(Path configPath, String key) throws IOException {
        List<String> lines = Files.readAllLines(configPath);

        for (String line : lines) {
            String trimmedLine = line.trim();
            // 빈 줄과 주석 줄은 실제 설정 항목이 아니므로 건너뛴다.
            if (trimmedLine.isEmpty() || trimmedLine.startsWith("#") || trimmedLine.startsWith("!")) {
                continue;
            }

            int separatorIndex = trimmedLine.indexOf('=');
            // '=' 문자가 없으면 기대한 설정 형식이 아니므로 무시한다.
            if (separatorIndex < 0) {
                continue;
            }

            String currentKey = trimmedLine.substring(0, separatorIndex).trim();
            // 요청한 키와 일치하는 항목만 반환하고 나머지는 계속 탐색한다.
            if (!key.equals(currentKey)) {
                continue;
            }

            return trimmedLine.substring(separatorIndex + 1).trim();
        }

        return "";
    }
}
