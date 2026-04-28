import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.net.URISyntaxException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;

/**
 * app.properties에서 경로 설정을 읽어 앱 전반에서 사용할 경로를 제공한다.
 * 배포 환경별 경로는 JAR 옆의 app.properties로 관리하고,
 * 실행 폴더에 설정 파일이 없으면 JAR에 포함된 기본 설정을 클래스패스에서 로드한다.
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
        String configuredPath = readConfigValue(key);
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
    }

    /**
     * 실행 폴더, 애플리케이션 폴더, 클래스패스 순서로 app.properties에서 특정 키의 값을 찾는다.
     *
     * @param key 읽어 올 설정 키
     * @return 일치하는 값, 없거나 설정 파일을 읽을 수 없으면 빈 문자열
     */
    private static String readConfigValue(String key) {
        try {
            String configuredValue = readConfigValueFromWorkingDirectory(key);
            if (!configuredValue.isEmpty()) {
                return configuredValue;
            }
            configuredValue = readConfigValueFromApplicationDirectory(key);
            if (!configuredValue.isEmpty()) {
                return configuredValue;
            }
            return readConfigValueFromClasspath(key);
        } catch (IOException exception) {
            // 설정 파일을 읽지 못해도 앱 기능은 계속 사용할 수 있어야 하므로 기본 경로를 사용한다.
            return "";
        }
    }

    /**
     * 실행 폴더에 별도로 배치된 app.properties에서 특정 키의 값을 찾는다.
     *
     * @param key 읽어 올 설정 키
     * @return 일치하는 값, 없으면 빈 문자열
     * @throws IOException 설정 파일을 읽을 수 없는 경우
     */
    private static String readConfigValueFromWorkingDirectory(String key) throws IOException {
        Path configPath = Paths.get(CONFIG_FILE_NAME);
        if (!Files.exists(configPath)) {
            return "";
        }

        return findValue(Files.readAllLines(configPath, StandardCharsets.UTF_8), key);
    }

    /**
     * JAR 파일 또는 컴파일된 클래스가 있는 폴더의 app.properties에서 특정 키의 값을 찾는다.
     *
     * @param key 읽어 올 설정 키
     * @return 일치하는 값, 없으면 빈 문자열
     * @throws IOException 설정 파일을 읽을 수 없는 경우
     */
    private static String readConfigValueFromApplicationDirectory(String key) throws IOException {
        Path applicationPath = getApplicationPath();
        if (applicationPath == null) {
            return "";
        }

        Path applicationDirectory = Files.isDirectory(applicationPath)
                ? applicationPath
                : applicationPath.getParent();
        if (applicationDirectory == null) {
            return "";
        }

        Path configPath = applicationDirectory.resolve(CONFIG_FILE_NAME);
        if (!Files.exists(configPath)) {
            return "";
        }

        return findValue(Files.readAllLines(configPath, StandardCharsets.UTF_8), key);
    }

    /**
     * src/main/resources에 두어 JAR 내부에 포함된 기본 app.properties에서 특정 키의 값을 찾는다.
     *
     * @param key 읽어 올 설정 키
     * @return 일치하는 값, 없으면 빈 문자열
     * @throws IOException 설정 파일을 읽을 수 없는 경우
     */
    private static String readConfigValueFromClasspath(String key) throws IOException {
        try (InputStream stream = AppConfig.class.getResourceAsStream("/" + CONFIG_FILE_NAME)) {
            if (stream == null) {
                return "";
            }
            try (BufferedReader reader = new BufferedReader(new InputStreamReader(stream, StandardCharsets.UTF_8))) {
                return findValue(reader.lines().toList(), key);
            }
        }
    }

    /**
     * 현재 실행 중인 JAR 파일 또는 컴파일된 클래스 폴더의 경로를 반환한다.
     *
     * @return 실행 파일 또는 클래스 폴더 경로, 확인할 수 없으면 null
     */
    private static Path getApplicationPath() {
        try {
            var codeSource = AppConfig.class.getProtectionDomain().getCodeSource();
            if (codeSource == null) {
                return null;
            }
            return Paths.get(codeSource.getLocation().toURI());
        } catch (SecurityException | URISyntaxException exception) {
            return null;
        }
    }

    /**
     * 단순한 {@code key=value} 형식의 설정 목록에서 특정 키의 값을 찾는다.
     *
     * @param lines 설정 파일 내용
     * @param key   읽어 올 설정 키
     * @return 일치하는 값, 없으면 빈 문자열
     */
    private static String findValue(List<String> lines, String key) {
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
