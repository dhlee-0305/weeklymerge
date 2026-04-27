import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Properties;

/**
 * messages.properties에서 사용자 노출 메시지를 읽어 제공한다.
 * 현재 작업 폴더의 messages.properties를 우선 로드하고, 없으면 클래스패스에서 로드한다.
 * 키가 없으면 키 이름을 그대로 반환한다.
 */
public final class Messages {
    private static final String MESSAGES_FILE_NAME = "messages.properties";
    private static final Properties PROPERTIES = new Properties();

    static {
        boolean loaded = tryLoadFromWorkingDirectory();
        if (!loaded) {
            tryLoadFromClasspath();
        }
    }

    private Messages() {
    }

    private static boolean tryLoadFromWorkingDirectory() {
        var path = Paths.get(MESSAGES_FILE_NAME);
        if (!Files.exists(path)) {
            return false;
        }
        try (var reader = Files.newBufferedReader(path, StandardCharsets.UTF_8)) {
            PROPERTIES.load(reader);
            return true;
        } catch (IOException e) {
            return false;
        }
    }

    private static void tryLoadFromClasspath() {
        try (InputStream stream = Messages.class.getResourceAsStream("/" + MESSAGES_FILE_NAME)) {
            if (stream != null) {
                PROPERTIES.load(new InputStreamReader(stream, StandardCharsets.UTF_8));
            }
        } catch (IOException e) {
            // 어디서도 파일을 읽지 못하면 키 이름을 그대로 표시한다.
        }
    }

    /**
     * 키에 대응하는 메시지를 반환한다. 없으면 키 이름을 반환한다.
     *
     * @param key 메시지 키
     * @return 메시지 문자열
     */
    public static String get(String key) {
        return PROPERTIES.getProperty(key, key);
    }

    /**
     * 키에 대응하는 메시지의 {@code {0}}, {@code {1}} 등 자리표시자를 인자로 순서대로 치환한다.
     *
     * @param key  메시지 키
     * @param args 치환할 인자 목록
     * @return 인자가 치환된 메시지 문자열
     */
    public static String format(String key, Object... args) {
        String pattern = get(key);
        for (int i = 0; i < args.length; i++) {
            pattern = pattern.replace("{" + i + "}", args[i] != null ? args[i].toString() : "");
        }
        return pattern;
    }
}
