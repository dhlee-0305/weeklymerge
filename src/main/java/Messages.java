import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Properties;

/**
 * messages.properties에서 사용자 노출 메시지를 읽어 제공한다.
 * 실행 폴더의 messages.properties를 우선 로드하고, 없으면 JAR에 포함된 기본 메시지를 클래스패스에서 로드한다.
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

    /**
     * 메시지 제공 유틸리티 클래스이므로 외부에서 인스턴스를 생성하지 못하게 막는다.
     */
    private Messages() {
    }

    /**
     * 실행 폴더에 별도로 배치된 messages.properties를 읽어 메시지 목록에 적재한다.
     * 배포 후 문구만 바꾸고 싶을 때 JAR 내부 기본 메시지보다 이 파일이 우선 적용된다.
     *
     * @return 실행 폴더에서 메시지 파일을 정상 로드했으면 true
     */
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

    /**
     * src/main/resources에 두어 JAR 내부에 포함된 기본 messages.properties를 읽어 메시지 목록에 적재한다.
     */
    private static void tryLoadFromClasspath() {
        try (InputStream stream = Messages.class.getResourceAsStream("/" + MESSAGES_FILE_NAME)) {
            if (stream != null) {
                PROPERTIES.load(new InputStreamReader(stream, StandardCharsets.UTF_8));
            }
        } catch (IOException e) {
            // 기본 메시지도 읽지 못하면 키 이름을 그대로 표시한다.
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
