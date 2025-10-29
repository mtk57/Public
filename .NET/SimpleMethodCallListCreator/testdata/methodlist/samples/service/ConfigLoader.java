package samples.service;

public final class ConfigLoader {

    private ConfigLoader() {
    }

    public static void load() {
        log("loading");
    }

    public static void log(int value) {
        System.out.println("value=" + value);
    }

    public static void log(String message) {
        System.out.println(message);
    }
}
