package samples.util;

public class MultiLineFormatter {

    public String compose(
        int id,
        String name,
        boolean enabled) {
        return String.format("%d-%s-%s", id, name, enabled);
    }

    public static void invokeRepeated(
        Runnable action,
        int repeatCount) {
        for (int i = 0; i < repeatCount; i++) {
            action.run();
        }
    }
}
