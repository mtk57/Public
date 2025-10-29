package samples.util;

import java.util.Locale;

public final class Converters {

    private Converters() {
    }

    public static String toDisplayName(String source) {
        if (source == null) {
            return "";
        }
        return source.toUpperCase(Locale.ROOT);
    }

    public static String join(String separator, String[] values) {
        if (values == null) {
            return "";
        }
        StringBuilder builder = new StringBuilder();
        for (int i = 0; i < values.length; i++) {
            if (i > 0) {
                builder.append(separator);
            }
            builder.append(values[i]);
        }
        return builder.toString();
    }
}
