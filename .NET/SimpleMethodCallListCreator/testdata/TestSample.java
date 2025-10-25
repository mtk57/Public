package samples;

public class FooController {

    public void execute() {
        prepare();
        Service service = new Service();
        service.process("data", 10);

        Utility.log("execute finished");
    }

    private void prepare() {
        String message = buildMessage("ready");
        // comments should be ignored by analyzer
        System.out.println(message);
        Utility.logFormatted(
        	"prepared: %s",
        	message
        );
    }

    private String buildMessage(String text) {
        return "Message:" + text;
    }
}

class Service {

    public void process(String name, int count) {
        validate(name);
        Helper helper = new Helper();
        helper.runTask(count);
    }

    private void validate(String value) {
        if (value == null || value.isEmpty()) {
            throw new IllegalArgumentException("value is required");
        }
    }
}

class Helper {

    public void runTask(int times) {
        for (int i = 0; i < times; i++) {
            step(i);
        }
    }

    private void step(int index) {
        Utility.log("step:" + index);
    }
}

class Utility {

    public static void log(String message) {
        System.out.println(message);
    }

    public static void logFormatted(String format, String value) {
        log(String.format(format, value));
    }

    public static void check(
		String format,
		String value) throws Exception {

        Utility.log("step:" + index);
    }

}
