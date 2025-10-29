package samples.service;

import java.util.Collections;
import java.util.List;

public class TaskService {

    public void execute(int count) {
        for (int i = 0; i < count; i++) {
            processStep(i);
        }
    }

    private void processStep(int index) {
        Helper helper = new Helper();
        helper.handle(index);
    }

    public static void prepareEnvironment() {
        ConfigLoader.load();
    }

    public List<String> summarize(List<String> inputs) throws IllegalArgumentException {
        if (inputs == null) {
            throw new IllegalArgumentException("inputs is null");
        }
        return Collections.unmodifiableList(inputs);
    }

    private static class Helper {
        void handle(int value) {
            ConfigLoader.log(value);
        }
    }
}
