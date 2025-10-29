package samples.app;

import java.util.ArrayList;
import java.util.List;
import samples.service.TaskService;
import samples.util.Converters;

public class MainEntry {

    private final TaskService service = new TaskService();

    public void run(String[] args) {
        int count = args != null ? args.length : 0;
        service.execute(count);
    }

    public static void bootstrap() {
        TaskService.prepareEnvironment();
    }

    protected List<String> buildItems(List<String> sources) {
        List<String> results = new ArrayList<>();
        for (String item : sources) {
            results.add(Converters.toDisplayName(item));
        }
        return results;
    }

    private String formatMessage(final String source) {
        return String.format("source:%s", source);
    }
}
