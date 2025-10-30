package tagsample;

public class Worker {
    public static void doWork() {
        logInternal();
    }

    public void run() {
        logInternal();
    }

    private static void logInternal() {
        System.out.println("working");
    }
}
