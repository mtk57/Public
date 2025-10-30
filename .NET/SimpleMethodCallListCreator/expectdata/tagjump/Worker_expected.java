package tagsample;

public class Worker {
    public static void doWork() {
        logInternal();//@ C:\_git\Public\.NET\SimpleMethodCallListCreator\testdata\tagjump\src\Worker.java	private static void logInternal()
    }

    public void run() {
        logInternal();//@ C:\_git\Public\.NET\SimpleMethodCallListCreator\testdata\tagjump\src\Worker.java	private static void logInternal()
    }

    private static void logInternal() {
        System.out.println("working");
    }
}
