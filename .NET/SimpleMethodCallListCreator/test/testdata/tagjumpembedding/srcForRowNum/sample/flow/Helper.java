package sample.flow;

public class Helper {
    public void prepare() {
        Worker worker = new Worker();
        worker.compute(); //@ legacy-tag
    }

    public static void prepareStatic() {
        Worker.performStatic();
    }
}
