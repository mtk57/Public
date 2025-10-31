package sample.flow;

public class Worker {
    public void execute() {
        compute();
    }

    public void compute() {
        System.out.println("computing");
    }

    public static void performStatic() {
        Utility.touch();
    }
}
