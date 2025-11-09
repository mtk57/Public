package sample.flow;

public class Worker {
    public void execute() {
        compute(); //@ C:\_git\Public\.NET\SimpleMethodCallListCreator\test\testdata\tagjumpembedding\srcForRowNum\sample\flow\Worker.java	
    }

    public void compute() {
        System.out.println("computing");
    }

    public static void performStatic() {
        Utility.touch();
    }

    public String ToString(Object o) {
        o.ToString(); //@ C:\_git\Public\.NET\SimpleMethodCallListCreator\test\testdata\tagjumpembedding\srcForRowNum\sample\flow\Worker.java	
    }
}
