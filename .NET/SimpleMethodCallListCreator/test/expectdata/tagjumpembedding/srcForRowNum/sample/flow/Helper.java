package sample.flow;

public class Helper {
    public void prepare() {
        Worker worker = new Worker();
        worker.compute(); //@ C:\_git\Public\.NET\SimpleMethodCallListCreator\test\testdata\tagjumpembedding\srcForRowNum\sample\flow\Worker.java	8
    }

    public static void prepareStatic() {
        Worker.performStatic(); //@ C:\_git\Public\.NET\SimpleMethodCallListCreator\test\testdata\tagjumpembedding\srcForRowNum\sample\flow\Worker.java	12
    }
}
