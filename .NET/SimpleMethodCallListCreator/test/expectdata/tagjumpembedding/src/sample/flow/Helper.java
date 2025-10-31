package sample.flow;

public class Helper {
    public void prepare() {
        Worker worker = new Worker();
        worker.compute();	//@ C:\_git\Public\.NET\SimpleMethodCallListCreator\test\testdata\tagjumpembedding\src\sample\flow\Worker.java	public void compute()	C:\_git\Public\.NET\SimpleMethodCallListCreator\test\testdata\tagjumpembedding\tagjump_method_list.tsv
    }

    public static void prepareStatic() {
        Worker.performStatic();	//@ C:\_git\Public\.NET\SimpleMethodCallListCreator\test\testdata\tagjumpembedding\src\sample\flow\Worker.java	public static void performStatic()	C:\_git\Public\.NET\SimpleMethodCallListCreator\test\testdata\tagjumpembedding\tagjump_method_list.tsv
    }
}
