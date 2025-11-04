package sample.flow;

public class MainWorkflow {
    private final Worker worker = new Worker();
    private final java.util.logging.Logger logger =
            java.util.logging.Logger.getLogger(MainWorkflow.class.getName());

    public void start() {
        initialize();//@ C:\_git\Public\.NET\SimpleMethodCallListCreator\test\testdata\tagjumpembedding\src\sample\flow\MainWorkflow.java	void initialize()	C:\_git\Public\.NET\SimpleMethodCallListCreator\test\testdata\tagjumpembedding\tagjump_method_list.tsv
        worker.execute(); //@ C:\_git\Public\.NET\SimpleMethodCallListCreator\test\testdata\tagjumpembedding\src\sample\flow\Worker.java	public void execute()	C:\_git\Public\.NET\SimpleMethodCallListCreator\test\testdata\tagjumpembedding\tagjump_method_list.tsv
        Worker.performStatic();//@ C:\_git\Public\.NET\SimpleMethodCallListCreator\test\testdata\tagjumpembedding\src\sample\flow\Worker.java	public static void performStatic()	C:\_git\Public\.NET\SimpleMethodCallListCreator\test\testdata\tagjumpembedding\tagjump_method_list.tsv
        Helper helper = new Helper();
        helper.prepare();//@ C:\_git\Public\.NET\SimpleMethodCallListCreator\test\testdata\tagjumpembedding\src\sample\flow\Helper.java	public void prepare()	C:\_git\Public\.NET\SimpleMethodCallListCreator\test\testdata\tagjumpembedding\tagjump_method_list.tsv
        logInfo();//@ C:\_git\Public\.NET\SimpleMethodCallListCreator\test\testdata\tagjumpembedding\src\sample\flow\MainWorkflow.java	private void logInfo()	C:\_git\Public\.NET\SimpleMethodCallListCreator\test\testdata\tagjumpembedding\tagjump_method_list.tsv
        overloadFunc1("123"); //★ メソッド特定失敗
        overloadFunc1("123");  //@ C:\_git\Public\.NET\SimpleMethodCallListCreator\test\testdata\tagjumpembedding\src\sample\flow\MainWorkflow.java	private void overloadFunc1(String)	C:\_git\Public\.NET\SimpleMethodCallListCreator\test\testdata\tagjumpembedding\tagjump_method_list.tsv

        if(isCorrect("hoge")){//@ C:\_git\Public\.NET\SimpleMethodCallListCreator\test\testdata\tagjumpembedding\src\sample\flow\MainWorkflow.java	private boolean isCorrect(String)	C:\_git\Public\.NET\SimpleMethodCallListCreator\test\testdata\tagjumpembedding\tagjump_method_list.tsv
            logInfo();//@ C:\_git\Public\.NET\SimpleMethodCallListCreator\test\testdata\tagjumpembedding\src\sample\flow\MainWorkflow.java	private void logInfo()	C:\_git\Public\.NET\SimpleMethodCallListCreator\test\testdata\tagjumpembedding\tagjump_method_list.tsv
        }
    }

    void initialize() {
        Helper.prepareStatic();//@ C:\_git\Public\.NET\SimpleMethodCallListCreator\test\testdata\tagjumpembedding\src\sample\flow\Helper.java	public static void prepareStatic()	C:\_git\Public\.NET\SimpleMethodCallListCreator\test\testdata\tagjumpembedding\tagjump_method_list.tsv
        worker.compute();//@ C:\_git\Public\.NET\SimpleMethodCallListCreator\test\testdata\tagjumpembedding\src\sample\flow\Worker.java	public void compute()	C:\_git\Public\.NET\SimpleMethodCallListCreator\test\testdata\tagjumpembedding\tagjump_method_list.tsv
    }

    private void logInfo() {
        logger.debug("start");
    }

    private boolean isCorrect(String prm) {
        return true;
    }

    private void overloadFunc1(String a) {
    }

    private void overloadFunc1(int a) {
    }
}
