package sample.flow;

public class MainWorkflow {
    private final Worker worker = new Worker();
    private final java.util.logging.Logger logger =
            java.util.logging.Logger.getLogger(MainWorkflow.class.getName());

    public void start() {
        initialize();
        worker.execute(); //@ STALE TAG
        Worker.performStatic();
        Helper helper = new Helper();
        helper.prepare();
        logInfo();
        overloadFunc1("123");
        overloadFunc1("123");  //@ C:\_git\Public\.NET\SimpleMethodCallListCreator\test\testdata\tagjumpembedding\src\sample\flow\MainWorkflow.java	private void overloadFunc1(String)	C:\_git\Public\.NET\SimpleMethodCallListCreator\test\testdata\tagjumpembedding\tagjump_method_list.tsv

        if(isCorrect("hoge")){
            logInfo();
        }
    }

    void initialize() {
        Helper.prepareStatic();
        worker.compute();
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
