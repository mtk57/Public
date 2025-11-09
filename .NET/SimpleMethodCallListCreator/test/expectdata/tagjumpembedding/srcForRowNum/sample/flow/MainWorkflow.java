package sample.flow;

public class MainWorkflow {
    private final Worker worker = new Worker();
    private final java.util.logging.Logger logger =
            java.util.logging.Logger.getLogger(MainWorkflow.class.getName());

    public void start() {
        initialize(); //@ C:\_git\Public\.NET\SimpleMethodCallListCreator\test\testdata\tagjumpembedding\srcForRowNum\sample\flow\MainWorkflow.java	33
        worker.execute(); //@ C:\_git\Public\.NET\SimpleMethodCallListCreator\test\testdata\tagjumpembedding\srcForRowNum\sample\flow\Worker.java	4
        Worker.performStatic(); //@ C:\_git\Public\.NET\SimpleMethodCallListCreator\test\testdata\tagjumpembedding\srcForRowNum\sample\flow\Worker.java	12
        Helper helper = new Helper();
        helper.prepare(); //@ C:\_git\Public\.NET\SimpleMethodCallListCreator\test\testdata\tagjumpembedding\srcForRowNum\sample\flow\Helper.java	4
        logInfo(); //@ C:\_git\Public\.NET\SimpleMethodCallListCreator\test\testdata\tagjumpembedding\srcForRowNum\sample\flow\MainWorkflow.java	38
        overloadFunc1("123"); //★ メソッド特定失敗
        overloadFunc1("123");  //@ C:\_git\Public\.NET\SimpleMethodCallListCreator\test\testdata\tagjumpembedding\srcForRowNum\sample\flow\MainWorkflow.java	50

        if(isCorrect("hoge")){ //@ C:\_git\Public\.NET\SimpleMethodCallListCreator\test\testdata\tagjumpembedding\srcForRowNum\sample\flow\MainWorkflow.java	42
            logInfo(); //@ C:\_git\Public\.NET\SimpleMethodCallListCreator\test\testdata\tagjumpembedding\srcForRowNum\sample\flow\MainWorkflow.java	38
        }

        if(isCorrect("hoge") != false){ //@ C:\_git\Public\.NET\SimpleMethodCallListCreator\test\testdata\tagjumpembedding\srcForRowNum\sample\flow\MainWorkflow.java	42
            logInfo(); //@ C:\_git\Public\.NET\SimpleMethodCallListCreator\test\testdata\tagjumpembedding\srcForRowNum\sample\flow\MainWorkflow.java	38
        }

        Worker w = new Worker();

        if(getWorkerString(w).Equal("hoge")){ //@ C:\_git\Public\.NET\SimpleMethodCallListCreator\test\testdata\tagjumpembedding\srcForRowNum\sample\flow\MainWorkflow.java	46
            logInfo(); //@ C:\_git\Public\.NET\SimpleMethodCallListCreator\test\testdata\tagjumpembedding\srcForRowNum\sample\flow\MainWorkflow.java	38
        }
    }

    void initialize() {
        Helper.prepareStatic(); //@ C:\_git\Public\.NET\SimpleMethodCallListCreator\test\testdata\tagjumpembedding\srcForRowNum\sample\flow\Helper.java	9
        worker.compute(); //@ C:\_git\Public\.NET\SimpleMethodCallListCreator\test\testdata\tagjumpembedding\srcForRowNum\sample\flow\Worker.java	8
    }

    private void logInfo() {
        logger.debug("start");
    }

    private boolean isCorrect(String prm) {
        return true;
    }

    private String getWorkerString(Worker w) {
        return w.ToString(); //@ C:\_git\Public\.NET\SimpleMethodCallListCreator\test\testdata\tagjumpembedding\srcForRowNum\sample\flow\Worker.java	16
    }

    private void overloadFunc1(String a) {
    }

    private void overloadFunc1(int a) {
    }
}
