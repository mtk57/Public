package tagsample;

public class TagRoot {
    public void start() {
        prepare();
        Helper.execute();
        Utility util = new Utility();
        util.log("done");
    }

    private void prepare() {
        Worker.doWork();
    }
}
