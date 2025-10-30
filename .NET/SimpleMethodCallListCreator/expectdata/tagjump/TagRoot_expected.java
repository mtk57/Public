package tagsample;

public class TagRoot {
    public void start() {
        prepare();//@ C:\_git\Public\.NET\SimpleMethodCallListCreator\testdata\tagjump\src\TagRoot.java	private void prepare()
        Helper.execute();//@ C:\_git\Public\.NET\SimpleMethodCallListCreator\testdata\tagjump\src\Helper.java	public static void execute()
        Utility util = new Utility();
        util.log("done");//@ C:\_git\Public\.NET\SimpleMethodCallListCreator\testdata\tagjump\src\Utility.java	public void log(String)
    }

    private void prepare() {
        Worker.doWork();//@ C:\_git\Public\.NET\SimpleMethodCallListCreator\testdata\tagjump\src\Worker.java	public static void doWork()
    }
}
