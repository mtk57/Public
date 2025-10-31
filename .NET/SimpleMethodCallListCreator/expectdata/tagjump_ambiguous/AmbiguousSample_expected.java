package tagjump.ambiguous;

public class AmbiguousSample {
    public void start() {
        target(); //★ メソッド特定失敗
    }

    public void target() {
    }
}
