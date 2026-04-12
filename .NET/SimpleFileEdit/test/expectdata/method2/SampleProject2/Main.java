package sample;

// ファイル冒頭のコメント
public class Main {
    
    /**
     * 削除対象メソッドのJavadocコメント
     * @return 値
     */
    public int deleteFunc() {  //★DEL

        if (aaaa == null) {
            throw new Exception();
        }

        for (int i = 1; i <= maxNum; i++) {
            
            // 4. if文（偶数か奇数かを判定）
            if (i % 2 == 0) {
                System.out.println(i + " は偶数です！");
            } else {
                System.out.println(i + " は奇数です。");
            }
        }

    }

    /*
     * 削除対象のブロックコメント
     */
    public int deleteFunc2() {  //★DEL
    }

    /**
     * 制御構文を多用するメソッド
     * 否定モードで本体が消えないことを確認
     */
    public void controlFlowMethod() {  //★DEL

        // if文
        if (flag) {
            System.out.println("flag is true");
        } else {
            System.out.println("flag is false");
        }

        // for文
        for (int i = 0; i < 10; i++) {
            System.out.println(i);
        }

        // while文
        while (condition) {
            doSomething();
        }

        // do-while文
        do {
            doSomething();
        } while (condition);

        // switch文
        switch (value) {
            case 1:
                break;
            default:
                break;
        }

        // try-catch文
        try {
            riskyOperation();
        } catch (Exception e) {
            handleError(e);
        }

        // throw文
        if (input == null) {
            throw new IllegalArgumentException("input is null");
        }

        // return文
        return;
    }

    // フィールド初期化（メソッドと誤検出しないこと）
    String[] data = new String[] { "aa", "bb" };

}
