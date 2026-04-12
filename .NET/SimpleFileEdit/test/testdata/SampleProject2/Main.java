package sample;

import aaa.bbb;
import aaa.bbb.ccc;

// ファイル冒頭のコメント
public class Main {
    

    /**
     *
     */
    public static void main(String[] args) {
        String url = "http://example.com"; // URLコメント
        String tricky = "/* not comment */ // still string";
        char slash = '/';
        
        System.out.println(url + tricky + slash); /* inline block */
        // 単純なコメント
        /*
            複数行
            コメント
        */
        System.out.println("finish");
        
        String keep = "// コメントではない";
        String block = "文中の /* コメント */ も保持";
    }

    /**
     *
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

    /**
     *
     */
    private int nonDeleteFunc() {
    }

    /**
     *
     */
    public int deleteFunc2() {  //★DEL
    }

    /**
     *
     */
    public static void sub(String[] args) {
        

        if (aaaa == null) {
            throw new Exception();
        }


        // 1. 変数の宣言と代入（ここで数字を1つ覚えさせます）
        int maxNum = 5;
        
        // 2. 画面に出力（最初に説明を表示）
        System.out.println("1から" + maxNum + "までの数字をチェックします！");
        
        // 3. for文（1からmaxNumまで順番に繰り返す）
        for (int i = 1; i <= maxNum; i++) {
            
            // 4. if文（偶数か奇数かを判定）
            if (i % 2 == 0) {
                System.out.println(i + " は偶数です！");
            } else {
                System.out.println(i + " は奇数です。");
            }
        }
        
        // 全部終わったらメッセージ
        System.out.println("チェック完了！");
    }

    /**
     * 制御構文を多用するメソッド（否定モードで本体が消えないことを確認）
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

    /**
     * 制御構文を多用するメソッド（削除対象外）
     */
    public void controlFlowMethodNonDel() {

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

}
