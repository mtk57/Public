package sample;

import aaa.bbb;
import aaa.bbb.ccc;

// ファイル冒頭のコメント
public class Main {
    
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

    public int deleteFunc() {  //★DEL
    }

    private int nonDeleteFunc() {
    }

    public int deleteFunc2() {  //★DEL
    }

    public static void sub(String[] args) {
        
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


}
