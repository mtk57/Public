package sample;

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
}
