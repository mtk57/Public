package sample;


public class Main {
    
    public static void main(String[] args) {
        String url = "http://example.com"; 
        String tricky = "/* not comment */ // still string";
        char slash = '/';
        
        System.out.println(url + tricky + slash); 
        
        
        System.out.println("finish");
        
        String keep = "// コメントではない";
        String block = "文中の /* コメント */ も保持";
    }
}
