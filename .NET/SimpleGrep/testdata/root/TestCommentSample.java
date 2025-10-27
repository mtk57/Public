package testdata;

// コメント
//コメント
/* コメント */
/*
コメント
*/


public class TestCommentSample {

    // GrepTarget should be ignored because it is in a line comment.
    /*
     * GrepTarget should be ignored even in a block comment.
     */
    public void printMessage() {
        String message = "GrepTarget appears in a string literal and should be found.";
        System.out.println(message);

        String multiLine = "Line1"
                + "GrepTarget in concatenated string should be found.";
        System.out.println(multiLine);

        char ch = 'G';
        char c2 = 'p'; // GrepTarget partial should not match unless literal search.
        System.out.println("" + ch + c2);
    }

    public void anotherMethod() {
        String example = "No keyword here.";
        System.out.println(example);
    }
}
