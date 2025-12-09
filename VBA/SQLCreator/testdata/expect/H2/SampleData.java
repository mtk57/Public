public class SampleData {

    private Long ID;
    private Long SEQ;
    private String NAME;

    public SampleData() {
    }

    public SampleData(Long ID, Long SEQ, String NAME) {
        this.ID = ID;
        this.SEQ = SEQ;
        this.NAME = NAME;
    }

    public Long getID() {
        return ID;
    }

    public void setID(Long ID) {
        this.ID = ID;
    }

    public Long getSEQ() {
        return SEQ;
    }

    public void setSEQ(Long SEQ) {
        this.SEQ = SEQ;
    }

    public String getNAME() {
        return NAME;
    }

    public void setNAME(String NAME) {
        this.NAME = NAME;
    }

    @Override
    public String toString() {
        return "SampleData{" + "ID=" + ID + ", " + "SEQ=" + SEQ + ", " + "NAME='" + NAME + "'" + "}";
    }
}