public class SampleData {

    private Long id;
    private Long seq;
    private String name;

    public SampleData() {
    }

    public SampleData(Long id, Long seq, String name) {
        this.id = id;
        this.seq = seq;
        this.name = name;
    }

    public Long getId() {
        return id;
    }

    public void setId(Long id) {
        this.id = id;
    }

    public Long getSeq() {
        return seq;
    }

    public void setSeq(Long seq) {
        this.seq = seq;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    @Override
    public String toString() {
        return "SampleData{" + "id=" + id + ", " + "seq=" + seq + ", " + "name='" + name + "'" + "}";
    }
}