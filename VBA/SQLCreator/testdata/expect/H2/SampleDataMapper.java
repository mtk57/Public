import java.util.List;
import org.apache.ibatis.annotations.Mapper;

import SampleData;

@Mapper
public interface SampleDataMapper {

    List<SampleData> findAll();

    int insert(SampleData data);

    int update(SampleData data);
}