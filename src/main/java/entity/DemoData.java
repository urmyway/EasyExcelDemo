package entity;

import com.alibaba.excel.annotation.format.DateTimeFormat;
import lombok.Data;

import java.util.Date;
@Data
public class DemoData {
    private String name;
    @DateTimeFormat("yyyy-mm-dd")
    private Date date;
    private String address;
}
