package entity;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.format.DateTimeFormat;
import lombok.Data;

import java.util.Date;
@Data
public class DemoData {
    /**
     * 强制读取第三个 这里不建议 index 和 name 同时用，要么一个对象只用index，要么一个对象只用name去匹配
     */
    @ExcelProperty(index = 0)
    private String name;
    @ExcelProperty(index = 1)
    private Date date;
    /**
     * 用名字去匹配，这里需要注意，如果名字重复，会导致只有一个字段读取到数据
     */
    @ExcelProperty("地址")
    private String address;
}
