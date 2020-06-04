import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.metadata.fill.FillConfig;
import entity.Demo;
import entity.FillData;
import org.junit.Test;
import util.TestFileUtil;

import java.io.File;
import java.util.*;

public class FillTest {
    /**
     * 最简单的填充
     *
     * @since 2.1.1
     */
    @Test
    public void simpleFill() {
        // 模板注意 用{} 来表示你要用的变量 如果本来就有"{","}" 特殊字符 用"\{","\}"代替
        String templateFileName = "D:\\test\\simple.xlsx";
        // 方案1 根据对象填充
        String fileName = "D:\\test\\" + "simpleFill" + System.currentTimeMillis() + ".xlsx";
        // 这里 会填充到第一个sheet， 然后文件流会自动关闭
        FillData fillData = new FillData();
        fillData.setName("张三");
        fillData.setNumber(5.2);
        EasyExcel.write(fileName).withTemplate(templateFileName).sheet().doFill(fillData);
        // 方案2 根据Map填充
        fileName = "D:\\test\\" + "simpleFill" + System.currentTimeMillis() + ".xlsx";
        // 这里 会填充到第一个sheet， 然后文件流会自动关闭
        Map<String, Object> map = new HashMap<String, Object>();
        map.put("name", "张三");
        map.put("number", 5.2);
        EasyExcel.write(fileName).withTemplate(templateFileName).sheet().doFill(map);
    }

    /**
     * 填充列表
     *
     * @since 2.1.1
     */
    @Test
    public void listFill() {
        // 模板注意 用{} 来表示你要用的变量 如果本来就有"{","}" 特殊字符 用"\{","\}"代替
        // 填充list 的时候还要注意 模板中{.} 多了个点 表示list
        String templateFileName = "D:\\test\\list.xlsx";

        // 方案1 一下子全部放到内存里面 并填充
        String fileName = "D:\\test\\" + "listFill" + System.currentTimeMillis() + ".xlsx";
        // 这里 会填充到第一个sheet， 然后文件流会自动关闭
        EasyExcel.write(fileName).withTemplate(templateFileName).sheet().doFill(data());

        // 方案2 分多次 填充 会使用文件缓存（省内存）
        fileName = "D:\\test\\" + "listFill" + System.currentTimeMillis() + ".xlsx";
        ExcelWriter excelWriter = EasyExcel.write(fileName).withTemplate(templateFileName).build();
        WriteSheet writeSheet = EasyExcel.writerSheet().build();
        excelWriter.fill(data(), writeSheet);
        excelWriter.fill(data(), writeSheet);
        // 千万别忘记关闭流
        excelWriter.finish();
    }

    //通用数据生成
    private List<FillData> data() {
        List<FillData> list = new ArrayList<FillData>();
        for (int i = 0; i < 10; i++) {
            FillData fillData = new FillData();
            list.add(fillData);
            fillData.setName("张三");
            fillData.setNumber(5.2);
        }
        return list;
    }

    /**
     * 复杂的填充
     *
     * @since 2.1.1
     */
    @Test
    public void complexFill() {
        // 模板注意 用{} 来表示你要用的变量 如果本来就有"{","}" 特殊字符 用"\{","\}"代替
        // {} 代表普通变量 {.} 代表是list的变量
        String templateFileName = "D:\\test\\complex.xlsx";

        String fileName = "D:\\test\\" + "complexFill" + System.currentTimeMillis() + ".xlsx";
        ExcelWriter excelWriter = EasyExcel.write(fileName).withTemplate(templateFileName).build();
        WriteSheet writeSheet = EasyExcel.writerSheet().build();
        // 这里注意 入参用了forceNewRow 代表在写入list的时候不管list下面有没有空行 都会创建一行，然后下面的数据往后移动。默认 是false，会直接使用下一行，如果没有则创建。
        // forceNewRow 如果设置了true,有个缺点 就是他会把所有的数据都放到内存了，所以慎用
        // 简单的说 如果你的模板有list,且list不是最后一行，下面还有数据需要填充 就必须设置 forceNewRow=true 但是这个就会把所有数据放到内存 会很耗内存
        // 如果数据量大 list不是最后一行 参照下一个 √
        FillConfig fillConfig = FillConfig.builder().forceNewRow(Boolean.TRUE).build();
        excelWriter.fill(data(), fillConfig, writeSheet);
        excelWriter.fill(data(), fillConfig, writeSheet);
        Map<String, Object> map = new HashMap<String, Object>();
        map.put("date", "2019年10月9日13:28:28");
        map.put("total", 1000);
        //分多次填充
        excelWriter.fill(map, writeSheet);
        excelWriter.finish();
    }
}
