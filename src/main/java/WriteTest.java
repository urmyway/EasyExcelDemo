import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.write.metadata.WriteSheet;
import entity.*;
import org.junit.Test;
import util.TestFileUtil;

import java.util.*;

public class WriteTest {
    //通用数据生成
    private List<Demo> data() {
        List<Demo> list = new ArrayList<Demo>();
        for (int i = 0; i < 10; i++) {
            Demo data = new Demo();
            data.setString("字符串" + i);
            data.setDate(new Date());
            data.setDoubleData(0.56);
            list.add(data);
        }
        return list;
    }

    /**
     * 最简单的写
     * <p>1. 创建excel对应的实体对象 参照{@link DemoData}
     * <p>2. 直接写即可
     */
    @Test
    public void simpleWrite() {
        // 写法1
        String fileName = "D:\\test\\" + "simpleWrite" + System.currentTimeMillis() + ".xlsx";
        // 这里 需要指定写用哪个class去写，然后写到第一个sheet，名字为模板 然后文件流会自动关闭
        // 如果这里想使用03 则 传入excelType参数即可
        EasyExcel.write(fileName, Demo.class).sheet("模板").doWrite(data());

/*        // 写法2
        fileName = TestFileUtil.getPath() + "simpleWrite" + System.currentTimeMillis() + ".xlsx";
        // 这里 需要指定写用哪个class去写
        ExcelWriter excelWriter = EasyExcel.write(fileName, DemoData.class).build();
        WriteSheet writeSheet = EasyExcel.writerSheet("模板").build();
        excelWriter.write(data(), writeSheet);
        // 千万别忘记finish 会帮忙关闭流
        excelWriter.finish();*/
    }

    /**
     * 根据参数只导出指定列
     * <p>
     * 1. 创建excel对应的实体对象 参照{@link DemoData}
     * <p>
     * 2. 根据自己或者排除自己需要的列
     * <p>
     * 3. 直接写即可
     *
     * @since 2.1.1
     */
    @Test
    public void excludeOrIncludeWrite() {
        String fileName = "D:\\test\\" + "excludeOrIncludeWrite" + System.currentTimeMillis() + ".xlsx";

        // 根据用户传入字段 假设我们要忽略 date
        Set<String> excludeColumnFiledNames = new HashSet<String>();
        excludeColumnFiledNames.add("date");
        // 这里 需要指定写用哪个class去写，然后写到第一个sheet，名字为模板 然后文件流会自动关闭
        EasyExcel.write(fileName, Demo.class).excludeColumnFiledNames(excludeColumnFiledNames).sheet("模板")
                .doWrite(data());

        fileName = "D:\\test\\" + "excludeOrIncludeWrite" + System.currentTimeMillis() + ".xlsx";
        // 根据用户传入字段 假设我们只要导出 date
        Set<String> includeColumnFiledNames = new HashSet<String>();
        includeColumnFiledNames.add("date");
        // 这里 需要指定写用哪个class去写，然后写到第一个sheet，名字为模板 然后文件流会自动关闭
        EasyExcel.write(fileName, Demo.class).includeColumnFiledNames(includeColumnFiledNames).sheet("模板")
                .doWrite(data());
    }

    /**
     * 指定写入的列
     * <p>1. 创建excel对应的实体对象 参照{@link IndexData}
     * <p>2. 使用{@link @ExcelProperty}注解指定写入的列
     * <p>3. 直接写即可
     */
    @Test
    public void indexWrite() {
        String fileName = "D:\\test\\" + "indexWrite" + System.currentTimeMillis() + ".xlsx";
        // 这里 需要指定写用哪个class去写，然后写到第一个sheet，名字为模板 然后文件流会自动关闭
        EasyExcel.write(fileName, IndexData.class).sheet("模板").doWrite(data());
    }

    /**
     * 复杂头写入
     * <p>1. 创建excel对应的实体对象 参照{@link ComplexHeadData}
     * <p>2. 使用{@link @ExcelProperty}注解指定复杂的头
     * <p>3. 直接写即可
     */
    @Test
    public void complexHeadWrite() {
        String fileName = "D:\\test\\" + "complexHeadWrite" + System.currentTimeMillis() + ".xlsx";
        // 这里 需要指定写用哪个class去写，然后写到第一个sheet，名字为模板 然后文件流会自动关闭
        EasyExcel.write(fileName, ComplexHeadData.class).sheet("模板").doWrite(data());
    }

    /**
     * 重复多次写入
     * <p>
     * 1. 创建excel对应的实体对象 参照{@link ComplexHeadData}
     * <p>
     * 2. 使用{@link @ExcelProperty}注解指定复杂的头
     * <p>
     * 3. 直接调用二次写入即可
     */
    @Test
    public void repeatedWrite() {
        // 方法1 如果写到同一个sheet
        String fileName = "D:\\test\\" + "repeatedWrite" + System.currentTimeMillis() + ".xlsx";
        // 这里 需要指定写用哪个class去写
        ExcelWriter excelWriter = EasyExcel.write(fileName, Demo.class).build();
        // 这里注意 如果同一个sheet只要创建一次
        WriteSheet writeSheet = EasyExcel.writerSheet("模板").build();
        // 去调用写入,这里我调用了五次，实际使用时根据数据库分页的总的页数来
        for (int i = 0; i < 5; i++) {
            // 分页去数据库查询数据 这里可以去数据库查询每一页的数据
            List<Demo> data = data();
            excelWriter.write(data, writeSheet);
        }
        /// 千万别忘记finish 会帮忙关闭流
        excelWriter.finish();

        // 方法2 如果写到不同的sheet 同一个对象
        fileName = "D:\\test\\" + "repeatedWrite" + System.currentTimeMillis() + ".xlsx";
        // 这里 指定文件
        excelWriter = EasyExcel.write(fileName, Demo.class).build();
        // 去调用写入,这里我调用了五次，实际使用时根据数据库分页的总的页数来。这里最终会写到5个sheet里面
        for (int i = 0; i < 5; i++) {
            // 每次都要创建writeSheet 这里注意必须指定sheetNo 而且sheetName必须不一样
            writeSheet = EasyExcel.writerSheet(i, "模板" + i).build();
            // 分页去数据库查询数据 这里可以去数据库查询每一页的数据
            List<Demo> data = data();
            excelWriter.write(data, writeSheet);
        }
        /// 千万别忘记finish 会帮忙关闭流
        excelWriter.finish();

        // 方法3 如果写到不同的sheet 不同的对象
        fileName = "D:\\test\\" + "repeatedWrite" + System.currentTimeMillis() + ".xlsx";
        // 这里 指定文件
        excelWriter = EasyExcel.write(fileName).build();
        // 去调用写入,这里我调用了五次，实际使用时根据数据库分页的总的页数来。这里最终会写到5个sheet里面
        for (int i = 0; i < 5; i++) {
            // 每次都要创建writeSheet 这里注意必须指定sheetNo 而且sheetName必须不一样。这里注意Demo.class 可以每次都变，我这里为了方便 所以用的同一个class 实际上可以一直变
            writeSheet = EasyExcel.writerSheet(i, "模板" + i).head(Demo.class).build();
            // 分页去数据库查询数据 这里可以去数据库查询每一页的数据
            List<Demo> data = data();
            excelWriter.write(data, writeSheet);
        }
        /// 千万别忘记finish 会帮忙关闭流
        excelWriter.finish();
    }

    /**
     * 日期、数字或者自定义格式转换
     * <p>1. 创建excel对应的实体对象 参照{@link ConverterData}
     * <p>2. 使用{@link @ExcelProperty}配合使用注解{@link @DateTimeFormat}、{@link @NumberFormat}或者自定义注解
     * <p>3. 直接写即可
     */
    @Test
    public void converterWrite() {
        String fileName = "D:\\test\\" + "converterWrite" + System.currentTimeMillis() + ".xlsx";
        // 这里 需要指定写用哪个class去写，然后写到第一个sheet，名字为模板 然后文件流会自动关闭
        EasyExcel.write(fileName, ConverterData.class).sheet("模板").doWrite(data());
    }
}
