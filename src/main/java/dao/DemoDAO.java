package dao;

import entity.DemoData;
import entity.SheetData;

import java.util.List;

public class DemoDAO {
    public void save(List<DemoData> list) {
        // 如果是mybatis,尽量别直接调用多次insert,自己写一个mapper里面新增一个方法batchInsert,所有数据一次性插入
    }

    public void save1(List<SheetData> list){
        //
    }
}
