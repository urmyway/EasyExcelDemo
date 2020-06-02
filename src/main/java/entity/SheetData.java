package entity;

import com.alibaba.excel.annotation.ExcelProperty;

public class SheetData {
    //员工姓名
    @ExcelProperty("员工姓名")
    private String userName;
    //员工工号
    @ExcelProperty("员工工号")
    private String workcode;
    //部门id
    @ExcelProperty("部门id")
    private int departmentid;

    public String getUserName() {
        return userName;
    }

    public void setUserName(String userName) {
        this.userName = userName;
    }

    public String getWorkcode() {
        return workcode;
    }

    public void setWorkcode(String workcode) {
        this.workcode = workcode;
    }

    public int getDepartmentid() {
        return departmentid;
    }

    public void setDepartmentid(int departmentid) {
        this.departmentid = departmentid;
    }
}
