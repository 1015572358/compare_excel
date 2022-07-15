package com.yuancc.compare.domain;

import com.yuancc.compare.utils.ExcelColumn;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.Date;

/**
 * @author ycc
 */
@Data
@NoArgsConstructor
@AllArgsConstructor
public class Wgtxz {
    @ExcelColumn(value = "序号",col = 1)
    private String dno;
    @ExcelColumn(value = "所属网格",col = 2)
    private String sswg;
    @ExcelColumn(value = "姓名",col = 3)
    private String xm;
    @ExcelColumn(value = "身份证号",col = 4)
    private String sfzh;
    @ExcelColumn(value = "住址",col = 5)
    private String address;
    @ExcelColumn(value = "手机号",col = 6)
    private String phone;
    @ExcelColumn(value = "人员类型",col = 7)
    private String rylx;
    @ExcelColumn(value = "是否出东兴",col = 8)
    private String sfcdx;
    @ExcelColumn(value = "注册时间",col = 9)
    private String zcsj;
    @ExcelColumn(value = "备注",col = 10)
    private String bz;
    @ExcelColumn(value = "通行证状态",col = 11)
    private String txzzt;
    @ExcelColumn(value = "申请类型",col = 12)
    private String sqlx;

}
