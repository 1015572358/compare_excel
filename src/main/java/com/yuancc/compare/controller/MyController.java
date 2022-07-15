package com.yuancc.compare.controller;

import com.yuancc.compare.handleexcel.CompareExcel;
import org.apache.log4j.Logger;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.servlet.http.HttpServletResponse;

@Controller
public class MyController {
    private static Logger log = Logger.getLogger(MyController.class);
    @RequestMapping("/getLog")
    public String getLog(HttpServletResponse response){
//        CompareExcel.getCompareExcel(response);

        return "mylog";
    }
}
