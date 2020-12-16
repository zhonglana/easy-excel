package com.lz.easyexcel;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.EasyExcelFactory;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.write.builder.ExcelWriterBuilder;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.style.HorizontalCellStyleStrategy;
import com.lz.easyexcel.domain.ExportVo;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.util.IOUtils;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.metadata.Table;
import org.springframework.stereotype.Controller;
import org.springframework.util.ResourceUtils;
import org.springframework.util.StringUtils;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.util.ArrayList;
import java.util.List;

@Slf4j
@Controller
public class TestController {


    @RequestMapping(value = "/export", method = RequestMethod.GET, produces = "application/json")
    public void export(HttpServletResponse response) throws IOException {
        String fileName = System.currentTimeMillis() + ".xlsx";
        response.setHeader("Content-disposition", "attachment;filename=" + fileName);
        OutputStream fileOutputStream = response.getOutputStream();
        try {
            //查询数据
            List<ExportVo> list = new ArrayList<>();
            list.add(new ExportVo("123", "lz"));
            //生成Excel
            String filePath = ResourceUtils.getURL("classPath:").getPath().concat(fileName);
            EasyExcel.write(fileOutputStream, ExportVo.class).sheet("sheet").doWrite(list);

        } catch (Exception e) {
            log.error("export failead:", e);
        } finally {
            //关闭流
            if (fileOutputStream != null) {
                IOUtils.closeQuietly(fileOutputStream);
            }
            //删除本地文件
            if (!StringUtils.isEmpty(fileName)) {
                File file = new File(fileName);
                if (file.exists()) {
                    file.delete();
                }
            }
        }
    }



    @RequestMapping(value = "/export2", method = RequestMethod.GET, produces = "application/json")
    public void export2(HttpServletResponse response) throws IOException {
        String fileName = System.currentTimeMillis() + ".xlsx";
        OutputStream fileOutputStream = response.getOutputStream();
        response.setHeader("Content-disposition", "attachment;filename=" + fileName);
        // 文件输出位置
//        String outPath = "C:\\Users\\oukele\\Desktop\\test.xlsx";

        try {
            // 所有行的集合
            List<List<Object>> list = new ArrayList<List<Object>>();

            for (int i = 1; i <= 10; i++) {
                // 第 n 行的数据
                List<Object> row = new ArrayList<Object>();
                row.add("第" + i + "单元格");
                row.add("第" + i + "单元格");
                list.add(row);
            }

//            ExcelWriter excelWriter = EasyExcelFactory.getWriter(new FileOutputStream(outPath));

            ExcelWriter excelWriter = EasyExcelFactory.getWriter(fileOutputStream);
            // 表单
            Sheet sheet = new Sheet(1,0);
            sheet.setSheetName("第一个Sheet");
            // 创建一个表格
            Table table = new Table(1);
            // 动态添加 表头 headList --> 所有表头行集合
            List<List<String>> headList = new ArrayList<List<String>>();
            // 第 n 行 的表头
            List<String> headTitle0 = new ArrayList<String>();
            List<String> headTitle1 = new ArrayList<String>();
            List<String> headTitle2 = new ArrayList<String>();
            headTitle0.add("最顶部-1");
            headTitle0.add("标题1");
            headTitle1.add("最顶部-1");
            headTitle1.add("标题2");
            headTitle2.add("最顶部-1");
            headTitle2.add("标题3");

            headList.add(headTitle0);
            headList.add(headTitle1);
            headList.add(headTitle2);
            table.setHead(headList);

            excelWriter.write1(list,sheet,table);
            // 记得 释放资源
            excelWriter.finish();

            System.out.println("ok");
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            //关闭流
            if (fileOutputStream != null) {
                IOUtils.closeQuietly(fileOutputStream);
            }
            //删除本地文件
            if (!StringUtils.isEmpty(fileName)) {
                File file = new File(fileName);
                if (file.exists()) {
                    file.delete();
                }
            }
        }
    }



    @RequestMapping(value = "/export3", method = RequestMethod.GET, produces = "application/json")
    public void export3(HttpServletResponse response) throws IOException {
        String fileName = System.currentTimeMillis() + ".xlsx";
        response.setHeader("Content-disposition", "attachment;filename=" + fileName);
        OutputStream fileOutputStream = response.getOutputStream();
        try {
            //查询数据
            List<ExportVo> list = new ArrayList<>();
            list.add(new ExportVo("123", "lz"));

            // 头的策略
            WriteCellStyle headWriteCellStyle = new WriteCellStyle();
            headWriteCellStyle.setFillForegroundColor(IndexedColors.BLUE.getIndex());

            WriteCellStyle contentWriteCellStyle = new WriteCellStyle();
            contentWriteCellStyle.setFillForegroundColor(IndexedColors.BLUE.getIndex());
            contentWriteCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            contentWriteCellStyle.setHorizontalAlignment(HorizontalAlignment.CENTER);
            contentWriteCellStyle.setBorderBottom(BorderStyle.THIN);
            contentWriteCellStyle.setBorderLeft(BorderStyle.THIN);
            contentWriteCellStyle.setBorderRight(BorderStyle.THIN);
            contentWriteCellStyle.setBorderTop(BorderStyle.THIN);

            HorizontalCellStyleStrategy horizontalCellStyleStrategy = new HorizontalCellStyleStrategy(headWriteCellStyle, contentWriteCellStyle);


            //生成Excel
            EasyExcel.write(response.getOutputStream(), ExportVo.class)
                    .head(getMorningCheckHead("表头"))
                    .registerWriteHandler(horizontalCellStyleStrategy)
                    .sheet("晨午检")
                    .doWrite(list);

        } catch (Exception e) {
            log.error("export failead:", e);
        } finally {
            //关闭流
            if (fileOutputStream != null) {
                IOUtils.closeQuietly(fileOutputStream);
            }
            //删除本地文件
            if (!StringUtils.isEmpty(fileName)) {
                File file = new File(fileName);
                if (file.exists()) {
                    file.delete();
                }
            }
        }
    }


    /**
     * 晨午检的头
     * @param bigTitle
     * @return
     */
    private  List<List<String>> getMorningCheckHead(String bigTitle){
        List<List<String>> head = new ArrayList<List<String>>();
        List<String> head0 = new ArrayList<>();
        head0.add(bigTitle);
        head0.add("班级");
        List<String> head1 = new ArrayList<>();
        head1.add(bigTitle);
        head1.add("姓名");
        head.add(head0);
        head.add(head1);
        return head;
    }

}
