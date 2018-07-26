package com.demo.api;

import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.read.context.AnalysisContext;
import com.alibaba.excel.read.event.AnalysisEventListener;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.demo.jdbc.MySqlJdbcFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.util.List;

/**
 * @Title: EasyExcel
 * @ProjectName demo
 * @Description: TODO
 * @author guanjian
 * @date 2018/7/25 14:35
 */

public class EasyExcel {
    private final static Logger LOG = LoggerFactory.getLogger(EasyExcel.class);
    /**
     * 生成excel
     * 支持xls 、xlsx
     * 参考github上的示例
     * @param url
     * @param dataList
     */
    public static void write(String url,List dataList){

        try(  OutputStream out = new FileOutputStream(new File(url))) {
            ExcelWriter writer = null;
            //根据文件后缀名判断生成Workbook的具体实现，仅仅修改后缀名而文本格式内容不一致也会抛异常
            if(url.endsWith(".xlsx")){
                writer = new ExcelWriter(out, ExcelTypeEnum.XLSX);
            }else{
                writer = new ExcelWriter(out, ExcelTypeEnum.XLS);
            }
            //写第一个sheet, sheet1  数据全是List<String> 无模型映射关系
            Sheet sheet1 = new Sheet(1, 0);
            //数据写入
            writer.write(dataList, sheet1);
            //关闭
            writer.finish();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 读取excel
     * 支持xls 、xlsx
     * @param url
     * @return
     */
    public static void read(String url){
        //所有数据容器
        List dataList = null;
        //单元格数据容器
        List cellList = null;

        try(InputStream inputStream = new FileInputStream(url)) {
            ExcelReader reader = null;
            if(url.endsWith(".xlsx")){
               reader = new ExcelReader(inputStream, ExcelTypeEnum.XLSX, null,
                        new AnalysisEventListener<List<String>>() {
                            @Override
                            public void invoke(List<String> object, AnalysisContext context) {
                              System.out.println(
                                    "当前sheet:" + context.getCurrentSheet().getSheetNo() + " 当前行：" + context.getCurrentRowNum()
                                            + " data:" + object);
                            }
                            @Override
                            public void doAfterAllAnalysed(AnalysisContext context) {

                            }
                        });
            }else{
                reader = new ExcelReader(inputStream, ExcelTypeEnum.XLS, null,
                        new AnalysisEventListener<List<String>>() {
                            @Override
                            public void invoke(List<String> object, AnalysisContext context) {
                            System.out.println(
                                    "当前sheet:" + context.getCurrentSheet().getSheetNo() + " 当前行：" + context.getCurrentRowNum()
                                            + " data:" + object);
                            }
                            @Override
                            public void doAfterAllAnalysed(AnalysisContext context) {

                            }
                        });
            }

            reader.read();

            reader.finish();
        } catch (Exception e) {
            e.printStackTrace();

        }
    }

    public static void main(String[] args) {
        /**
         * 文件路径和数据容器
         */
        String xlsUrl = "\\path\\测试.xls";
        String xlsxUrl = "\\path\\测试.xlsx";
        String exportUrl = "\\path\\导出.xls";
        List dataList = null;

        /**
         * 数据库相关
         */
        //获取jdbc
        MySqlJdbcFactory.Builder builder = MySqlJdbcFactory.build();
        //批量插入
        builder.batchInsert(dataList);
        //查询
        dataList = builder.select();

        /**
         * 方法调用
         */
        read(xlsUrl);
        write(exportUrl,dataList);

        /**
         * 时间消耗
         */
        //开始时间
        long start = System.currentTimeMillis();
        //结束时间
        long end = System.currentTimeMillis();
        //时间差
        LOG.info("时间消耗：{} ms",end - start);
    }
}
