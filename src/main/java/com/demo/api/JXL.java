package com.demo.api;

import com.demo.jdbc.MySqlJdbcFactory;
import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

/**
 * @Title: JXL
 * @ProjectName demo
 * @Description: TODO
 * @author guanjian
 * @date 2018/7/25 15:01
 */

public class JXL {

    /**
     * JXL生成excel
     * 仅支持xls
     * @param url
     * @param dataList
     */
    public static void write(String url,List dataList){

        try {
            //接收输入流
            File file=new File(url);
            file.createNewFile();
            //创建工作簿
            WritableWorkbook workbook=Workbook.createWorkbook(file);
            //创建sheet
            WritableSheet sheet=workbook.createSheet("sheetName",0);
            //设置单元格
            Label label=null;

            //行数
            int rowNum = dataList.size();
            //遍历行
            for(int i=0;i<rowNum;i++){
                //当前行的列数
                int columuNum = ((List)dataList.get(i)).size();
                //遍历列
                for (int j=0;j<columuNum;j++){
                    //添加单元格
                    label=new Label(j,i,String.valueOf(((List)dataList.get(i)).get(j)));//new Label(列, 行, 值)
                    sheet.addCell(label);//存单元格
                }
                label = null;//help gc
            }

            //写入数据，一定记得写入数据，不然前面全白费了
            workbook.write();
            //最后一步，关闭工作簿
            workbook.close();

        } catch (IOException e) {
            e.printStackTrace();
        } catch (RowsExceededException e) {
            e.printStackTrace();
        } catch (WriteException e) {
            e.printStackTrace();
        }


    }

    /**
     * 读取Excel
     * @param url
     * @return
     */
    public static List read(String url){
        //所有数据容器
        List dataList = null;
        //单元格数据容器
        List cellList = null;

        try {
                //创建workbook
                Workbook workbook = Workbook.getWorkbook(new File(url));
                //获取第一个工作表sheet
                Sheet sheet= workbook.getSheet(0);
                //一次生成，避免resize
                dataList = Collections.synchronizedList(new ArrayList(sheet.getRows()));
                //遍历所有行
                for(int i=0;i<sheet.getRows();i++){
                    //一次生成，避免resize
                    cellList = Collections.synchronizedList(new ArrayList(sheet.getColumns()));
                    //遍历所有列
                    for(int j=0;j<sheet.getColumns();j++){
                        //获取当前单元格
                        Cell cell=sheet.getCell(j,i);
                        //存单元格数据
                        cellList.add(cell.getContents());
                    }
                    dataList.add(cellList);//存当前行数据
                    cellList = null;//help gc
                }
                //关闭
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            } catch (BiffException e) {
                e.printStackTrace();
            }
            return dataList;
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
        System.out.println(new Double(end - start)/1000 +"s");
    }

}
