package com.demo.api;

import com.demo.jdbc.MySqlJdbcFactory;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;


/**
 * @Title: POI
 * @ProjectName demo
 * @Description: TODO
 * @author guanjian
 * @date 2018/7/24 8:36
 */

public class POI {

    /** POI读取
     *  @description 根据excelUrl路径读取excel文本，支持xls、xlsx读取，大体积的会出现OOM
     *  @param excelUrl 文本路径
     *  @return List 数据集合 （数据结构List<List>）
     */
    public static List<List> read(String excelUrl){
        //所有数据容器
        List dataList = null;
        //单元格数据容器
        List cellList = null;

        try {
            File file = new File(excelUrl);
            //WorkBookFactory会根据格式创建XSSFWorkbook 或者 HSSFWorkbook的实现
            Workbook wk = WorkbookFactory.create(file);
            //获取sheet
            Sheet sheet = wk.getSheetAt(0);
            //根据文本行数设置size，避免resize减少性能消耗
            dataList = Collections.synchronizedList(new ArrayList(sheet.getLastRowNum()));

            //遍历所有的行
            for (Row row : sheet) {
                //根据当前行的列数设置size，避免resize减少性能消耗
                cellList = Collections.synchronizedList( new ArrayList(row.getLastCellNum()));
                //遍历所有的列
                for (Cell cell : row) {
                    cellList.add(getValue(cell));//存单元格数据  注意数据类型转换
                }
                dataList.add(cellList); //存当前行数据
                cellList = null;//help gc
            }
            //关闭
            wk.close();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }
        return dataList;
    }

    /**
     * 生成excel
     * @description 根据dataList生成excel数据，支持xls、xlsx的生成，大体积的会出现OOM
     * @param exportUrl 文本路径
     * @param dataList 数据容器 （数据结构 List<List>）
     */
    public static void write(String exportUrl,List dataList){

        try(FileOutputStream out = new FileOutputStream(exportUrl)) {
            Workbook wb = null;
            //根据文件后缀名判断生成Workbook的具体实现，仅仅修改后缀名而文本格式内容不一致也会抛异常
            if(exportUrl.endsWith(".xlsx")){
                wb = new HSSFWorkbook();//EXCEL 97-2003
            }else{
                wb = new XSSFWorkbook();//EXCEL 2007
            }
            //创建sheet
            Sheet sheet = wb.createSheet();

            //行数
            int rowNum = dataList.size();

            //循环创建行
            for (int i = 0;i<rowNum;i++) {
                //创建行
                Row row = sheet.createRow(i);
                //当前列数
                int columnNum = ((List)dataList.get(i)).size();
                //循环创建列
                for (int j =0;j<columnNum;j++) {
                    //创建列
                    Cell cell = row.createCell(j);
                    //存单元格数据
                    cell.setCellValue(String.valueOf(((List)dataList.get(i)).get(j)));
                }
            }
            dataList = null; // help gc

            //将数据写入流
            wb.write(out);
            //关闭
            wb.close();

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 生成excel
     * 通过SXSSF导出
     * @param excelUrl
     * @param dataList
     */
    public static void writeBySXSSF(String excelUrl,List dataList){
        try(FileOutputStream out = new FileOutputStream(excelUrl)) {
            //SXSSFWorkbook仅支持xlsx的读取
            SXSSFWorkbook wb = new SXSSFWorkbook(100); // 在内存当中保持 100 行 , 超过的数据放到硬盘中
            //创建sheet
            Sheet sh = wb.createSheet();

            //行数
            int rowNum = dataList.size();

            //循环创建行
            for(int i = 0; i < rowNum; i++){
                //创建行
                Row row = sh.createRow(i);
                //当前行的列数
                int cellNum = ((List)dataList.get(i)).size();
                //循环创建单元格
                for(int j = 0; j < cellNum; j++){
                    //创建单元格
                    Cell cell = row.createCell(j);
                    //存单元格数据，注意格式转换
                    cell.setCellValue(String.valueOf(((List)dataList.get(i)).get(j)));
                }

            }
            //数据写入磁盘
            wb.write(out);
            //销毁临时文件
            wb.dispose();
            //关闭
            wb.close();

        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    /**
     * 单元格类型转换
     * @param cell
     * @return
     */
    private static Object getValue(Cell cell) {
        Object obj = null;
        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_BOOLEAN:
                obj = cell.getBooleanCellValue();
                break;
            case Cell.CELL_TYPE_ERROR:
                obj = cell.getErrorCellValue();
                break;
            case Cell.CELL_TYPE_NUMERIC:
                obj = cell.getNumericCellValue();
                break;
            case Cell.CELL_TYPE_STRING:
                obj = cell.getStringCellValue();
                break;
            default:
                break;
        }
        return obj;
    }

    /**
     * 测试
     * @throws Exception
     */
    public static void main(String[] args) throws Exception {

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
        writeBySXSSF(exportUrl,dataList);

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
