package com.demo.api;

import com.demo.jdbc.MySqlJdbcFactory;
import com.monitorjbl.xlsx.StreamingReader;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

/**
 * @Title: ExcelStreamingReader
 * @ProjectName demo
 * @Description: TODO
 * @author guanjian
 * @date 2018/7/26 9:03
 */

public class ExcelStreamingReader {
    private final static Logger LOG = LoggerFactory.getLogger(ExcelStreamingReader.class);
    /**
     * excel读取
     * 通过excel-streaming-reader实现，仅支持xlsx读取
     * @param excelUrl 文本路径
     * @return List 数据集合 （数据结构 List<List>）
     */
    public static List readByStreamingReader(String excelUrl) {
        //所有数据容器
        List dataList = null;
        //单元格数据容器
        List cellList = null;

        try (FileInputStream in = new FileInputStream(excelUrl)) {

            //通过构造器配置参数获得WorkBook，仅支持xlsx
            Workbook wk = StreamingReader.builder()
                    .rowCacheSize(100)  //缓存到内存中的行数，默认是10
                    .bufferSize(4096)  //读取资源时，缓存到内存的字节大小，默认是1024
                    .open(in);  //打开资源，必须，可以是InputStream或者是File，注意：只能打开XLSX格式的文件
            //获取sheet
            Sheet sheet = wk.getSheetAt(0);
            //设置size，减少resize
            dataList = Collections.synchronizedList(new ArrayList(sheet.getLastRowNum()));

            //遍历所有的行
            for (Row row : sheet) {
                //设置size，减少resize
                cellList = Collections.synchronizedList(new ArrayList(row.getLastCellNum()));
                //遍历所有的列
                for (Cell cell : row) {
                    //存单元格数据
                    cellList.add(cell.getStringCellValue());
                }
                dataList.add(cellList);//存行数据
                cellList = null;//help gc
            }

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        //返回数据
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
        readByStreamingReader(exportUrl);

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
