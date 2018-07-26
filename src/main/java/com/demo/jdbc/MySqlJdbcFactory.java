package com.demo.jdbc;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.util.StringUtils;

import java.sql.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

/**
 * @Title: MySqlBatchUtil
 * @ProjectName demo
 * @Description: TODO
 * @author guanjian
 * @date 2018/7/24 16:02
 */
public class MySqlJdbcFactory {
    private static final Logger LOG = LoggerFactory.getLogger(MySqlJdbcFactory.class);

    /**
     * 构建方法
     * @return
     */
    public static MySqlJdbcFactory.Builder build(String driverClassName,String driverUrl,String userName,String passWord){
        return new Builder(driverClassName,driverUrl,userName,passWord);
    }
    public static MySqlJdbcFactory.Builder build(){
        return new MySqlJdbcFactory.Builder(null,null,null,null);
    }
    /**
     * 构建类
     */
    public static class Builder{
        private String driverClassName = "com.mysql.jdbc.Driver";
        private String driverUrl ="jdbc:mysql://localhost:3306/sakila";
        private String userName = "root";
        private String passWord = "root";

        public Builder(String driverClassName,String driverUrl,String userName,String passWord){
            if (!StringUtils.isEmpty(driverClassName))
                this.driverClassName = driverClassName;
            if (!StringUtils.isEmpty(driverUrl))
                this.driverUrl = driverUrl;
            if (!StringUtils.isEmpty(userName))
                this.userName = userName;
            if (!StringUtils.isEmpty(passWord))
                this.passWord = passWord;
        }

        /**
         * 批量导入方法
         */
        public void batchInsert(List dataList) {
            long start = System.currentTimeMillis();
            try(Connection conn = DriverManager.getConnection(driverUrl, userName, passWord)) {
                //配置驱动类
                Class.forName(driverClassName);
                //关闭自动提交事务
                conn.setAutoCommit(false);
                //组合sql
                String sql = "INSERT INTO TEST (NAME,ADDRESS,CREATE_DATE) VALUES ( ?, ?,SYSDATE())";

                PreparedStatement pst = conn.prepareStatement(sql, ResultSet.TYPE_SCROLL_SENSITIVE,
                        ResultSet.CONCUR_READ_ONLY);

                for (int i = 0; i < dataList.size(); i++) {
                    List list = (List) dataList.get(i);
                    pst.setString(1,String.valueOf(list.get(0)));
                    pst.setString(2,String.valueOf(list.get(1)));
                    pst.addBatch();
                }
                //批量执行并提交
                pst.executeBatch();
                conn.commit();

                long end = System.currentTimeMillis();

                LOG.info("插入时间 {} ms",end - start);
            } catch (Exception ex) {
                ex.printStackTrace();
            }
        }

        /**
         * 查询方法
         */
        public List<HashMap> select() {

            long start = System.currentTimeMillis();

            List dataList = new ArrayList();

            try(Connection conn = DriverManager.getConnection(driverUrl, userName, passWord)) {
                //配置驱动类
                Class.forName(driverClassName);

                //查询SQL
                String sql = "SELECT * FROM TEST";

                PreparedStatement pst = conn.prepareStatement(sql, ResultSet.TYPE_SCROLL_SENSITIVE,
                        ResultSet.CONCUR_READ_ONLY);

                ResultSet rs = pst.executeQuery();

                ResultSetMetaData md = rs.getMetaData();

                int columnCount = md.getColumnCount(); //Map rowData;

                while (rs.next()) { //rowData = new HashMap(columnCount);

                    List cellList = new ArrayList(columnCount);//一次创建，避免resize

                    for (int i = 1; i <= columnCount; i++) {
                        cellList.add(String.valueOf(rs.getObject(i)));
                    }

                    dataList.add(cellList);
                    cellList = null;//help gc

                }
                long end = System.currentTimeMillis();

                LOG.info("查询时间 {} ms",end - start);
            } catch (Exception ex) {
                ex.printStackTrace();
            }
            return dataList;
        }

        /**
         *  getter setter
         */
        public String getDriverClassName() {
            return driverClassName;
        }

        public String getDriverUrl() {
            return driverUrl;
        }

        public String getUserName() {
            return userName;
        }

        public String getPassWord() {
            return passWord;
        }

        public Builder driverClassName(String driverClassName) {
            this.driverClassName = driverClassName;
            return this;
        }
        public Builder driverUrl(String driverUrl) {
            this.driverUrl = driverUrl;
            return this;
        }
        public Builder userName(String userName) {
            this.userName = userName;
            return this;
        }
        public Builder passWord(String passWord) {
            this.passWord = passWord;
            return this;
        }

        @Override
        public String toString() {
            return "Builder{" +
                    "driverClassName='" + driverClassName + '\'' +
                    ", driverUrl='" + driverUrl + '\'' +
                    ", userName='" + userName + '\'' +
                    ", passWord='" + passWord + '\'' +
                    '}';
        }
    }
}
