package org.zz.persionalTool.ExcelTool;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;

public class OracleTableExportToExcel {

    private static String userName = "SAMPLE";
    private static String password = "SAMPLE";
    private static String url = "jdbc:oracle:thin:@127.0.0.1:1521:test";

    private static final String sql =
            "SELECT " +
//                    "t1.Table_Name || chr(13) || t3.comments AS \"表名称及说明\", " +
//            "       t3.comments                                 AS \"表说明\", " +
                    "       t1.Column_Name AS \"字段名称\", " +
                    "       t1.DATA_TYPE || '(' || t1.DATA_LENGTH || ')' AS \"数据类型\", " +
                    "       t1.NullAble AS \"是否为空\", " +
                    "       t2.Comments AS \"字段说明\", " +
                    "       t1.Data_Default \"默认值\" " +
                    "  FROM cols t1 " +
                    "  LEFT JOIN user_col_comments t2 " +
                    "    ON t1.Table_name = t2.Table_name " +
                    "   AND t1.Column_Name = t2.Column_Name " +
                    "  LEFT JOIN user_tab_comments t3 " +
                    "    ON t1.Table_name = t3.Table_name " +
                    "  LEFT JOIN user_objects t4 " +
                    "    ON t1.table_name = t4.OBJECT_NAME " +
                    " WHERE NOT EXISTS (SELECT t4.Object_Name " +
                    "          FROM User_objects t4 " +
                    "         WHERE t4.Object_Type = 'TABLE' " +
                    "           AND t4.Temporary = 'Y' " +
                    "           AND t4.Object_Name = t1.Table_Name) " +
                    " and t1.TABLE_NAME = ? " +
                    " ORDER BY  t1.Column_ID ";
//            " ORDER BY t1.Table_Name, t1.Column_ID ";

    public static void main(String[] args) throws Exception {

        XSSFWorkbook xsb = new XSSFWorkbook();

        Connection conn = null;
        PreparedStatement pstat = null;
        ResultSet rs = null;
        FileOutputStream out = null;
        try {
            Class.forName("oracle.jdbc.driver.OracleDriver");
            conn = DriverManager.getConnection(url, userName, password);
            pstat = conn.prepareStatement(sql);

            XSSFSheet sheet;
            XSSFRow row;
            XSSFCellStyle normalStyle = getNormalStyle(xsb);
            XSSFCell cell ;
            for (String table : tableNames) {
                info("处理表："+table);
                sheet = xsb.createSheet(table.toUpperCase());
                createHeaderRow(sheet,xsb);

                pstat.clearParameters();
                pstat.setString(1, table.toUpperCase());
                rs = pstat.executeQuery();
                int i = 1;
                while (rs.next()) {
                    row = sheet.createRow(i);
                    cell = row.createCell(0);
                    cell.setCellStyle(normalStyle);
                    cell.setCellValue(rs.getString(1));

                    cell = row.createCell(1);
                    cell.setCellStyle(normalStyle);
                    cell.setCellValue(rs.getString(2));

                    cell = row.createCell(2);
                    cell.setCellStyle(normalStyle);
                    cell.setCellValue(rs.getString(3));

                    cell = row.createCell(3);
                    cell.setCellStyle(normalStyle);
                    cell.setCellValue(rs.getString(4));
                    i++;

                    //自动适应列宽
                    for (int j = 0; j < 4 ; j++) {
                        sheet.autoSizeColumn(j, true);
                    }
                }

                rs.close();
            }

            File file = new File("D:/tableStructure.xlsx");
            //输出文件
            out = new FileOutputStream(file);
            info("输出文件");
            xsb.write(out);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (null != rs)
                rs.close();
            if (null != pstat)
                pstat.close();
            if (null != conn)
                conn.close();
            if (null != out)
                out.close();
        }

    }

    private static void info(String msg){
        System.out.println(msg);
    }

    private static XSSFSheet createHeaderRow(XSSFSheet sheet,XSSFWorkbook xsb) {

        XSSFRow headerRow = sheet.createRow(0);
        XSSFCell cell = headerRow.createCell(0);
        XSSFCellStyle headerStyle = getHeaderStyle(xsb);
        cell.setCellStyle(headerStyle);
        cell.setCellValue("字段名称");
        cell = headerRow.createCell(1);
        cell.setCellStyle(headerStyle);
        cell.setCellValue("数据类型");
        cell = headerRow.createCell(2);
        cell.setCellStyle(headerStyle);
        cell.setCellValue("是否为空");
        cell = headerRow.createCell(3);
        cell.setCellStyle(headerStyle);
        cell.setCellValue("字段说明");
//        cell = headerRow.createCell(4);
//        cell.setCellStyle(headerStyle);
//        cell.setCellValue("默认值");

        return sheet;
    }

    private static XSSFCellStyle getHeaderStyle(XSSFWorkbook xssfWorkbook){
        XSSFCellStyle cellStyle = xssfWorkbook.createCellStyle();
        Font font = xssfWorkbook.createFont();
        font.setBold(true);
        cellStyle.setFont(font);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        return cellStyle;
    }


    private static XSSFCellStyle getNormalStyle(XSSFWorkbook xssfWorkbook){
        XSSFCellStyle cellStyle = xssfWorkbook.createCellStyle();
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        return cellStyle;
    }


    public static String[] tableNames = {
            "t_task_dic",
    };
}
