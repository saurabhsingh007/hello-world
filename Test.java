package com.jilit.irp.util;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DatabaseMetaData;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.util.HashMap;
import java.util.Map;
import oracle.jdbc.OracleConnection;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;

public class Test {

    private static final String FILE_NAME = "D:/MyFirstExcel.xls";

    public static void main(String args[]) throws Exception {
        String databaseName = "camplx2";
        String schema = "NITD24112017";
        Class.forName("oracle.jdbc.driver.OracleDriver");
        Connection conn = DriverManager.getConnection("jdbc:oracle:thin:@//172.16.7.156:1521/cmp11", "NITD24112017", "NITD24112017");
        ((OracleConnection)conn).setRemarksReporting(true);
        Map<String, String> tableMap = new HashMap();
        String[] types = {"TABLE"};
        DatabaseMetaData meta = conn.getMetaData();
        ResultSet resultSet = meta.getTables(databaseName, schema, "%", types);
        while (resultSet.next()) {
            if (resultSet.getString(4).equalsIgnoreCase("TABLE")) {
                tableMap.put(resultSet.getString("TABLE_NAME"), "");
            }
        }
        resultSet.close();
        // --- LISTING DATABASE COLUMN NAMES ---
        ResultSet resultSet1 = null;
        ResultSet resultSet2 = null;
        HSSFWorkbook hwb = new HSSFWorkbook();
        HSSFSheet sheet = hwb.createSheet("Table With Column Details");
        HSSFRow tabrow = null;
        HSSFRow colrow = null;
        HSSFCell tabeCell = null;
        HSSFCell colCell = null;
        int i = 0;
        String columntype="";
        boolean pk=false;
        String pkindex="";
        HSSFCellStyle headerstyle = hwb.createCellStyle();
        headerstyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        headerstyle.setFillForegroundColor(HSSFColor.BLUE_GREY.index);
        HSSFFont font = hwb.createFont();
        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        headerstyle.setFont(font);
        headerstyle.setAlignment(HSSFCellStyle.ALIGN_LEFT);

        HSSFCellStyle tablestyle = hwb.createCellStyle();
        tablestyle.setFont(font);
        tablestyle.setAlignment(HSSFCellStyle.ALIGN_LEFT);
        for (Map.Entry entry : tableMap.entrySet()) {
            tabrow = sheet.createRow((short) i++);            
            tabeCell = tabrow.createCell((short) 0);
            tabeCell.setCellValue("Sr.No.");
            tabeCell.setCellStyle(headerstyle);
            tabeCell = tabrow.createCell((short) 1);
            tabeCell.setCellValue("Column Name");
            tabeCell.setCellStyle(headerstyle);
            tabeCell = tabrow.createCell((short) 2);
            tabeCell.setCellValue("ID");
            tabeCell.setCellStyle(headerstyle);
            tabeCell = tabrow.createCell((short) 3);
            tabeCell.setCellValue("PK");
            tabeCell.setCellStyle(headerstyle);
            tabeCell = tabrow.createCell((short) 4);
            tabeCell.setCellValue("NULL?");
            tabeCell.setCellStyle(headerstyle);
            tabeCell = tabrow.createCell((short) 5);
            tabeCell.setCellValue("Data Type");
            tabeCell.setCellStyle(headerstyle);
            tabeCell = tabrow.createCell((short) 6);
            tabeCell.setCellValue("Comments");
            tabeCell.setCellStyle(headerstyle);
            tabeCell = tabrow.createCell((short) 7);

            tabrow = sheet.createRow((short) i++);
            tabeCell = tabrow.createCell((short) 0);
            tabeCell.setCellValue(entry.getKey().toString());
            tabeCell.setCellStyle(tablestyle);
            
            resultSet1 = meta.getColumns(databaseName, schema, entry.getKey().toString(), "%");
            resultSet2 = meta.getPrimaryKeys(databaseName, schema, entry.getKey().toString());

            while (resultSet1.next()) {
                colrow = sheet.createRow((short) i++);
                colCell = colrow.createCell((short) 1);
                colCell.setCellValue(resultSet1.getString("COLUMN_NAME"));
                columntype = resultSet1.getString("TYPE_NAME").toString();
                if("NUMBER".equalsIgnoreCase(columntype)){
                    columntype=columntype+"("+resultSet1.getString("COLUMN_SIZE")+")";
                }else if (("VARCHAR2".equalsIgnoreCase(columntype)) || ("CHAR".equalsIgnoreCase(columntype))){
                   columntype=columntype+"("+resultSet1.getString("COLUMN_SIZE")+" Byte)";
                }

                colCell = colrow.createCell((short) 2);
                colCell.setCellValue(resultSet1.getString("ORDINAL_POSITION"));

                colCell = colrow.createCell((short) 3);
                pk=false;
                while(resultSet2.next()){
                    if(resultSet2.getString("COLUMN_NAME").toString().equalsIgnoreCase(resultSet1.getString("COLUMN_NAME").toString())){
                        pk=true;
                        pkindex=resultSet2.getString(5);
                        break;
                    }
                }
                colCell.setCellValue((pk==true ? pkindex:""));

                colCell = colrow.createCell((short) 4);
                colCell.setCellValue(("0".equalsIgnoreCase(resultSet1.getString("NULLABLE").toString()) ? "N" :"Y"));
                colCell = colrow.createCell((short) 5);
                colCell.setCellValue(columntype);
                colCell = colrow.createCell((short) 6);
                colCell.setCellValue(resultSet1.getString("REMARKS"));
            }
            resultSet1.close();
            resultSet2.close();
        }
        try {
            FileOutputStream outputStream = new FileOutputStream(FILE_NAME);
            hwb.write(outputStream);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } 

        System.out.println("Done");
    }
}