package cn.xiaozhi.main;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.List;
import java.util.Map;
import java.util.Properties;

/**
 * Created by xiaozhi on 14-11-1.
 */
public class ExcelUtilByPoi {
    //Singleton
    private ExcelUtilByPoi(){
        readPropertes();
    }

    private static class HolderClass{
        private final static ExcelUtilByPoi instance = new ExcelUtilByPoi();
    }

    public static ExcelUtilByPoi getInstance(){
        return HolderClass.instance;
    }



    private Workbook wb;
    private Properties properties;

    /**
     *
     * @param outFilePath
     * @param header
     * @param content
     */
    public void exportExcel2007ByPoi(String outFilePath,String header[],Map<String,List<Object[]>> content){
        try {
            wb = createWorkBook(XSSFWorkbook.class.getName());
            exportExcel(wb, header, content);
            wb.write(new FileOutputStream(outFilePath));

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     *
     * @param outFilePath
     * @param header
     * @param content
     */
    public void exportExcel2003ByPoi(String outFilePath,String header[],Map<String,List<Object[]>> content){
        try {
            wb = createWorkBook(HSSFWorkbook.class.getName());
            exportExcel(wb, header, content);
            wb.write(new FileOutputStream(outFilePath));

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void exportExcel(Workbook wb, String[] header, Map<String, List<Object[]>> content) {
        for(Map.Entry<String,List<Object[]>> entry:content.entrySet()){
            String sheetName = entry.getKey();
            List<Object[]> sheetContent = entry.getValue();

            Sheet sheet =  wb.createSheet(sheetName);
            createSheet(sheet,header,sheetContent);
        }
    }



    private Workbook createWorkBook(String excelType) throws Exception {
        return (Workbook)Class.forName(excelType).newInstance();
    }

    private void createSheet(Sheet sheet, String[] header, List<Object[]> sheetContent) {
        int row=0;
        //创建header行
        row = createHeader(sheet,row,header);
        for(Object[] objects:sheetContent){
            Row rowE = sheet.createRow(row);
            for(int i=0;i<objects.length;i++){
                rowE.createCell(i).setCellValue(objects[i].toString());
            }
            row++;
        }
    }

    private int createHeader(Sheet sheet, int row, String[] header) {
        Row rowE = sheet.createRow(row);
        CellStyle cs = null;
        if(properties!=null){
            cs = createCellStyle("header");
        }
        for(int i=0;i<header.length;i++){
            Cell cell = rowE.createCell(i);
            if(cs!=null){
                cell.setCellStyle(cs);
            }
            rowE.createCell(i).setCellValue(header[i]);
        }
        return ++row;
    }

    /**
     *
     * @param cellType hader,content
     * @return
     */
    private CellStyle createCellStyle(String cellType){
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setFillBackgroundColor(
        Short.parseShort(properties.getProperty(cellType+"background.color")));
//        cellStyle.setFont();
        return cellStyle;
    }

    private void readPropertes(){
        InputStream is = ClassLoader.getSystemResourceAsStream("excelUtil.properties");
        properties=new Properties();
        try {
            properties.load(is);
        } catch (IOException e) {
            ;
        }
    }

}
