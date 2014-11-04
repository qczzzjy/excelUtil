package cn.xiaozhi.main;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.util.List;
import java.util.Map;

/**
 * Created by xiaozhi on 14-11-1.
 */
public class ExcelUtilByPoi {
    //Singleton
    private ExcelUtilByPoi(){
    }

    private static class HolderClass{
        private final static ExcelUtilByPoi instance = new ExcelUtilByPoi();
    }

    public static ExcelUtilByPoi getInstance(){
        return HolderClass.instance;
    }

    /**
     *
     * @param outFilePath
     * @param header
     * @param content
     */
    public void exportExcel2007ByPoi(String outFilePath,String header[],Map<String,List<Object[]>> content){
        try {
            Workbook wb = createWorkBook(XSSFWorkbook.class.getName());
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
            Workbook wb = createWorkBook(HSSFWorkbook.class.getName());
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
        Integer row=0;
        //创建header行
        createHeader(sheet,row,header);
        for(Object[] objects:sheetContent){
            Row rowE = sheet.createRow(row);
            for(int i=0;i<objects.length;i++){
                rowE.createCell(i).setCellValue(objects[i].toString());
            }
        }
    }

    private void createHeader(Sheet sheet, Integer row, String[] header) {
        Row rowE = sheet.createRow(row);
        for(int i=0;i<header.length;i++){
            rowE.createCell(i).setCellValue(header[i]);
        }
        row++;
    }

    /**
     *
     * @param wb
     * @param cellType hader,content
     * @return
     */
    private CellStyle createCellStyle(Workbook wb,String cellType){
        CellStyle cellStyle = wb.createCellStyle();
//        cellStyle.setFillBackgroundColor();
//        cellStyle.setFont();
        return cellStyle;
    }

}
