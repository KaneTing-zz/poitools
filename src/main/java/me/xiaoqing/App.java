package me.xiaoqing;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args )
    {
        try {
            //1.生成xls
            //geneXlsByHSSF();

            //2.生成xlsx
            geneXlsByXSSF();
        }catch (FileNotFoundException e){
            System.out.print(e.getMessage());
        }catch (IOException e){
            System.out.print(e.getMessage());
        }

    }

    private static void geneXlsByHSSF() throws FileNotFoundException, IOException{
        HSSFWorkbook workbook = new HSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        workbook.setSheetName(0, "mysheet");
        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0, CellType.STRING);
        cell.setCellValue("I love coding");

        String fileName = String.valueOf(new Date().getTime())+ ".xls";
        FileOutputStream outputStream = new FileOutputStream("/Users/ManKane/Documents/workspaces/exercises/"+fileName);
        workbook.write(outputStream);
        outputStream.close();
        workbook.close();
    }

    private static void geneXlsByXSSF() throws FileNotFoundException, IOException{
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet();
        workbook.setSheetName(0, "mysheet");
        XSSFRow row = sheet.createRow(0);
        XSSFCell cell = row.createCell(0, CellType.NUMERIC);
        cell.setCellValue(100);

        String fileName = String.valueOf(new Date().getTime())+ ".xlsx";
        FileOutputStream outputStream = new FileOutputStream("/Users/ManKane/Documents/workspaces/exercises/"+fileName);
        workbook.write(outputStream);
        outputStream.close();
        workbook.close();
    }
}
