package me.xiaoqing;

import com.monitorjbl.xlsx.StreamingReader;
import junit.framework.Assert;
import me.xiaoqing.bigexcel.*;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Date;
import java.util.Iterator;

/**
 * Hello world!
 *
 */
public class App 
{

    //private static String fileDirectory = "/Users/ManKane/Documents/workspaces/exercises/";
    private static String fileDirectory = "D://execises//";

    public static void main( String[] args )
    {
        try {
            //1.生成xls, 使用HSSF
            //geneXlsByHSSF();

            //2.生成xlsx, 使用XSSF
            //geneXlsByXSSF();

            //3.生成大文件xlsx, 使用XSSF
            //geneLargeFileByXSSF();

            //4.生成大文件,使用SXSSF
            //writeXlsBySXSSF();

            //5.读取xls或者xlsx
            //readXlsByFactory();

            //6.读取大文件
            readLargeXls();

            //7.通过SAX读取大excel文件

            //readFileByEventModel();

            //8.写大excel文件
            //writeBigFileByXSSF();
        }
//        catch (InvalidFormatException e){
//            e.printStackTrace();
//        }
        catch (IOException e){
            e.printStackTrace();
        }
        catch (Exception e){
            e.printStackTrace();
        }

    }

    /**
     * 通过HSSF生成xls文件
     * @throws IOException
     */
    private static void geneXlsByHSSF() throws IOException{
        HSSFWorkbook workbook = new HSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        workbook.setSheetName(0, "mysheet");
        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0, CellType.STRING);
        cell.setCellValue("I love coding");

        String fileName = String.valueOf(new Date().getTime())+ ".xls";
        FileOutputStream outputStream = new FileOutputStream(fileDirectory+fileName);
        workbook.write(outputStream);
        outputStream.close();
        workbook.close();
    }

    /**
     * 通过XSSF生成xlsx文件
     * @throws IOException
     */
    private static void geneXlsByXSSF() throws IOException{
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet();
        workbook.setSheetName(0, "mysheet");
        XSSFRow row = sheet.createRow(0);
        XSSFCell cell = row.createCell(0, CellType.NUMERIC);
        cell.setCellValue(100);

        String fileName = String.valueOf(new Date().getTime())+ ".xlsx";
        FileOutputStream outputStream = new FileOutputStream(fileDirectory+fileName);
        workbook.write(outputStream);
        outputStream.close();
        workbook.close();
    }

    /**
     * 通过XSSF生成大xlsx文件
     * Exception in thread "main" java.lang.OutOfMemoryError: Java heap space
     * 经测试37000行数据是可以生成的
     * 改用 writeXlsBySXSSF
     * @throws IOException
     */
    private static void geneLargeFileByXSSF() throws IOException{
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet();
        workbook.setSheetName(0, "mysheet");

        for(int rownum = 0; rownum < 370000; rownum ++){
            Row row = sheet.createRow(rownum);
            for(int cellnum = 0; cellnum < 10; cellnum ++){
                Cell cell = row.createCell(cellnum);
                cell.setCellValue("i love coding! Rownum is " + rownum);
            }
        }

        String fileName = String.valueOf(new Date().getTime())+ ".xlsx";
        FileOutputStream outputStream = new FileOutputStream(fileDirectory+fileName);
        workbook.write(outputStream);
        outputStream.close();
        workbook.close();
    }

    /**
     * 通过SXSSF生成大文件
     */
    private static void writeXlsBySXSSF() throws IOException{

        long start = System.currentTimeMillis();

        SXSSFWorkbook wb = new SXSSFWorkbook(1000);// keep 1000 rows in memory, exceeding rows will be flushed to disk
        Sheet sh = wb.createSheet();
        for(int rownum = 0; rownum < 370000; rownum ++){
            Row row = sh.createRow(rownum);
            for(int cellnum = 0; cellnum < 10; cellnum ++){
                Cell cell = row.createCell(cellnum);
                cell.setCellValue("i love coding! Rownum is " + rownum);
            }
        }

        // Rows with rownum < 9900 are flushed and not accessible
        for(int rownum = 0; rownum < 369000; rownum++){
            Assert.assertNull(sh.getRow(rownum));
        }

        // ther last 1000 rows are still in memory
        for(int rownum = 369000; rownum < 370000; rownum++){
            Assert.assertNotNull(sh.getRow(rownum));
        }

        FileOutputStream out = new FileOutputStream(fileDirectory+ "sxssf.xlsx");
        wb.write(out);
        out.close();
        // dispose of temporary files backing this workbook on disk
        wb.dispose();

        long end = System.currentTimeMillis();
        System.out.println("....................."+(end-start)/1000);
    }

    /**
     * 读取excel文件，usermodel只能用于小文件，因为会将整个文件一次性加载到内存中
     * @throws IOException
     * @throws InvalidFormatException
     */
    private static void readXlsByFactory() throws IOException, InvalidFormatException {
        //String fileName = String.valueOf(new Date().getTime())+ ".xlsx";
        String fileName = "1513915637097.xlsx";
        Workbook workbook = WorkbookFactory.create(new File(fileDirectory+ fileName));
        Sheet sheet = workbook.getSheetAt(0);
        Iterator iterator = sheet.rowIterator();
        while (iterator.hasNext()){
            Row row = (Row)iterator.next();
            Cell cell = row.getCell(0);
            System.out.println(cell.getStringCellValue());
        }
    }

    /**
     * 读取excel文件，使用内存缓存来分批次读取;只能打开XLSX格式的文件
     * @throws IOException
     * @throws InvalidFormatException
     */
    private static void readLargeXls() throws IOException, InvalidFormatException {
        Date startDate = new Date();
        //String fileName = String.valueOf(new Date().getTime())+ ".xlsx";
        String fileName = "37万.xlsx";
        FileInputStream in = new FileInputStream(fileDirectory+ fileName);
        Workbook wk = StreamingReader.builder().rowCacheSize(100)  //缓存到内存中的行数，默认是10
                .bufferSize(2048)  //读取资源时，缓存到内存的字节大小，默认是1024
                .open(in);  //打开资源，必须，可以是InputStream或者是File，注意：只能打开XLSX格式的文件
        Sheet sheet = wk.getSheetAt(0);
        for(Row row : sheet){
            if(row.getRowNum() % 1000 == 0){
                Cell cell1 = row.getCell(0);
                System.out.print(cell1.getStringCellValue()+ "  ");
                Cell cell2 = row.getCell(1);
                System.out.println(cell2.getStringCellValue());
            }
        }
        Date endDate = new Date();
        System.out.println("------------used time is: " + (endDate.getTime()-startDate.getTime()) + " 毫秒");
    }

    /**
     * 读取excel文件，使用SAX的事件方式读取xml,内存消耗小
     * @throws Exception
     */
    private static void readFileByEventModel() throws Exception {
        Date startDate = new Date();
        IRowReaderInterface reader = new IRowReader();
        String fileName = "37万.xlsx";
        //ExcelReaderUtil.readExcel(reader, "F://te03.xls");
        ExcelReaderUtil.readExcel(reader, fileDirectory+ fileName);
        Date endDate = new Date();
        System.out.println("------------used time is: " + (endDate.getTime()-startDate.getTime()) + " 毫秒");
    }

    /**
     * 写excel文件,可支持大数据量
     * POI提供了很好的支持，主要流程是：
     * 第一步构建工作薄和电子表格对象，
     * 第二步在一个流中构建文本文件，
     * 第三步使用流中产生的数据替换模板中的电子表格。
     * @throws Exception
     */
    private static void writeBigFileByXSSF() throws Exception{
        long start = System.currentTimeMillis();
        //构建excel2007写入器
        AbstractExcel2007Writer excel07Writer = new Excel2007WriterImpl();
        //调用处理方法
        String fileName = String.valueOf(new Date().getTime()) + ".xlsx";
        excel07Writer.process(fileDirectory+ fileName);
        long end = System.currentTimeMillis();
        System.out.println("....................."+(end-start)/1000);
    }

}
