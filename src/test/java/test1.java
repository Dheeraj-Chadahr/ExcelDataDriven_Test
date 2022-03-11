//
//import com.sun.rowset.internal.Row;
//import javafx.scene.control.Cell;
//import org.apache.commons.compress.archivers.dump.InvalidFormatException;
//import org.apache.poi.ss.usermodel.CellType;
//import org.apache.poi.ss.usermodel.Sheet;
//import org.apache.poi.xssf.usermodel.XSSFCell;
//import org.apache.poi.xssf.usermodel.XSSFRow;
//import org.apache.poi.xssf.usermodel.XSSFSheet;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//import org.openqa.selenium.By;
//import org.openqa.selenium.WebDriver;
//import org.openqa.selenium.chrome.ChromeDriver;
//import org.testng.annotations.Test;
//import java.io.File;
//import java.io.IOException;
//
//
//public class test1{
//
//    @Test(priority = 1)
//    public static String[][] readExcel() throws InvalidFormatException, IOException, org.apache.poi.openxml4j.exceptions.InvalidFormatException {
//        String[][] data; // array for storing excel data
//        String filePath = "C:\\Users\\Desktop-PC\\Downloads\\dataDriven.xlsx";
//        File file = new File(filePath);
//        XSSFWorkbook workbook = new XSSFWorkbook(file);
//        XSSFSheet sheet = workbook.getSheet("Sheet1");
//        int noRows = sheet.getPhysicalNumberOfRows();
//        System.out.println("Rows are" + noRows);
//        data = new String[noRows][]; // initialize the size of array to no rows
//        // in excel
//        for (int i = 0; i < noRows; i++) {
//            XSSFRow row = sheet.getRow(i);
//            int noCells = row.getPhysicalNumberOfCells();
//            data[i] = new String[noCells];// initialize the size of col of array
//            // to no cell of excel
//            System.out.println("No of cells" + noCells);
//            for (int c = 0; c < noCells; c++) {
//                XSSFCell cell = row.getCell(c);
//                cell.setCellType(CellType.STRING);
//                String value = cell.getStringCellValue();
//                data[i][c] = value;
//                System.out.println("Value is" + value);
//            }
//        }
//        return data;
//    }}