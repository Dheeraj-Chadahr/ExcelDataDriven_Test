import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class WriteExcel {
    @Test(priority = 3)
    public void Write_ExcelData() throws InterruptedException,  IOException, IOException {
        File src=new File("C:\\Users\\HP\\Desktop\\HeyOliver.xlsx");
        FileInputStream fis=new FileInputStream(src);
        XSSFWorkbook wb=new XSSFWorkbook(fis);
        XSSFSheet sheet1=wb.getSheetAt(0);
        sheet1.getRow(0).createCell(2).setCellValue("123");
        sheet1.getRow(0).createCell(2).setCellValue("xyz");
        FileOutputStream fout=new FileOutputStream(src);
        wb.write(fout);
        wb.close();
    }
}
