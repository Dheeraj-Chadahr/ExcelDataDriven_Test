import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;
import java.io.File;

public class newTest {
    public static void main(String[] args) throws Exception {
        FileInputStream fis = new FileInputStream(new File("C:/Users/Desktop-PC/dataDriven.xlsx"));
        HSSFWorkbook workbook = new HSSFWorkbook(fis);

        int sheets = workbook.getNumberOfSheets();
        for (int i=0; i<sheets; i++)
        {
            if (workbook.getSheetName(i).equalsIgnoreCase("testdata"))
            {
                System.out.println("Hello");
                HSSFSheet sheet = workbook.getSheetAt(i);

                Iterator<Row> rows=sheet.iterator(); //sheet is collection of rows
                Row firstRow= rows.next();
                Iterator<Cell>  ce=firstRow.cellIterator(); //row is collection of cells
                ce.next();
                int k=0;
                int coloumn = 0;
                while (ce.hasNext())
                {
                    Cell value=ce.next();
                    if (value.getStringCellValue().equalsIgnoreCase("Data2"))
                    {
                        coloumn=k;

                    }
                    k++;
                }
                System.out.println(coloumn);

            }
        }

    }
}
