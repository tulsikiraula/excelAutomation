import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;

/**
 * Created by Tiger(AMIT) on 16-12-2017.
 */
public class CreateWorkbookTest {

    public static void main(String[] args) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook( );
        FileOutputStream outputStream = new FileOutputStream(new File("Vinworkbook.xlsx"));
        XSSFSheet sheet1 = workbook.createSheet("Vin Definition");
        sheet1.createRow(51);
        Map<Integer, Object[]> details = new HashMap<Integer, Object[]>();
        details.put(1,new Object[]{"Model","Version","Transmision","StartDate","EndDate"});
        details.put(2,new Object[]{"Honda","1.8","Manaul","2011","2015"});
        details.put(3,new Object[]{"Honda","1.8","Automatic","2016","2017"});

        XSSFRow row;
        int rowid=0;

        Set<Integer> keys = details.keySet();
        for (int key : keys){
            row=sheet1.createRow(rowid++);

            Object detail[]=details.get(key);
            int cellId=0;
            for (Object val: detail) {
                Cell cell = row.createCell(cellId++);
                cell.setCellValue((String)val);
                
            }
            
        }
        workbook.write(outputStream);
        System.out.println("sheet created");


    }



}
