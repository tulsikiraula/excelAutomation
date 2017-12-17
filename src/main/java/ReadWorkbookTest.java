import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.*;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

/**
 * Created by Tiger(AMIT) on 17-12-2017.
 */
public class ReadWorkbookTest {
    public static void main(String[] args) throws FileNotFoundException {
        XSSFWorkbook xssfWorkbook = null;
        FileInputStream fileInputStream = new FileInputStream(new File("Vinworkbook.xlsx"));
        try {
             xssfWorkbook = new XSSFWorkbook(fileInputStream);
        } catch (IOException io) {
            io.printStackTrace();
        }
        XSSFSheet sheet = xssfWorkbook.getSheet("Vin Definition");

        Iterator<Row> rowIterator = sheet.iterator();

        Map<Integer,ArrayList<Object>> details = new TreeMap<>();

        while(rowIterator.hasNext()){
            XSSFRow row = (XSSFRow)rowIterator.next();
            Iterator<Cell> cellIterator = row.cellIterator();
            ArrayList<Object> cellVal = new ArrayList<>();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                cellVal.add(cell.getStringCellValue());
            }
                details.put(row.getRowNum()+1, cellVal);
            }
        Set<Integer> set = details.keySet();
        ArrayList<Object> list = null;
        for (int setVal : set) {
             list = details.get(setVal);
            System.out.println(list);

        }

            
        }


    }
  

