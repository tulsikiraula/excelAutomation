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
        Map<Integer,ArrayList<Object>> vinData = null;
        Map<Integer,ArrayList<Object>> GGData = null;

        vinData = readWorkbook("Vinworkbook.xlsx");
        GGData = readWorkbook("GG_Cars.xlsx");

        showDetails(vinData,GGData);
        System.out.println("good");

    }

    private static Map<Integer,ArrayList<Object>> readWorkbook(String fileName ) throws FileNotFoundException {

        XSSFWorkbook xssfWorkbook = null;
        FileInputStream fileInputStream = new FileInputStream(new File(fileName));
        try {
             xssfWorkbook = new XSSFWorkbook(fileInputStream);
        } catch (IOException io) {
            io.printStackTrace();
        }
        //XSSFSheet sheet = xssfWorkbook.getSheet("Vin Definition");
        XSSFSheet sheet = xssfWorkbook.getSheetAt(0);
        Iterator<Row> rowIterator = sheet.iterator();

        Map<Integer,ArrayList<Object>> details = new HashMap<>();

        while(rowIterator.hasNext()){
            XSSFRow row = (XSSFRow)rowIterator.next();
            Iterator<Cell> cellIterator = row.cellIterator();
            ArrayList<Object> cellVal = new ArrayList<>();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                switch(cell.getCellType()){
                    case Cell.CELL_TYPE_NUMERIC:
                        cellVal.add(cell.getNumericCellValue());
                        break;
                    case Cell.CELL_TYPE_STRING:
                        cellVal.add(cell.getStringCellValue());
                        break;
                    default:

                }
            }
                details.put(row.getRowNum(), cellVal);
            }
        return details;
    }

    private static void showDetails(Map<Integer, ArrayList<Object>> vin, Map<Integer, ArrayList<Object>> gg) {
        Set<Integer> vinSet = vin.keySet();
        Set<Integer> ggSet = gg.keySet();

        ArrayList<Object> vinList = null;
        ArrayList<Object> ggList = null;

        Map<Integer,ArrayList<Object>> newSheet = new HashMap<>();
        ArrayList<Object> newList = new ArrayList<>();
        int i=0;
        for (int vinVal : vinSet) {
             vinList = vin.get(vinVal+1);

             if (vinList != null) {
                 for (int ggVal : ggSet) {
                     ggList = gg.get(ggVal + 1);
                     if (ggList != null) {
                         if (ggList.get(0).toString().equalsIgnoreCase(vinList.get(0).toString()) && ggList.get(1) == vinList.get(1) &&
                                 ggList.get(5).toString().equalsIgnoreCase(vinList.get(2).toString()) && (int) ggList.get(6) <= (int) vinList.get(4) &&
                                 (int) ggList.get(7) >= (int) vinList.get(3)) {
                             newList.add(ggList.get(0));
                             newList.add(ggList.get(1));
                             newSheet.put(++i, newList);
                         }
                     }

                 }
             }
        }
    }


}
  

