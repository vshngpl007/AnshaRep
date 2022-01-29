package ExcelTest;

import com.sun.org.apache.bcel.internal.generic.SWITCH;
import com.sun.scenario.effect.impl.sw.sse.SSEBlend_SRC_OUTPeer;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.print.DocFlavor;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
public class ExcelTest {
    public static void main(String[] args) throws IOException {

        FileInputStream fs = new FileInputStream("/Users/vishnugopal/Documents/AnushaTest.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(fs);
        XSSFSheet sheet = workbook.getSheetAt(0);
        Iterator sheetIT = sheet.iterator();
        //
        while (sheetIT.hasNext()) {
            XSSFRow row = (XSSFRow) sheetIT.next();

            Iterator cellIT = row.iterator();

            while (cellIT.hasNext()) {
String name = cellIT.next().toString();
             if(name.equalsIgnoreCase("Vishnu"))
             {
                 System.out.println(name);
                 while (cellIT.hasNext())
                 {
                     System.out.println(cellIT.next());
                 }
             }

            }

        }
    }
public static void AnushaNew()
{
    System.out.println("This is NewAnusha Changes");
}
}

