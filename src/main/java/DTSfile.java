import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;

public class DTSfile implements ExcelFile {
    private static final String file = "D://stock/DTS.xls";
    private static final ArrayList<String> rezult = new ArrayList<String>();

    public ArrayList<String> findMatches(String request) throws IOException {
        HSSFWorkbook myExcelBook = new HSSFWorkbook(new FileInputStream(file));
        HSSFSheet myExcelSheet = myExcelBook.getSheetAt(0);
        HSSFRow row ;

        System.out.println("----------------------------------------------------------");
        System.out.println(" Ä Ò Ñ ");

        if(request.contains(" ")) {
            String request1 = request.substring(0, request.lastIndexOf(" "));
            String request2 = request.substring(request.lastIndexOf(" ") + 1);
            for (int i = 1; i <= myExcelSheet.getPhysicalNumberOfRows() - 1; i++) {
                row = myExcelSheet.getRow(i);
                try {
                    String name = row.getCell(2).getStringCellValue();
                    if (name.toUpperCase().contains(request1.toUpperCase()) && name.toUpperCase().contains(request2.toUpperCase())) {
                        Double cuantaty = row.getCell(3).getNumericCellValue();
                        Double price = row.getCell(4).getNumericCellValue();

                        String answer = name + " " + String.format("%.0f", cuantaty * 1000) + " ì  " + String.format("%.2f", price / 1000) + "ãðí/ì";
                        System.out.println(answer);
                        rezult.add(answer+" ÄÒÑ");
                    }
                } catch (IllegalStateException e){}
                catch (NullPointerException e5){}
            }
        }
        else{
            for (int i = 1; i <= myExcelSheet.getPhysicalNumberOfRows() - 1; i++) {
                row = myExcelSheet.getRow(i);
                try {
                    String name = row.getCell(2).getStringCellValue();
                    if (name.toUpperCase().contains(request.toUpperCase())) {
                        Double cuantaty = row.getCell(3).getNumericCellValue();
                        Double price = row.getCell(4).getNumericCellValue();

                        String answer = name + " " + String.format("%.0f", cuantaty * 1000) + " ì  " + String.format("%.2f", price / 1000) + "ãðí/ì";
                        System.out.println(answer);
                        rezult.add(answer+" ÄÒÑ");
                    }
                }
                catch (IllegalStateException e){}
                catch (NullPointerException e5){}
            }
        }


        myExcelBook.close();

        return rezult;
    }
}
