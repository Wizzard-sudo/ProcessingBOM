import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class main {
    public static void main(String[] args) throws IOException {
        readFromExcel("D:\\Java\\Excel\\1234.xlsx");
    }



    public static void readFromExcel(String file) throws IOException {

        int numberColumnWithPN = 0;
        int numberColumnWithManufacturer = 1;
        int numberSheetDocument = 0;
        int numberColumnWithFindPN = 3;


        List<String> PNList = new ArrayList<>();

        XSSFWorkbook myExcelBook = new XSSFWorkbook(new FileInputStream(file));
        XSSFSheet myExcelSheet = myExcelBook.getSheetAt(numberSheetDocument);

        int r = 2;
        XSSFRow row = myExcelSheet.getRow(r);

        String PN = new String();

        while (myExcelSheet.getRow(r) != null) {
            if(PN != null)
            {
                PN = row.getCell(numberColumnWithPN).toString();
                PN = PN.replaceAll(" ", "")
                        .replaceAll("[^\\x00-\\x7F]", "")
                        .replaceAll(".115", "");
                System.out.print(PN + ", ");
                PNList.add(PN);
                if(row.getCell(numberColumnWithManufacturer) != null)
                    System.out.println(row.getCell(numberColumnWithManufacturer));
            }
            r++;
            row = myExcelSheet.getRow(r);
        }
        myExcelBook.close();

          System.out.println("--!--");

        int count = 2;
        row = myExcelSheet.getRow(count);

        for (String name:PNList) {
            System.out.println("---");
            System.out.println(name);
            StringBuffer nameSB = new StringBuffer(name);
            List<String> DNList = findPN(nameSB);
            System.out.println(DNList);

            row.getCell(numberColumnWithFindPN).setCellValue(DNList.toString());

            count++;
            row = myExcelSheet.getRow(count);
        }
        myExcelBook.close();
    }





    public static List<String> findPN(StringBuffer name) throws IOException {

        XSSFWorkbook myExcelBook = new XSSFWorkbook(new FileInputStream("D:\\Java\\Excel\\All PN.xlsx"));
        XSSFSheet myExcelSheet = myExcelBook.getSheetAt(0);

        int clarNum = (int) Math.round(name.length() * 0.33);

        List<String> DNList = new ArrayList<>();

        int r = 1;
        int clN = 0;
        XSSFRow row = myExcelSheet.getRow(r);

        while (DNList.isEmpty() & clN <= clarNum) {
            while (myExcelSheet.getRow(r) != null) {
                if (row.getCell(4).toString().contains(name.toString())) {
                    DNList.add(String.valueOf(row.getCell(1)));
                }
                r++;
                row = myExcelSheet.getRow(r);
            }
            if(DNList.isEmpty()) {
                name.delete(name.length() - 1, name.length());
                r = 1;
                row = myExcelSheet.getRow(r);
                clN++;
            }
        }
        if (DNList.isEmpty())
            DNList.add("PN не был найден");

        myExcelBook.close();

        return DNList;
    }
}
