import javafx.application.Application;

import javafx.fxml.FXMLLoader;
import javafx.geometry.Pos;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.control.Label;
import javafx.stage.Stage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.net.URL;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;

public class main extends Application {

    public static void main(String[] args) throws IOException {
        Application.launch();
        //readFromExcel("D:\\Java\\Excel\\1234.xlsx");
    }

    @Override
    public void start(Stage stage) throws Exception {
        FXMLLoader loader = new FXMLLoader();
        URL xmlUrl = getClass().getResource("/mainScene.fxml");
        loader.setLocation(xmlUrl);
        MainSceneController controller = loader.getController();
        Parent root = loader.load();

        stage.setScene(new Scene(root));
        stage.show();
    }

    public static void readFromExcel(String file,
                                     int numberColumnWithPN,
                                     int numberColumnWithManufacturer,
                                     double accuracySearch, String nameFileNew) throws IOException {

//        int numberColumnWithPN = 0;
//        int numberColumnWithManufacturer = 1;
//        int numberSheetDocument = 0;
//        double accuracySearch = 0.33;
//        int numberColumnWithPN = 1;
//        int numberColumnWithManufacturer = 2;
        int numberSheetDocument = 0;
//        double accuracySearch = 0.33;


        String nameFile = file.substring(0, file.length() - nameFileNew.length() - 2) + "NEW.xlsx";


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

                if(row.getCell(numberColumnWithManufacturer) != null) {
                    System.out.println(row.getCell(numberColumnWithManufacturer));
                }
            }
            r++;
            row = myExcelSheet.getRow(r);
        }
        myExcelBook.close();

                System.out.println("--!--");

        r = 2;

        List<List<String>> DNs = new ArrayList<>();

        for (String name:PNList) {
            XSSFWorkbook myExcelBookFind = new XSSFWorkbook(new FileInputStream("D:\\Java\\Excel\\All PN.xlsx"));

            System.out.println("---");
            System.out.println(name);
            StringBuffer nameSB = new StringBuffer(name);
            List<String> DNList = findPN(nameSB, accuracySearch, myExcelBookFind);
            System.out.println(DNList);
            DNs.add(DNList);

            myExcelBook.close();
        }

        XSSFWorkbook copyBook = copyWorkbookFile(myExcelBook, r, DNs);
        XSSFSheet copySheet = copyBook.getSheet("BOM");


        //writeWorkbook(myExcelBook, file);
        myExcelBook.close();

        r = 2;
        int i = 0;
        row = copySheet.getRow(r);

        while (row.getCell(i) != null)
            copySheet.autoSizeColumn(++i);


        copyBook.write(new FileOutputStream(nameFile));
        copyBook.close();
    }





    public static List<String> findPN(StringBuffer name, double accuracySearch, XSSFWorkbook myExcelBook) throws IOException {


        XSSFSheet myExcelSheet = myExcelBook.getSheetAt(0);

        int clarNum = (int) Math.round(name.length() * accuracySearch);

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



        return DNList;
    }

    public static XSSFWorkbook copyWorkbookFile(XSSFWorkbook originWorkbook, int rowBegin, List<List<String>> DNs){

        XSSFSheet originSheet = originWorkbook.getSheetAt(0);
        XSSFWorkbook copyWorkbook = new XSSFWorkbook();
        XSSFSheet copySheet = copyWorkbook.createSheet("BOM");

        int cellNum = 0;
        int rowConst = rowBegin;

        XSSFRow origenRow = originSheet.getRow(rowBegin);
        XSSFCell originCell;
        XSSFRow copyRow = copySheet.createRow(rowBegin);
        XSSFCell copyCell = copyRow.createCell(cellNum);

        while (origenRow != null){
            originCell = origenRow.getCell(cellNum);

            while (originCell != null){
                copyCell = copyRow.createCell(cellNum);

                copyCell.setCellValue(originCell.toString());
                System.out.print(originCell + " ");
                originCell = origenRow.getCell(++cellNum);
            }

            copyCell = copyRow.createCell(++cellNum);
            copyCell.setCellValue(DNs.get(rowBegin - rowConst).toString().substring(1, DNs.get(rowBegin - rowConst).toString().length() - 1));

            rowBegin++;
            origenRow = originSheet.getRow(rowBegin);
            copyRow = copySheet.createRow(rowBegin);
            cellNum = 0;

            System.out.println();
        }

        return copyWorkbook;
    }
}