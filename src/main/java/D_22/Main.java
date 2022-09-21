package D_22;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class Main {

    /*
    Napraviti fajl domaci22.xslx i u prve dve kolone upisati svoje ime pa prezime, u drugom redu
    upisati nasumicno ime i prezime. Napisati metodu u javi koja ce ispisati ta 2 imena i prezimena.
    Napisati metodu koja ce upisivati u xslx fajl u prvu i drugu kolonu ime i prezime. Treba da dodate
    jos 8 nasumicno generisanih imena i prezimena koriscenjem biblioteke Java Faker. Opet ih ispisati
    koristeci metodu za ispisivanje. Metodu za ispisivanje napraviti da radi tako da kada dobijete
    prazane vrednosti u redu da prestane petlja. Za kolone mozete fiksirati samo 2 kolone A i B.
     Upisivanje ne mora da bude dinamicki, moze samo da upisuje od treceg reda (drugi index) do 10og reda.

     */
    public static void main(String[] args) throws IOException {

        readExcel("domaci22.xlsx");
        try {
            writeExcel("test.xlsx");
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        }
    }

    private static void writeExcel(String filename) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("test");

        for (int i = 0; i < 1; i++) {
            XSSFRow row = sheet.createRow(i);
            for (int j = 0; j < 1; j++) {
                XSSFCell cell = row.createCell(j);
                cell.setCellValue("Goran");
            }
            for (int j = 1; j < 2; j++) {
                XSSFCell cell = row.createCell(j);
                cell.setCellValue("Pincir");
            }
        }
        // cell.setCellValue("Goran");

        FileOutputStream fileOutputStream = new FileOutputStream(new File(filename));
        workbook.write(fileOutputStream);
        fileOutputStream.close();


    }

    public static void readExcel(String path) {
        try {
            FileInputStream inputStream = new FileInputStream(new File("domaci22.xlsx"));

            XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
            XSSFSheet sheet = workbook.getSheet("Imena");

            for (int j = 0; j < 2; j++) {

                XSSFRow row = sheet.getRow(j);

                for (int i = 0; i < 2; i++) {
                    XSSFCell cell = row.getCell(i);
                    String name = cell.getStringCellValue();
                    System.out.println(name);
                }
            }
        } catch (FileNotFoundException ex) {
            System.out.println("FIleNotFound.class");
        } catch (IOException e) {
            // e.printStackTrace();
        }
    }



}
// Ovo meni nista ne funkcionise

