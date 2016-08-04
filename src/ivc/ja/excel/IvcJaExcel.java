/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package ivc.ja.excel;

/**
 *
 * @author n-dan_000
 */
import java.io.FileNotFoundException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class IvcJaExcel {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        try {
            writeToExcel("workbook.xlsx");
        } catch (IOException ex) {
            Logger.getLogger(IvcJaExcel.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
 
    public static void writeToExcel (String filePath) throws IOException {
            int value;
            Workbook book = new XSSFWorkbook();
            Sheet[] sheets = new Sheet[2];
            Row[] rows = new Row[3];
            
            for (Sheet sheet : sheets) {
                sheet = book.createSheet();
                value = 1;
                for (int i = 1; i < 4; i++) {
                rows[i-1] = sheet.createRow(i);
                for (int j = 1; j < 4; j++) {
                    rows[i-1].createCell(j).setCellValue(String.valueOf(value));
                    value++;
                }
            }
        }
            //book.write();
            try {
            book.write(new FileOutputStream(filePath));
        } catch (FileNotFoundException fnfe) {
                System.out.println(fnfe.toString());
        } catch (Exception e){
                System.out.println(e.toString());
        }
            
            book.close();
    }
    
}
