import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class Main {
    public static void main( String[] args){
        System.out.println("Hello Gradle!");


        Workbook workbook = null;
        try{
            FileInputStream in = new FileInputStream("Purchaseorder2018ver.xlsx");
            workbook = WorkbookFactory.create(in);
            System.out.println("success access.");
        }catch (InvalidFormatException e){
            e.printStackTrace();
        }catch (IOException e){
            e.printStackTrace();
        }

        Sheet sheet = workbook.getSheet("発注書");
        Row row = null;

        List<Row> template = new ArrayList<Row>();
        for(int i = 1;row == null;i++){
            template.add(sheet.getRow(i));
        }
//        Cell cell = row.getCell()


        FileOutputStream out = null;
        try{
            out = new FileOutputStream("template_copy.xlsx");
            workbook.write(out);
        }catch (IOException e){
            e.printStackTrace();
        }finally{
            try {
                out.close();
                workbook.close();
            }catch (IOException e){
                e.printStackTrace();
            }
        }
    }
}
