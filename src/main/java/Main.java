import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

public class Main {
    public static void main( String[] args){

        Workbook workbook = null;
        try{
            FileInputStream in = new FileInputStream("template.xlsx");
            workbook = WorkbookFactory.create(in);
            System.out.println("[INFO]Access Success.");
        }catch (Exception e){
            e.printStackTrace();
        }

        Sheet sheet = workbook.getSheet("発注書");

        List<Row> template = new ArrayList<>();

//        template.add(sheet.getRow(0));//1番上の空行
        //テンプレ全コピー(最後に空行入る？)
        int count = 1;
        while(true){
            template.add(sheet.getRow(count));
            if(template.get(count - 1) == null)break;
            count++;
        }
//        for(int i = 1;template.get(i - 1) != null;i++){
//            template.add(sheet.getRow(i));
//        }
        System.out.println("[INFO]size:" + template.size());

        List<Row> template_copy = new ArrayList<>();
        //テンプレの一番下から1行空けてcreateRow
        for(int i = template.size();i < 2*template.size();i++){
            template_copy.add(sheet.createRow(i));
        }

        Cell origin_cell;
        Cell target_cell;
        //i = row , j = column
        for(int i = 0;i < template.size() - 1;i++) {
            for(int j = 0;j < 13;j++) {//todo
                try {
//                    System.out.println("i:" + i + "\tj:" + j);
                    origin_cell = template.get(i).getCell(j);
                    target_cell = template_copy.get(i).createCell(j);

                    if (origin_cell != null) {
                        target_cell.setCellStyle(origin_cell.getCellStyle());

                        switch (origin_cell.getCellType()){
                            case Cell.CELL_TYPE_STRING:
                                target_cell.setCellValue(origin_cell.getStringCellValue());
                                System.out.println(origin_cell.getStringCellValue());
                                break;
                            case Cell.CELL_TYPE_FORMULA:
                                target_cell.setCellFormula(origin_cell.getCellFormula());
                                System.out.println(origin_cell.getCellFormula());
                                break;
                            default:
                                break;
                        }
                    }
                }
                catch (Exception e){
                    e.printStackTrace();
                }
            }
        }
        System.out.println("\n[INFO]copy_finish");
//        row = sheet.getRow(2);
//        Row copy_row = sheet.createRow(19);
//        copy_row.createCell(2).setCellValue(row.getCell(2).getStringCellValue());
//        copy_row.getCell(2).setCellStyle(row.getCell(2).getCellStyle());

//        Cell cell = row.createCell(0);
//        cell.setCellValue(template.size());


        FileOutputStream out = null;
        try{
            out = new FileOutputStream("template_copy.xlsx");
            workbook.write(out);
        }catch (Exception e){
            e.printStackTrace();
        }finally{
            try {
                out.close();
                workbook.close();
            }catch (Exception e){
                e.printStackTrace();
            }
        }
    }
}
