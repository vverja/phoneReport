package com.eskaro.javautils;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

public class ExcelExchange {
    private static Map<String, Map<String, Double>> table = new HashMap<>();
    public static void main(String[] args) throws IOException {
        if (args.length==0)
            throw new IOException("argument is not set!");
        for (int i = 0; i < args.length-1; i++) {
            getMappedData(getFile(args[i]));
        }
        writeMappedData(getFile(args[args.length-1]));

        
        }
        
        private static void getMappedData(File filePath) throws IOException{
            try(FileInputStream file = new FileInputStream(filePath)){

                Workbook wb = new HSSFWorkbook(file);

                Sheet sh = wb.getSheetAt(0);
                Row headerRow = sh.getRow(2);
                for (int j = 3; j < sh.getLastRowNum()+1; j++) {
                    for (int i = 1; i < headerRow.getLastCellNum()-1; i++) {
                        Map<String, Double> column;
                        String keyRow = headerRow.getCell(i).getStringCellValue().replaceAll("[^0-9]", "");
                        {
                            if (table.get(keyRow) == null) {
                                column = new HashMap<>();
                            }
                            else{
                                column = table.get(keyRow);
                            }
                            if (sh.getRow(j).getCell(i).getCellType() == CellType.NUMERIC)
                                column.put(sh.getRow(j).getCell(0).getStringCellValue(), sh.getRow(j).getCell(i).getNumericCellValue());
                            else if (sh.getRow(j).getCell(i).getStringCellValue().equals(" "))
                                column.put(sh.getRow(j).getCell(0).getStringCellValue(), 0.0);

                        }
                        table.put(keyRow,column);
                    }
                }
            }
        }
    public static void writeMappedData(File filePath) throws IOException {
        FileInputStream file = new FileInputStream(filePath);
            Workbook wb = new XSSFWorkbook(file);
            Sheet mainSheet =  wb.getSheetAt(0);
            Map<String, Integer> phoneAdressList = new HashMap<>();
            Map<String, Integer> expenceAdressList = new HashMap<>();

            for (Map.Entry<String, Map<String, Double>> mapRow: table.entrySet()
                 ) {
                final String keyPhone = mapRow.getKey();
                Cell cell = findCell(keyPhone, mainSheet);
                if (cell != null) {
                    phoneAdressList.put(keyPhone, cell.getRowIndex());
                    for (Map.Entry<String, Double> expense: table.get(keyPhone).entrySet()
                    ) {
                        String expenseKey = expense.getKey();
                        if (expenceAdressList.containsKey(expenseKey))
                            continue;
                        cell = findCell(expenseKey, mainSheet);
                        if(cell!=null) {
                            expenceAdressList.put(expenseKey, cell.getColumnIndex());
                        }else{
                            System.out.println("В шаблоне не найдена затрата - " + expenseKey);
                        }
                    }
                } else {
                    System.out.println("В шаблоне не найден телефон - " + keyPhone);
                }
            }
            for (Map.Entry<String, Integer> phone: phoneAdressList.entrySet()
                 ) {
                Row row = mainSheet.getRow(phone.getValue());
                for (Map.Entry<String, Double> expences: table.get(phone.getKey()).entrySet()
                ){
                    try {
                        row.getCell(expenceAdressList.get(expences.getKey())).setCellValue(expences.getValue());
                    }catch(NullPointerException e){
                        System.out.printf("Затрата %s не найдена в шаблоне\n", expences.getKey());
                    }
                }
            }
            wb.setForceFormulaRecalculation(true);
            Date now = new Date();
            SimpleDateFormat sf  = new SimpleDateFormat("dd_MM_yy");
            FileOutputStream outputStream = new FileOutputStream(String.format("report%s.xlsx",sf.format(now)));
            wb.write(outputStream);
            wb.close();
            file.close();
            outputStream.close();

    }

    private static File getFile(String pathToXLSFile) {
        Path path = Paths.get(pathToXLSFile);
        return new File(path.toAbsolutePath().toString());
    }
    private static Cell findCell(String key, Sheet sheet){

        for (Row row: sheet
             ) {
            for (Cell cell: row
                 ) {
                if(cell.getCellType()==CellType.STRING) {
                    if (cell.getStringCellValue().equals(key))
                        return cell;
                }
                else if(cell.getCellType()==CellType.NUMERIC){
                    if(Double.parseDouble(key)==cell.getNumericCellValue())
                        return cell;
                }
            }
        }
        return null;
    }
}

