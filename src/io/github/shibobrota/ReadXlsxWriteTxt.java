package io.github.shibobrota;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileWriter;
import java.util.ArrayList;
import java.util.Iterator;

public class ReadXlsxWriteTxt {

    public static final String INPUT_XLSX_PATH = "./Status_and_error_codes.xlsx";
    public static final String OUTPUT_ENGLISH_TXT_PATH = "newEnglishStrings.txt";
    public static final String OUTPUT_HINDI_TXT_PATH = "newHindiStrings.txt";

    public static void main(String[] args) {
        createStringsFileFromXml();
    }

    static void createStringsFileFromXml() {
        ArrayList<String> keyStrings = new ArrayList<>();
        try {
            FileInputStream xlsxFIleStream = new FileInputStream(INPUT_XLSX_PATH);
            Workbook workbook = new XSSFWorkbook(xlsxFIleStream);
            Sheet firstSheet = workbook.getSheetAt(0);
            Iterator sheetIterator = firstSheet.iterator();
            FileWriter stringsEnglishFileWriter = new FileWriter(OUTPUT_ENGLISH_TXT_PATH);
            FileWriter stringsHindiFileWriter = new FileWriter(OUTPUT_HINDI_TXT_PATH);
            while (sheetIterator.hasNext()) {
                Row row = (Row) sheetIterator.next();
                Cell englishCell = row.getCell(6); //English Messages
                Cell hindiCell = row.getCell(7); //HindiMessages
                if (englishCell != null &&
                        CellType.STRING == englishCell.getCellType() &&
                        !englishCell.getStringCellValue().trim().equals("") &&
                        !englishCell.getStringCellValue().trim().equals("--") &&
                        hindiCell != null &&
                        CellType.STRING == hindiCell.getCellType() &&
                        !hindiCell.getStringCellValue().trim().equals("") &&
                        !hindiCell.getStringCellValue().trim().equals("--")
                ){
                    String key = englishCell.getStringCellValue().toLowerCase()
                            .replaceAll(" ","_")
                            .replaceAll("\\.","")
                            .replaceAll(",","");
                    if (!keyStrings.contains(key)) {
                        keyStrings.add(key);
                        //Put it into String File
                        stringsEnglishFileWriter.append("<string name=\""+key+"\">"+englishCell.getStringCellValue().trim()+"</string>\n");
                        stringsHindiFileWriter.append("<string name=\""+key+"\">"+hindiCell.getStringCellValue().trim()+"</string>\n");
                    }
                }
            }
            stringsEnglishFileWriter.close();
            stringsHindiFileWriter.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
