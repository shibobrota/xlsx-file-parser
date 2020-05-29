package io.github.shibobrota;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ReadTwoTxtWriteXlsx {

    static void createXmlFromStringsTxt() throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        //Create a blank sheet
        XSSFSheet sheet = workbook.createSheet("Employee Data");

        File englishFile = new File("./english_input.txt"),
                hindiFile = new File("./hindi_input.txt");

        BufferedReader engBr = new BufferedReader(new FileReader(englishFile));
        BufferedReader hinBr = new BufferedReader(new FileReader(hindiFile));

        String stEng  = engBr.readLine(), stHin  = hinBr.readLine();
        Pattern patternDQ = Pattern.compile("\"(.*?)\"");
        Pattern patternLG = Pattern.compile(">(.*?)<");
        Matcher matcherEng, matcherHin, matcherEngVal, matcherHinVal;
        int rownum = 0;
        Row rowTitle = sheet.createRow(rownum++);
        Cell titleKeyCell = rowTitle.createCell(0), engCell = rowTitle.createCell(1), hinCell = rowTitle.createCell(2);
        titleKeyCell.setCellValue("KEY");
        engCell.setCellValue("English Translations");
        hinCell.setCellValue("Hindi Translations");

        while ((stEng) != null && (stHin) != null){

            matcherEng = patternDQ.matcher(stEng);
            matcherHin = patternDQ.matcher(stHin);


            if (matcherEng.find() && matcherHin.find())
            {
                if (matcherEng.group(1).equals(matcherHin.group(1))){
                    matcherEngVal = patternLG.matcher(stEng);
                    matcherHinVal = patternLG.matcher(stHin);

                    int cellnum = 0;
                    Row row = sheet.createRow(rownum++);
                    Cell keyCell = row.createCell(cellnum++);
                    keyCell.setCellValue(matcherEng.group(1));

                    matcherEngVal.find();
                    Cell engValCell = row.createCell(cellnum++);
                    engValCell.setCellValue(matcherEngVal.group(1));

                    matcherHinVal.find();
                    Cell hinValCell = row.createCell(cellnum++);
                    hinValCell.setCellValue(matcherHinVal.group(1));
                }
            }
            stEng  = engBr.readLine();
            stHin  = hinBr.readLine();
        }
        try
        {
            //Write the workbook in file system
            FileOutputStream out = new FileOutputStream(new File("merged.xlsx"));
            workbook.write(out);
            out.close();
            System.out.println("File written successfully on disk.");
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
    }
}
