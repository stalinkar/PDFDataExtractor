package org.example;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Date;

public class PDFTableExtractor {

    public static void main(String[] args) {
        String pdfFilePath = "inputFile.pdf";
        String excelFilePath = "outputFile.xlsx";

        try {
            PDDocument document = PDDocument.load(Files.newInputStream(Paths.get(pdfFilePath)));
            PDFTextStripper pdfStripper = new PDFTextStripper();
            String text = pdfStripper.getText(document);

            Workbook workbook = getWorkbook(text, excelFilePath);

            workbook.close();
            document.close();

            System.out.println("Tables extracted from PDF and saved to Excel.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static Workbook getWorkbook(String text, String excelFilePath) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Tables");

        // Split text into paragraphs
        String[] paragraphs = text.split("^\\s*$");

        int rowNum = 0;
        for (String para : paragraphs) {
            Row row = sheet.createRow(rowNum);
                Cell cell = row.createCell(rowNum);
                cell.setCellValue(para);
            rowNum++;
        }

        // Write to Excel file
        try (FileOutputStream fileOut = new FileOutputStream(excelFilePath)) {
            workbook.write(fileOut);
        }
        return workbook;
    }
}
