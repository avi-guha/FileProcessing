package org.example;

import java.awt.*;
import java.io.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FileProcessing {

    public static void main(String[] args) {
        Frame frame = new Frame();
        FileDialog fileDialog = new FileDialog(frame, "Select .txt files", FileDialog.LOAD);
        fileDialog.setMultipleMode(true);
        fileDialog.setVisible(true);

        File[] txtFiles = fileDialog.getFiles();
        if (txtFiles == null || txtFiles.length == 0) {
            System.out.println("No files selected.");
            System.exit(0);
        }

        File parentDir = txtFiles[0].getParentFile();
        File fullDir = new File(parentDir, "full");
        File filteredDir = new File(parentDir, "filtered");
        fullDir.mkdirs();
        filteredDir.mkdirs();

        for (File txtFile : txtFiles) {
            String fullExcelPath = new File(fullDir, txtFile.getName().replace(".txt", "_full.xlsx")).getAbsolutePath();
            String filteredExcelPath = new File(filteredDir, txtFile.getName().replace(".txt", "_filtered.xlsx")).getAbsolutePath();

            convertTextToExcel(txtFile, fullExcelPath);
            filterExcel(fullExcelPath, filteredExcelPath);
        }

        System.out.println("Processing complete!");
        frame.dispose();
    }

    private static void convertTextToExcel(File txtFile, String excelPath) {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("FullData");

        try (BufferedReader br = new BufferedReader(new FileReader(txtFile))) {
            String line;
            int rowNum = 0;
            while ((line = br.readLine()) != null) {
                if (line.trim().isEmpty()) continue;

                String[] tokens = line.trim().split("\\s+");
                Row row = sheet.createRow(rowNum++);
                for (int i = 0; i < tokens.length; i++) {
                    Cell cell = row.createCell(i);
                    cell.setCellValue(tokens[i]);
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
            return;
        }

        try (FileOutputStream fos = new FileOutputStream(excelPath)) {
            workbook.write(fos);
            System.out.println("Created full Excel: " + excelPath);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void filterExcel(String fullExcelPath, String filteredExcelPath) {
        Workbook filteredWorkbook = new XSSFWorkbook();

        try (FileInputStream fis = new FileInputStream(fullExcelPath);
             Workbook fullWorkbook = new XSSFWorkbook(fis)) {

            Sheet fullSheet = fullWorkbook.getSheetAt(0);
            if (fullSheet == null) {
                System.err.println("No sheet found in " + fullExcelPath);
                return;
            }

            int headerRowIndex = 1;
            Row headerRow = fullSheet.getRow(headerRowIndex);
            if (headerRow == null) {
                System.err.println("Header row is empty in " + fullExcelPath);
                return;
            }

            int hcIndex = -1, erIndex = -1, hIndex = -1;


            for (int i = 0; i < headerRow.getLastCellNum(); i++) {
                Cell cell = headerRow.getCell(i);
                if (cell == null) continue;
                String header = cell.getStringCellValue().trim();


                if (header.equalsIgnoreCase("hc(nm)")) hcIndex = i + 2;
                if (header.equalsIgnoreCase("Er(GPa)")) erIndex = i + 2;
                if (header.equalsIgnoreCase("H(GPa)")) hIndex = i + 2;
            }

            if (hcIndex == -1 || erIndex == -1 || hIndex == -1) {
                System.err.println("Could not find required columns in " + fullExcelPath);
                return;
            }

            Sheet filteredSheet = filteredWorkbook.createSheet("FilteredData");
            Row filteredHeader = filteredSheet.createRow(0);
            filteredHeader.createCell(0).setCellValue("hc(nm)");
            filteredHeader.createCell(1).setCellValue("Er(GPa)");
            filteredHeader.createCell(2).setCellValue("H(GPa)");

            int filteredRowNum = 1;
            for (int i = headerRowIndex + 1; i <= fullSheet.getLastRowNum(); i++) {
                Row fullRow = fullSheet.getRow(i);
                if (fullRow == null) continue;

                Row newRow = filteredSheet.createRow(filteredRowNum++);
                newRow.createCell(0).setCellValue(getCellValueAsString(fullRow.getCell(hcIndex)));
                newRow.createCell(1).setCellValue(getCellValueAsString(fullRow.getCell(erIndex)));
                newRow.createCell(2).setCellValue(getCellValueAsString(fullRow.getCell(hIndex)));
            }

        } catch (IOException e) {
            e.printStackTrace();
            return;
        }

        try (FileOutputStream fos = new FileOutputStream(filteredExcelPath)) {
            filteredWorkbook.write(fos);
            System.out.println("Created filtered Excel: " + filteredExcelPath);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static String getCellValueAsString(Cell cell) {
        if (cell == null) return "";
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue().trim();
            case NUMERIC -> Double.toString(cell.getNumericCellValue());
            case BOOLEAN -> Boolean.toString(cell.getBooleanCellValue());
            case FORMULA -> cell.getCellFormula();
            default -> "";
        };
    }
}