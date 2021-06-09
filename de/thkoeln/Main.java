package de.thkoeln;



import org.apache.commons.compress.utils.Lists;
import org.apache.poi.ss.usermodel.*;

import javax.swing.*;
import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;

class Main {
    public static void main(String[] args) {
        try {
            // Variables
            //Optimizer myBasicOptimizer = new Optimizer();
            System.out.println("Optimizer instantiated...");

            // Select Excel File which the user wants to investigate regarding machine optimization potential
            // hier wieder ändern  String excelFileName = selectExcelFile();
            //hier wieder ändern   System.out.println("File via openFileDialog selected: " + excelFileName);

            // Read out data from excel sheet via Apache POI
            // hier wieder ändern   readExcelFileData(excelFileName);
            Sheet sheet = readExcelFileData("F:\\UNI\\AA SoSe 2021\\Projektmanagement 2\\PMII_Modell_v1\\ProductionSheet.xlsx");
            List<Row> rows = sortExcelFileData(sheet,19);

            // Optimize machine planning and scheduling
            //myBasicOptimizer.process();

            // Write optimized data back to excel sheet
            writeExcelFileData("F:\\UNI\\AA SoSe 2021\\Projektmanagement 2\\PMII_Modell_v2\\ProductionSheet NEU.xlsx", rows, 1);

            System.out.println("Optimization done.");
        } catch (Exception ex) {
            System.out.println(ex.toString());
        }
    }

    // Select Excel file which we want to read out and retrieve necessary machine data
    private static String selectExcelFile(){
        String strFilename = "";

        try {
            JFileChooser fileChooser = new JFileChooser();
            fileChooser.setCurrentDirectory(new File(System.getProperty("user.home")));

            int result = fileChooser.showOpenDialog(null);

            switch(result) {
                case JFileChooser.APPROVE_OPTION:
                    File selectedFile = fileChooser.getSelectedFile();

                    // user made a selection -> return selected file path
                    strFilename = selectedFile.toString();
                    break;

                case JFileChooser.CANCEL_OPTION:
                    // user made no selection -> shutdown application
                    System.out.println("No selection made. Shutting down application...");
                    System.exit(0);
                    break;

                default:
                    System.out.println("Something went wrong. Shutting down application...");
                    System.exit(0);
                    break;
            }
        } catch (Exception exFileSelection) {
            System.out.println("Error in selectExcelFile()");
            System.out.println("Error details: " + exFileSelection.toString());
        } finally {
            return strFilename;
        }
    }

    private static List<Row> sortExcelFileData(Sheet sheet, int sortColumn)
    {
        List<Row> rows = new ArrayList<>();
        try {

            for (int i = 1; i<sheet.getPhysicalNumberOfRows();i ++ )    // i = 2???? es werden keine 0-Zeilen erzeugt
            {
                rows.add(sheet.getRow(i)); //(new SortRow(sheet.getRow(i).getCell(19).getStringCellValue(), sheet.getRow(i)));
            }
            // ToDo: Sortieren bei Zahlenwerten?
            rows.sort((r1, r2) -> r1.getCell(sortColumn).getStringCellValue().compareTo(r2.getCell(sortColumn).getStringCellValue()));
        }
        catch (Exception exp) {
            System.out.println("Error in writeExcelFileData()");
            System.out.println("Error details: " + exp.toString());
        }
        return rows;
    }

    private static Sheet readExcelFileData(String fileName) {
        Sheet sheet = null;
        try {
            Workbook workbook = WorkbookFactory.create(new File(fileName));
            sheet = workbook.getSheetAt(1);
        } catch (Exception exReadExcelFile) {
            System.out.println("Error in readExcelFileData()");
            System.out.println("Error details: " + exReadExcelFile.toString());
        }
        return sheet;
    }

    private static void writeExcelFileData(String filename, List<Row> rows, int startRow){
        try {
            double Kabine = 0;
            int c = 0;
            Workbook workbook = WorkbookFactory.create(new File("F:\\UNI\\AA SoSe 2021\\Projektmanagement 2\\PMII_Modell_v1\\ProductionSheet Template.xlsx"));
            Sheet sheet = workbook.getSheet("Input");
            for (int i = 0; i < rows.size(); i++)
            {
                rows.get(i).getCell(0).setCellValue(i+1);
                String Farbe = rows.get(i).getCell(19).getStringCellValue();
                //String Blau = rows.get(0).getCell(20).getStringCellValue();
                //String Gelb = rows.get(0).getCell(25).getStringCellValue();
                //String Grün = rows.get(0).getCell(26).getStringCellValue();
                //String Rot = rows.get(0).getCell(27).getStringCellValue();
                //if(0<=i && i<22)
                if(Farbe.equals("Blau"))
                {
                    rows.get(i).getCell(3).setCellValue(1);
                    Kabine = Kabine+rows.get(i).getCell(14).getNumericCellValue();
                    c++;
                    if (Kabine>=30)
                    {
                        //for(int cc = c; c>0; c--)
                        //{
                        rows.get(i/*-c*/).getCell(5).setCellValue(Kabine);
                        //}
                        Kabine = 0;
                    }

                }
                else if(Farbe.equals("Gelb"))
                {
                    rows.get(i).getCell(3).setCellValue(2);
                }
                else if(Farbe.equals("Grün"))
                {
                    rows.get(i).getCell(3).setCellValue(1);
                }
                else if(Farbe.equals("Rot"))
                {
                    rows.get(i).getCell(3).setCellValue(2);
                }

                Row row = sheet.createRow(i + startRow);
                for (int x = 0; x < rows.get(i).getPhysicalNumberOfCells(); x++)
                {
                    if (rows.get(i).getCell(x).getCellType() == CellType.NUMERIC)
                    {
                        row.createCell(x).setCellValue(rows.get(i).getCell(x).getNumericCellValue());
                    }
                    else if (rows.get(i).getCell(x).getCellType() == CellType.STRING)
                    {
                        row.createCell(x).setCellValue(rows.get(i).getCell(x).getStringCellValue());
                    }
                    else if (rows.get(i).getCell(x).getCellType() == CellType.FORMULA)
                    {
                        String cellFormula = rows.get(i).getCell(x).getCellFormula();
                        String[] formulaSplitted = cellFormula.split("\\*");
                        for (int y = 0; y < formulaSplitted.length; y++) {
                            if (!formulaSplitted[y].matches("^[0-9]*$")) {
                                formulaSplitted[y] = formulaSplitted[y].replace(String.valueOf(rows.get(i).getRowNum() + 1), String.valueOf(i + startRow + 1));
                            }
                        }
                        row.createCell(x).setCellFormula(String.join("*", formulaSplitted));
                    }
                    else
                    {
                        System.out.println(rows.get(i).getCell(x).getCellType());
                    }
                }
            }
            try (FileOutputStream outputStream = new FileOutputStream(filename)) {
                workbook.write(outputStream);
            }
            System.out.println("Called writeExcelFileData()");
        } catch (Exception exWriteExcelFile) {
            System.out.println("Error in writeExcelFileData()");
            System.out.println("Error details: " + exWriteExcelFile.toString());
        }
    }

}

