import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.testng.Assert;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

@Slf4j
public class ExcelFunctions {

    public static void main(String args[]) throws IOException {


        //Write
        String fileWrite=System.getProperty("user.dir")+"/src/test/resources/Excel/WriteExcel.xlsx";
        ExcelFunctions ef=new ExcelFunctions();
        ef.writeExcel(fileWrite,"Sheet1",0,0,"First Name");
        ef.writeExcel(fileWrite,"Sheet1",0,1,"Last Name");
        ef.writeExcel(fileWrite,"Sheet1",1,0,"Jishnu");
        ef.writeExcel(fileWrite,"Sheet1",1,1,"Nambiar");
        ef.writeExcel(fileWrite,"Sheet1",2,0,"Parveen");
        ef.writeExcel(fileWrite,"Sheet1",2,1,"Banu");
        ef.writeExcel(fileWrite,"Sheet1",3,0,"Nishtha");
        ef.writeExcel(fileWrite,"Sheet1",3,1,"Kavin");
        ef.writeExcel(fileWrite,"Sheet1",4,0,"Puneeth");
        ef.writeExcel(fileWrite,"Sheet1",4,1,"Aishwarya");


        //Negative scenario
        String filePathExpected=System.getProperty("user.dir")+"/src/test/resources/Excel/ExpectedFile.xlsx";
        String filePathActualWrong=System.getProperty("user.dir")+"/src/test/resources/Excel/ActualFileWrong.xlsx";

        ef.verifyDataInAllSheets(filePathExpected,filePathActualWrong);


        //Postive Scenario
        String filePathActual=System.getProperty("user.dir")+"/src/test/resources/Excel/ActualFileWrong.xlsx";
        ef.verifyDataInAllSheets(filePathExpected,filePathActual);


    }
    /*
     Author:Jishnu Nambiar
     Description:Compare the data between the excel
     Input:String filePathExpected-File path of the expected excel
           String filePathActual-File path of the actual excel
     output:void
     */
    public void verifyDataInAllSheets(String filePathExpected, String filePathActual) throws IOException {
        System.out.println("Verify the excel has same data-->Started");
        //Verify the Sheets count are same before Validating Rows and Columns Count.
        verifySameNumberandNamesOfSheets(filePathExpected, filePathActual);
        //Verify the Rows and Columns count are same before Validating the data in cell.
        verifyRowsandColumnInSheets(filePathExpected, filePathActual);
        Workbook workbook1 = WorkbookFactory.create(new File(filePathExpected));
        Workbook workbook2 = WorkbookFactory.create(new File(filePathActual));
        //Since both the work book has same number of Sheets we are taking Workbook1's count
        try {
            int sheetCounts = workbook1.getNumberOfSheets();
            //We will convert all the cell data into String format
            DataFormatter df = new DataFormatter();

            for (int i = 0; i < sheetCounts; i++) {
                Sheet s1 = workbook1.getSheetAt(i);
                Sheet s2 = workbook2.getSheetAt(i);
                int rowCounts = s1.getPhysicalNumberOfRows();
                for (int j = 0; j < rowCounts; j++) {
                    Row row = s1.getRow(j);
                    Row row1 = s2.getRow(j);
                    //if whole row in between is null to increase the count of the loop
                    if (row == null && row1 == null) {
                        rowCounts = rowCounts + 1;
                        break;
                    } else {
                        //If only one Workbook has a null row but the other book has data in same Row
                        if (row == null) {
                            Assert.assertTrue(false, "WorkBook 1 is having empty row at " + (j + 1) + " in Sheet " + workbook1.getSheetAt(i) + "WorkBook 2 has Data at same position");
                        }
                        if (row1 == null) {
                            Assert.assertTrue(false, "WorkBook 2 is having empty row at " + (j + 1) + " in Sheet " + workbook1.getSheetAt(i) + "WorkBook 1 has data at same position");
                        }
                    }
                    //Get the column count in row
                    int cellCounts = s1.getRow(j).getPhysicalNumberOfCells();
                    for (int k = 0; k < cellCounts; k++) {
                        //Formating the cell data into String
                        String expectedValue = df.formatCellValue(s1.getRow(j).getCell(k));
                        String actualValue = df.formatCellValue(s2.getRow(j).getCell(k));
                        if (!expectedValue.equals(actualValue)) {
                            //Checking if the cell is empty/null
                            if (actualValue == "") {
                                actualValue = "Empty cell";
                            }
                            if (expectedValue == "") {
                                expectedValue = "Empty Cell";
                            }
                            Assert.assertTrue(expectedValue.equals(actualValue), "Expected Value in Row " + (j + 1) + " Column " + (k + 1) + " is " + expectedValue + " the same cell of actual file had " + actualValue);
                        }
                    }
                }

            }
            System.out.println("***Data is same in Excel***");
        }
        //Even if the code fails the Workbook should not be left open
        finally {
            workbook1.close();
            workbook2.close();
        }
        System.out.println("Verify the excel has same data-->Completed");

    }

    /*
      Author:Jishnu Nambiar
      Description:Verify the count of the row and columns in the excel
      Input:String filePathExpected-File path of the expected excel
            String filePathActual-File path of the actual excel
      output:void
      */
    public void verifyRowsandColumnInSheets(String filePathExpected, String filePathActual) throws IOException {
        System.out.println("Verifying if both work books have same number of Rows and Columns in a sheet--->Started");
        Workbook workbook1 = WorkbookFactory.create(new File(filePathExpected));
        Workbook workbook2 = WorkbookFactory.create(new File(filePathActual));
        try {
            int sheetCounts = workbook1.getNumberOfSheets();
            for (int i = 0; i < sheetCounts; i++) {
                Sheet s1 = workbook1.getSheetAt(i);
                Sheet s2 = workbook2.getSheetAt(i);
                int rowsInSheet1 = s1.getPhysicalNumberOfRows();
                int rowsInSheet2 = s2.getPhysicalNumberOfRows();
                Assert.assertTrue(rowsInSheet1 == rowsInSheet2, "Row count are not same");
                int j = 0;
                while ((j < rowsInSheet1)) {//Since Workbook1 is the expected file we are iterating with row1
                    Row row = s1.getRow(j);
                    if (row == null) {//if any row is null then we can skip and go to the next row
                        break;
                    }
                    j++;
                    int cellCounts1 = row.getPhysicalNumberOfCells();
                    int cellCounts2 = row.getPhysicalNumberOfCells();
                    Assert.assertTrue(cellCounts1 == cellCounts2, "Column count are not same");
                }
            }
            System.out.println("Both Workbooks same number of Rows and Columns in a sheet");
        } finally {
            workbook1.close();
            workbook2.close();
        }
        System.out.println("Verifying if both work books have same number of Rows and Columns in a sheet--->Completed");

    }

    /*
      Author:Jishnu Nambiar
      Description:Verify the Sheet in the excel
      Input:String filePathExpected-File path of the expected excel
            String filePathActual-File path of the actual excel
      output:void
      */
    public void verifySameNumberandNamesOfSheets(String filePathExpected, String filePathActual) throws IOException {
        System.out.println("Verifying if both work books have same number of sheets--->Started");
        Workbook workbook1 = WorkbookFactory.create(new File(filePathExpected));
        Workbook workbook2 = WorkbookFactory.create(new File(filePathActual));
        try {
            // Get total sheets count from first excel file
            int sheetsInWorkbook1 = workbook1.getNumberOfSheets();
            // Get total sheets count from second excel file
            int sheetsInWorkbook2 = workbook2.getNumberOfSheets();
            // Compare if both excel files have same number of sheets
            Assert.assertTrue(sheetsInWorkbook1 == sheetsInWorkbook2, "In Workbook1 sheet count is" + sheetsInWorkbook1 + "in Workbook1 sheet count is" + sheetsInWorkbook2);
            //Verifying sheet Names
            List<String> sheetsNameOfWb1 = new ArrayList<>();
            List<String> sheetsNameOfWb2 = new ArrayList<>();
            for (int i = 0; i < sheetsInWorkbook1; i++) {
                // Retrieving sheet names from both work books and adding to different lists
                sheetsNameOfWb1.add(workbook1.getSheetName(i));
                sheetsNameOfWb2.add(workbook2.getSheetName(i));
            }
            Collections.sort(sheetsNameOfWb1);
            Collections.sort(sheetsNameOfWb2);
            Assert.assertTrue(sheetsNameOfWb1.equals(sheetsNameOfWb2), "Sheet Name in Workbook1: "+sheetsNameOfWb1+" and WorkBook2: "+sheetsNameOfWb1+".Both are not same");
            System.out.println("The excel is having same number and name of Sheets");
        } finally {
            workbook1.close();
            workbook2.close();
        }
        System.out.println("Verifying if both work books have same number of sheets--->Completed");
    }

    /*
      Author:Jishnu Nambiar
      Description:Write data in the excel
      Input:String path-File path
            String sheetName-Sheet Name of the excel
            int rowNumber-Row in which data needs to be entered
            int columnNumber-Column in which data needs to be entered
            String value-Value to entered in cell
      output:void
      */
    public void writeExcel(String path, String sheetName, int rowNumber, int ColumnNumber, String value) throws IOException {
        File src = new File(path);
        FileInputStream fis = new FileInputStream(src);
        Workbook wb1 = WorkbookFactory.create(fis);
        try {
            Sheet xs = wb1.getSheet(sheetName);
            Row row = xs.getRow(rowNumber);
            if (row == null) {
                row = xs.createRow(rowNumber); // create a new row object if it is empty
            }
            xs.getRow(rowNumber).createCell(ColumnNumber).setCellValue(value);
            FileOutputStream fout = new FileOutputStream(src);
            wb1.write(fout);
            System.out.println("Data "+value+" added in row:"+rowNumber+" Column:"+ColumnNumber+" in sheet:"+sheetName);
        } finally {
            wb1.close();
        }

    }
}
