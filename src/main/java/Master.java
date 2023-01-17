import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Master {
    public Master() throws IOException {
        initTable();
        printFileData();
    }

    private void initTable() throws IOException {

//        Let's create a method that writes a list of persons to a sheet titled “Persons”.

//        First, we will create and style a header row that contains “Name” and “Age” cells:

        Workbook workbook = new XSSFWorkbook();

        Sheet sheet = workbook.createSheet("Persons");
        sheet.setColumnWidth(0, 6000);
        sheet.setColumnWidth(1, 4000);

        Row header = sheet.createRow(0);

        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        XSSFFont font = ((XSSFWorkbook) workbook).createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 16);
        font.setBold(true);
        headerStyle.setFont(font);

        Cell headerCell = header.createCell(0);
        headerCell.setCellValue("Name");
        headerCell.setCellStyle(headerStyle);

        headerCell = header.createCell(1);
        headerCell.setCellValue("Age");
        headerCell.setCellStyle(headerStyle);

//        Next, let's write the content of the table with a different style:

        CellStyle style = workbook.createCellStyle();
        style.setWrapText(true);

        Row row = sheet.createRow(1);
        Cell cell = row.createCell(0);
        cell.setCellValue("John Smith2");
        cell.setCellStyle(style);

        cell = row.createCell(1);
        cell.setCellValue(20);
        cell.setCellStyle(style);
//        fill 100 rows of data
        for (int i = 2; i < 100; i++) {
            row = sheet.createRow(i);
            cell = row.createCell(0);
            cell.setCellValue(i);//+"");
            cell = row.createCell(1);
            cell.setCellValue(i + 1);//+"");

        }

//        Finally, let's write the content to a “temp.xlsx” file in the current directory and close the workbook:

//        File currDir = new File(".");
//        String path = currDir.getAbsolutePath();
//        String fileLocation = path.substring(0, path.length() - 1) + "temp.xlsx";

        FileOutputStream outputStream = new FileOutputStream(getFilePath());
        workbook.write(outputStream);
        workbook.close();
    }

    private String getFilePath() {
        File currDir = new File(".");
        String path = currDir.getAbsolutePath();
        String filePath = path.substring(0, path.length() - 1) + "temp.xlsx";
        System.out.println(filePath);
        return filePath;
    }

    private void printFileData() throws IOException {
        FileInputStream file = new FileInputStream(new File(getFilePath()));
        Workbook workbook = new XSSFWorkbook(file);
        Sheet sheet = workbook.getSheetAt(0);
//        fill the map
        Map<Integer, List<String>> data = new HashMap<>();
        int i = 0;
        for (Row row : sheet) {
            data.put(i, new ArrayList<String>());
            for (Cell cell : row) {
                switch (cell.getCellType()) {
                    case STRING -> data.get(i).add(cell.getRichStringCellValue().getString());
                    case NUMERIC -> {
                        if (DateUtil.isCellDateFormatted(cell)) {
                            data.get(i).add(cell.getDateCellValue() + "");
                        } else {
                            data.get(i).add(cell.getNumericCellValue() + "");
                        }
                    }
                    case BOOLEAN -> data.get(i).add(cell.getBooleanCellValue() + "");
                    case FORMULA -> data.get(i).add(cell.getCellFormula() + "");
                    default -> data.get(i).add(" ");
                }
            }
            i++;
        }
//        print map
        for (int j = 0; j < data.size(); j++) {
            List row = data.get(j);
            for (int k = 0; k < row.size(); k++) {
                System.out.print(row.get(k) + "\t");
            }
            System.out.println();
        }
    }

    public static void main(String[] args) throws IOException {
        new Master();
    }
}
