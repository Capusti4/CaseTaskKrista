package com.example.CaseTask;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Main {
    static FileInputStream fis;
    static Workbook workbook;
    static FileOutputStream out;

    public static void main(String[] args) throws IOException {
        {
            fis = new FileInputStream("C:\\Users\\Capusti4\\Desktop\\test.xlsx");
            workbook = new XSSFWorkbook(fis);
            out = new FileOutputStream("C:\\Users\\Capusti4\\Desktop\\test_result.xlsx");
            Map<Integer, Object> criteria = new HashMap<>();
            List<List<Cell>> cRows = new ArrayList<>();
            List<List<Row>> groups = new ArrayList<>();

            System.out.println("Reading criteria");
            Sheet sheet = workbook.getSheetAt(0);
            for (Cell cell : sheet.getRow(0)) {
                criteria.put(cell.getColumnIndex(), cell.getStringCellValue().toLowerCase());
            }

            System.out.println("Reading data");
            for (Row row : sheet) {
                if (row.getRowNum() == 0) {
                    continue;
                }

                List<Cell> cRow = new ArrayList<>();
                for (Cell cell : row) {
                    if (!criteria.containsKey(cell.getColumnIndex())) {
                        cRow.add(cell);
                    }
                }
                cRows.add(cRow);
            }

            System.out.println("Grouping data");
            while (!cRows.isEmpty()) {
                List<Cell> row = cRows.getFirst();
                List<Row> group = new ArrayList<>();
                List<List<Cell>> cRowsToDelete = new ArrayList<>();

                group.add(row.getFirst().getRow());
                cRows.removeFirst();

                for (List<Cell> cRow : cRows) {
                    if (!equals(row, cRow)) {
                        continue;
                    }
                    group.add(cRow.getFirst().getRow());
                    cRowsToDelete.add(cRow);
                }
                cRows.removeAll(cRowsToDelete);
                groups.add(group);
            }

            System.out.println("Calculating criteria");
            for (List<Row> group : groups) {
                double sum = 0;
                double min = Double.MAX_VALUE;
                double max = Double.MIN_VALUE;
                Row newRow = sheet.createRow(groups.indexOf(group));
                for (Row row : group) {
                    for (Cell cell : row) {
                        if (criteria.get(cell.getColumnIndex()) == null) {
                            newRow.createCell(cell.getColumnIndex())
                                    .setCellValue(cell.getNumericCellValue());
                        } else if (criteria.get(cell.getColumnIndex()).equals("sum")) {
                            sum += cell.getNumericCellValue();
                            if (group.indexOf(row) == group.size() - 1) {
                                newRow.createCell(cell.getColumnIndex())
                                        .setCellValue(sum);
                            }
                        } else if (criteria.get(cell.getColumnIndex()).equals("min")) {
                            if (min > cell.getNumericCellValue()) min = cell.getNumericCellValue();
                            if (group.indexOf(row) == group.size() - 1) {
                                newRow.createCell(cell.getColumnIndex())
                                        .setCellValue(min);
                            }
                        } else if (criteria.get(cell.getColumnIndex()).equals("max")) {
                            if (max < cell.getNumericCellValue()) max = cell.getNumericCellValue();
                            if (group.indexOf(row) == group.size() - 1) {
                                newRow.createCell(cell.getColumnIndex())
                                        .setCellValue(max);
                            }
                        }
                    }
                }
            }

            List<Row> rowsToDelete = new ArrayList<>();
            for (Row row : sheet) {
                if (row.getRowNum() > groups.size() - 1) {
                    rowsToDelete.add(row);
                }
            }
            for (Row row : rowsToDelete) {
                sheet.removeRow(row);
            }

            System.out.println("Writing data to result file");
            workbook.write(out);
            System.out.println("Done");
        }
    }

    static boolean equals(List<Cell> row1, List<Cell> row2) {
        for (Cell cell : row1) {
            if (cell.getNumericCellValue() != row2.get(cell.getColumnIndex()).getNumericCellValue()) {
                return false;
            }
        }
        return true;
    }
}