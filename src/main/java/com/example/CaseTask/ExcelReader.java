package com.example.CaseTask;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelReader {
    private final Workbook workbook;
    private final FileOutputStream out;
    private final List<String> criteriaTypes = List.of(new String[]{"", "-", "min", "max", "sum", "concat"});
    private final Map<Integer, Object> criteria = new HashMap<>();
    private final List<List<Cell>> cRows = new ArrayList<>();
    private final List<List<Row>> groups = new ArrayList<>();
    private final Sheet sheet;

    private double sum;
    private double min;
    private double max;
    private String concat;

    public ExcelReader(String filePath) throws IOException {
        workbook = new XSSFWorkbook(new FileInputStream(filePath));
        out = new FileOutputStream(filePath.split(".xlsx")[0] + "_result.xlsx");
        sheet = workbook.getSheetAt(0);
    }

    public void formatData() throws IOException {
        readCriteria();
        readData();
        createGroups();
        calculateGroups();
        workbook.write(out);
    }

    private void readCriteria() {
        for (Cell cell : sheet.getRow(0)) {
            checkCriteria(cell);
        }
    }

    private void checkCriteria(Cell cell) {
        try {
            if (criteriaTypes.contains(cell.getStringCellValue().toLowerCase())) {
                criteria.put(cell.getColumnIndex(), cell.getStringCellValue().toLowerCase());
            } else {
                throw new ExcelException(
                        "Incorrect criteria in " +
                        new CellReference(cell.getRowIndex(), cell.getColumnIndex()).formatAsString() +
                        ", change file and try again"
                );
            }
        } catch (IllegalStateException e) {
            throw new ExcelException(
                    "Cell " + new CellReference(cell.getRowIndex(), cell.getColumnIndex()).formatAsString() +
                    " must be string"
            );
        }
    }

    private void readData() {
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
    }

    private void createGroups() {
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
    }

    private void calculateGroups() {
        for (List<Row> group : groups) {
            sum = 0;
            min = Double.MAX_VALUE;
            max = Double.MIN_VALUE;
            concat = "";
            Row newRow = sheet.createRow(groups.indexOf(group));
            for (Row row : group) {
                for (Cell cell : row) {
                    checkCell(row, cell, newRow, group);
                }
            }
        }
        cleanOldRows();
    }

    private void checkCell(Row row, Cell cell, Row newRow, List<Row> group) {
        try {
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
            } else if (criteria.get(cell.getColumnIndex()).equals("concat")) {
                concat += cell.getStringCellValue();
                if (group.indexOf(row) == group.size() - 1) {
                    newRow.createCell(cell.getColumnIndex())
                            .setCellValue(concat);
                }
            }
        }
        catch (IllegalStateException e) {
            throw new ExcelException(
                    "Cell " + new CellReference(cell.getRowIndex(), cell.getColumnIndex()).formatAsString() +
                    " must be numeric"
            );
        }
    }

    private void cleanOldRows() {
        List<Row> rowsToDelete = new ArrayList<>();
        for (Row row : sheet) {
            if (row.getRowNum() > groups.size() - 1) {
                rowsToDelete.add(row);
            }
        }
        for (Row row : rowsToDelete) {
            sheet.removeRow(row);
        }
    }

    private boolean equals(List<Cell> row1, List<Cell> row2) {
        if (row1.size() != row2.size()) {
            return false;
        }
        for (Cell cell : row1) {
            Cell secondCell = row2.get(cell.getColumnIndex());
            try {
                if (cell.getNumericCellValue() != secondCell.getNumericCellValue() || cell.getColumnIndex() != secondCell.getColumnIndex()) {
                    return false;
                }
            }
            catch (IllegalStateException e) {
                throw new ExcelException(
                        "Cell " + new CellReference(cell.getRowIndex(), cell.getColumnIndex()).formatAsString() +
                        " must be numeric"
                );
            }
        }
        return true;
    }
}
