package com.example.CaseTask;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class ExcelReader {
    private final Workbook resultWorkbook;
    private final FileOutputStream out;
    private final Map<Integer, Object> criteriaIndexes = new HashMap<>();
    private final Sheet sheet;
    private final Map<List<Object>, GroupCalculator> aggregateMap = new LinkedHashMap<>();

    public ExcelReader(String filePath) throws IOException {
        Workbook workbook = new XSSFWorkbook(new FileInputStream(filePath));
        resultWorkbook = new XSSFWorkbook();
        out = new FileOutputStream(filePath.split(".xlsx")[0] + "_result.xlsx");
        sheet = workbook.getSheetAt(0);
    }

    public void formatData() throws IOException {
        readCriteria();
        createGroups();
        writeResult();
    }

    private void readCriteria() {
        for (Cell cell : sheet.getRow(0)) {
            if (getStringValue(cell).isEmpty()) {
                break;
            }
            criteriaIndexes.put(cell.getColumnIndex(), cell.getStringCellValue().toLowerCase());
        }
    }

    private void createGroups() {
        for (Row row : sheet) {
            if (row.getRowNum() == 0 || isRowEmpty(row)) continue;

            List<Object> key = new ArrayList<>();
            for (Integer idx : criteriaIndexes.keySet()) {
                if (!List.of(new String[]{"-", "min", "max", "sum", "concat"}).contains(criteriaIndexes.get(idx).toString())) {
                    key.add(getCellValue(row.getCell(idx)));
                }
            }

            aggregateMap.computeIfAbsent(key, k -> new GroupCalculator())
                    .update(row, criteriaIndexes);
        }
    }

    private void writeResult() throws IOException {
        Sheet resultSheet = resultWorkbook.createSheet("Result");
        int rowIdx = 0;

        for (GroupCalculator stats : aggregateMap.values()) {
            Row newRow = resultSheet.createRow(rowIdx++);
            int colIdx = 0;

            for (int i = 0; i <= Collections.max(criteriaIndexes.keySet()); i++) {
                String type = (String) criteriaIndexes.get(i);

                if (type == null || type.equals("-")) continue;

                Object val = stats.getVal(i, type);
                Cell newCell = newRow.createCell(colIdx++);

                if (val instanceof Double d) newCell.setCellValue(d);
                else if (val != null) newCell.setCellValue(val.toString());
            }
        }

        resultWorkbook.write(out);
    }

    private boolean isRowEmpty(Row row) {
        if (row == null) {
            return true;
        }

        for (int c = row.getFirstCellNum(); c < row.getLastCellNum(); c++) {
            Cell cell = row.getCell(c);

            if (cell != null && cell.getCellType() != CellType.BLANK) {
                return false;
            }
        }

        return true;
    }

    static public Object getCellValue(Cell cell) {
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> cell.getNumericCellValue();
            case BOOLEAN -> cell.getBooleanCellValue();
            default -> "";
        };
    }

    private String getStringValue(Cell cell) {
        try {
            return cell.getStringCellValue();
        } catch (IllegalStateException e) {
            throw new MustBeStringException(cell);
        }
    }
}
