package com.example.CaseTask;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellReference;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static com.example.CaseTask.ExcelReader.getCellValue;

class GroupCalculator {
    private final Map<Integer, Double> sums = new HashMap<>();
    private final Map<Integer, Double> mins = new HashMap<>();
    private final Map<Integer, Double> maxs = new HashMap<>();
    private final Map<Integer, List<String>> concats = new HashMap<>();
    private final Map<Integer, Object> criteriaValues = new HashMap<>();

    public void update(Row row, Map<Integer, Object> colTypes) {
        colTypes.forEach((idx, type) -> {
            Cell cell = row.getCell(idx);
            if (cell == null) return;

            String typeStr = type.toString().toLowerCase();
            if (typeStr.equals("-")) return;
            switch (typeStr) {
                case "sum" -> sums.merge(idx, getNumeric(cell), Double::sum);
                case "min" -> mins.merge(idx, getNumeric(cell), Math::min);
                case "max" -> maxs.merge(idx, getNumeric(cell), Math::max);
                case "concat" -> concats.computeIfAbsent(idx, k -> new ArrayList<>())
                        .add(getCellValue(cell).toString());
                default -> criteriaValues.putIfAbsent(idx, getCellValue(cell));
            }
        });
    }

    private double getNumeric(Cell cell) {
        try {
            return cell.getNumericCellValue();
        } catch (Exception e) {
            throw new ExcelException("Cell " + new CellReference(cell.getRowIndex(), cell.getColumnIndex()).formatAsString() + " must be numeric");
        }
    }

    public Object getVal(int idx, String type) {
        return switch (type.toLowerCase()) {
            case "sum" -> sums.getOrDefault(idx, 0.0);
            case "min" -> mins.get(idx);
            case "max" -> maxs.get(idx);
            case "concat" -> String.join(", ", concats.getOrDefault(idx, List.of()));
            default -> criteriaValues.get(idx);
        };
    }
}
