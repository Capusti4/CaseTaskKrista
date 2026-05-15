package com.example.CaseTask;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellReference;

public class MustBeStringException extends ExcelException {
    public MustBeStringException(Cell cell) {
        super("Cell " + new CellReference(cell.getRowIndex(), cell.getColumnIndex()).formatAsString() + " must be string");
    }
}
