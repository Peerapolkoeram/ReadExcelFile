package com.TPK.ReadExcel.service;

import com.TPK.ReadExcel.modal.ColumnExcel;
import jxl.read.biff.BiffException;
import org.dhatim.fastexcel.reader.Cell;
import org.dhatim.fastexcel.reader.ReadableWorkbook;
import org.dhatim.fastexcel.reader.Row;
import org.dhatim.fastexcel.reader.Sheet;
import org.springframework.stereotype.Service;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Stream;

@Service
public class FastExcelService {

    public List<ColumnExcel> readExcel(String path, List<Integer> column, Integer rowStart, Integer sheetStart) throws BiffException, IOException {

        List<ColumnExcel> resultExcel = new ArrayList<>();
        List<String> rowDataArr = new ArrayList<>();

        try (FileInputStream file = new FileInputStream(path); ReadableWorkbook wb = new ReadableWorkbook(file)) {
            Sheet sheet = wb.getFirstSheet();
            try (Stream<Row> rows = sheet.openStream()) {
                rows.forEach(r -> {
                    if (r.getRowNum() >= rowStart) {
                        column.forEach(e -> {
                            for (Cell cell : r) {
                                try {
                                    rowDataArr.add(cell.getRawValue());
                                } catch (Exception ignored) {}
                            }
                        });
                    }
                    ColumnExcel excelFile = null;
                    try {
                        excelFile = new ColumnExcel(rowDataArr.get(0), rowDataArr.get(1), rowDataArr.get(2));
                    } catch (Exception e) {
                        System.out.println("No data");
                    }
                    resultExcel.add(excelFile);
                    rowDataArr.clear();
                });
            }
        }
        return resultExcel;
    }

}
