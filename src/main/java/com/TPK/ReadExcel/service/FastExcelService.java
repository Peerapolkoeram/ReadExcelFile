package com.TPK.ReadExcel.service;

import com.TPK.ReadExcel.modal.ColumnExcel;
import com.TPK.ReadExcel.utils.ExcelUtils;
import lombok.RequiredArgsConstructor;
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
@RequiredArgsConstructor
public class FastExcelService {

    private final ExcelUtils excelUtils;

    public List<ColumnExcel> readExcel(String path, List<Integer> column, Integer rowStart, Integer sheetStart) throws IOException {

        List<ColumnExcel> resultExcel = new ArrayList<>();
        List<String> rowDataArr = new ArrayList<>();
        FileInputStream file = excelUtils.fileInputStream(path);
        ReadableWorkbook wb = excelUtils.readableWorkbook(file);
        Sheet sheet = wb.getFirstSheet();
        Stream<Row> rows = sheet.openStream();
        rows.forEach(r -> {
            if (r.getRowNum() >= rowStart) {
                column.forEach(e -> {
                    for (Cell cell : r) {
                        try {
                            rowDataArr.add(cell.getRawValue());
                        } catch (Exception ignored) {
                        }
                    }
                });
            }
            ColumnExcel excelFile = excelUtils.addListDataToColumnExcel(rowDataArr);
            resultExcel.add(excelFile);
            rowDataArr.clear();
        });
        return resultExcel;
    }

}
