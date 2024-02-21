package com.TPK.ReadExcel.service;

import com.TPK.ReadExcel.modal.ColumnExcel;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@Service
public class JxlsExcelService {

    public List<ColumnExcel> readExcel(String path, List<Integer> column, Integer rowStart, Integer sheetStart) throws BiffException, IOException {
        List<ColumnExcel> resultExcel = new ArrayList<>();

        Workbook workbook = Workbook.getWorkbook(new File(path));
        Sheet sheet = workbook.getSheet(sheetStart);
        int rows = sheet.getRows();
        int columns = sheet.getColumns();
        List<String> rowDataArr = new ArrayList<>();
        for (int i = 0; i < rows; i++) {
            if (i >= rowStart) {
                column.forEach(e -> {
                    for (int j = 0; j < columns; j++) {
                        if (j == e) {
                            rowDataArr.add(sheet.getCell(j,e).getContents());
                        }
                    }
                });
                ColumnExcel excelFile = null;
                try {
                    excelFile = new ColumnExcel(rowDataArr.get(0), rowDataArr.get(1), rowDataArr.get(2));
                } catch (Exception e) {
                    System.out.println("No data");
                }
                resultExcel.add(excelFile);
                rowDataArr.clear();
            }
        }
        return resultExcel;
    }

}
