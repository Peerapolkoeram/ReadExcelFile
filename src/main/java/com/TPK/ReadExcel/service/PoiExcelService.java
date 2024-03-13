package com.TPK.ReadExcel.service;

import com.TPK.ReadExcel.modal.ColumnExcel;
import com.TPK.ReadExcel.utils.ExcelUtils;
import lombok.RequiredArgsConstructor;
import lombok.SneakyThrows;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@Service
@RequiredArgsConstructor
public class PoiExcelService {

    private final ExcelUtils excelUtils;

    @SneakyThrows
    public List<ColumnExcel> readExcel(String path, List<Integer> column, Integer rowStart, Integer sheetStart) {
        List<ColumnExcel> resultData = new ArrayList<>();
        // search filename
        FileInputStream file = excelUtils.fileInputStream(path);
        // read file
        List<String> rowDataArr = new ArrayList<>();
        Workbook workbook = new XSSFWorkbook(file);
        Sheet sheet = workbook.getSheetAt(sheetStart);
        for (Row rows : sheet) {
            if (rows.getRowNum() >= rowStart) {
                for (Cell cell : rows) {
                    /* loop column arr */
                    for (int j : column) {
                        if (cell == (cell.getRow().getCell(j))) {
                            switch (cell.getCellType()) {
                                case STRING -> rowDataArr.add(cell.getRichStringCellValue().getString());
                                case NUMERIC -> rowDataArr.add(String.valueOf((int) cell.getNumericCellValue()));
                                case BOOLEAN -> rowDataArr.add(String.valueOf(cell.getBooleanCellValue()));
                                default -> rowDataArr.add(null);
                            }
                        }
                    }
                }
                ColumnExcel excelFile = excelUtils.addListDataToColumnExcel(rowDataArr);
                resultData.add(excelFile);
                /* Clear Data in arr */
                rowDataArr.clear();
            }
        }
        return resultData;
    }
}
