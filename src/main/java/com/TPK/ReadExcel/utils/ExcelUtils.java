package com.TPK.ReadExcel.utils;

import com.TPK.ReadExcel.modal.ColumnExcel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.List;

public class ExcelUtils {

    public FileInputStream fileInputStream(String pathFile) {
        FileInputStream fileInputStream;
        try {
            fileInputStream = new FileInputStream(pathFile);
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        }
        return fileInputStream;
    }

    public ColumnExcel addListDataToColumnExcel(List<String> data) {
        ColumnExcel excelFile = null;
        try {
            excelFile = new ColumnExcel(data.get(0), data.get(1), data.get(2));
        } catch (Exception e) {
            System.out.println("No data");
        }
        return excelFile;
    }

}
