package com.TPK.ReadExcel.utils;

import com.TPK.ReadExcel.modal.ColumnExcel;
import lombok.SneakyThrows;
import org.dhatim.fastexcel.reader.ReadableWorkbook;
import org.springframework.stereotype.Component;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.List;

@Component
public class ExcelUtils {

    @SneakyThrows
    public FileInputStream fileInputStream(String pathFile)  {
        return new FileInputStream(pathFile);
    }

    @SneakyThrows
    public ColumnExcel addListDataToColumnExcel(List<String> data) {
        return new ColumnExcel(data.get(0), data.get(1), data.get(2));
    }

    @SneakyThrows
    public ReadableWorkbook readableWorkbook (FileInputStream file) {
        return new ReadableWorkbook(file);
    }

}