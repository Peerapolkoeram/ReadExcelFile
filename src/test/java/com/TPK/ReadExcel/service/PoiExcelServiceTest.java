package com.TPK.ReadExcel.service;


import com.TPK.ReadExcel.modal.ColumnExcel;
import com.TPK.ReadExcel.utils.ExcelUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.junit.jupiter.api.io.TempDir;
import org.mockito.InjectMocks;
import org.mockito.Mock;
import org.mockito.Mockito;
import org.mockito.junit.jupiter.MockitoExtension;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotEquals;

@ExtendWith(MockitoExtension.class)
class PoiExcelServiceTest {

    @TempDir
    private Path pathFile;

    private String path;
    private List<Integer> column;
    private Integer rowStart;
    private Integer sheetStart;
    private ColumnExcel excel;

    @InjectMocks
    private PoiExcelService poiExcelService;

    @Mock
    private ExcelUtils excelUtils;

    @BeforeEach
    void mockData() throws IOException {
        path = "TestExcel.xlsx";
        column = List.of(2, 3, 4);
        rowStart = 2;
        sheetStart = 0;
        excel = new ColumnExcel("1", "1", "1");

        // crate excel file
        createFileExcelForTest(path);
    }

    @Test
    void checkMockData() {
        assertEquals("TestExcel.xlsx", path);
        assertEquals(Arrays.asList(2, 3, 4), column);
        assertEquals(2, rowStart);
        assertEquals(0, sheetStart);
    }

    @Test
    void verifyDataWhenReadingExcel() throws IOException {

        String paths = pathFile.resolve(path).toString();

        // Mock class
        Mockito.when(excelUtils.fileInputStream(paths)).thenReturn(new FileInputStream(paths));
        Mockito.when(excelUtils.addListDataToColumnExcel(Mockito.any())).thenReturn(excel);

        // call service
        List<ColumnExcel> resultPoi = poiExcelService.readExcel(paths, column, rowStart, sheetStart);

        // result
        assertEquals("1", resultPoi.get(0).column1());
        assertEquals("1", resultPoi.get(0).column2());
        assertEquals("1", resultPoi.get(0).column3());
    }

    @Test
    void verifyDataWhenReadingExcelDoesNotMatchExpectation() throws IOException {

        String paths = pathFile.resolve(path).toString();

        // Mock class
        Mockito.when(excelUtils.fileInputStream(paths)).thenReturn(new FileInputStream(paths));
        Mockito.when(excelUtils.addListDataToColumnExcel(Mockito.any())).thenReturn(excel);

        // call service
        List<ColumnExcel> resultPoi = poiExcelService.readExcel(paths, column, rowStart, sheetStart);

        // result
        assertNotEquals("2", resultPoi.get(0).column1());
        assertNotEquals("2", resultPoi.get(0).column2());
        assertNotEquals("2", resultPoi.get(0).column3());
    }

    void createFileExcelForTest(String fileName) throws IOException {
        pathFile = Files.createTempDirectory("tempDir");
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Sheet1");
            Row row = sheet.createRow(2);
            column.forEach(e -> {
                Cell cell = row.createCell(e);
                cell.setCellValue(1);
            });
            try (FileOutputStream outputStream = new FileOutputStream(pathFile.resolve(fileName).toFile())) {
                workbook.write(outputStream);
            }
        }
    }

}