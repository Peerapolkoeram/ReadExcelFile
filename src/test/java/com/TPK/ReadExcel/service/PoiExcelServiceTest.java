package com.TPK.ReadExcel.service;


import com.TPK.ReadExcel.modal.ColumnExcel;
import com.TPK.ReadExcel.utils.ExcelUtils;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Order;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.InjectMocks;
import org.mockito.Mock;
import org.mockito.Mockito;
import org.mockito.junit.jupiter.MockitoExtension;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;

@ExtendWith(MockitoExtension.class)
class PoiExcelServiceTest {

    private String path;
    private List<Integer> column;
    private Integer rowStart;
    private Integer sheetStart;
    private  ColumnExcel excel;

    @InjectMocks
    PoiExcelService poiExcelService;

    @Mock
    ExcelUtils excelUtils;

    @BeforeEach
    void mockData() {
        path = "FileTest/TestExcel.xlsx";
        column = Arrays.asList(2,3,4);
        rowStart = 2;
        sheetStart = 0;
        excel = new ColumnExcel("1","2","3");
    }

    @Test
    @Order(1)
    void checkMockData() {
        assertEquals("FileTest/TestExcel.xlsx", path);
        assertEquals(Arrays.asList(2, 3, 4), column);
        assertEquals(2, rowStart);
        assertEquals(0, sheetStart);
    }

    @Test
    @Order(2)
    void poiExcel() throws IOException {

        // Mock class
        Mockito.when(excelUtils.fileInputStream(path)).thenReturn(new FileInputStream(path));
        Mockito.when(excelUtils.addListDataToColumnExcel(Mockito.any())).thenReturn(excel);

        // call service
        List<ColumnExcel> resultPoi = poiExcelService.readExcel(path, column,rowStart, sheetStart);

        // result
        assertEquals("1", resultPoi.get(0).column1());
        assertEquals("2", resultPoi.get(0).column2());
        assertEquals("3", resultPoi.get(0).column3());
    }

}