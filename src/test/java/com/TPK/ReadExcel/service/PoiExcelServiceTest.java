package com.TPK.ReadExcel.service;


import com.TPK.ReadExcel.modal.ColumnExcel;
import com.TPK.ReadExcel.utils.ExcelUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Tag;
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
import java.util.function.Function;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotEquals;

@ExtendWith(MockitoExtension.class)
class PoiExcelServiceTest {

    @TempDir
    private static Path pathFile;
    private static String path;
    private static List<Integer> column;
    private static Integer rowStart;
    private static Integer sheetStart;
    private static ColumnExcel excel;

    @InjectMocks
    private PoiExcelService poiExcelService;

    @Mock
    private ExcelUtils excelUtils;

    @BeforeAll
    static void mockData() throws IOException {
        path = "TestExcel.xlsx";
        column = List.of(2, 3, 4);
        rowStart = 2;
        sheetStart = 0;
        excel = new ColumnExcel("1", "1", "1");

        // crate excel file
        createFileExcelForTest(path);
    }

    static void createFileExcelForTest(String fileName) throws IOException {
        pathFile = Files.createTempDirectory("tempDir");
        try (Workbook workbook = new XSSFWorkbook()) {
            // create sheet[0] => name: Sheet1
            Sheet sheet = workbook.createSheet("Sheet1");
            // start rows => 2
            Row row = sheet.createRow(2);
            // create column => 2,3,4
            column.forEach(e -> {
                Cell cell = row.createCell(e);
                // set value = 1
                cell.setCellValue(1);
            });
            try (FileOutputStream outputStream = new FileOutputStream(pathFile.resolve(fileName).toFile())) {
                workbook.write(outputStream);
            }
        }
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

    @Test
    void TestFunction() {
        Function<String, Integer> strToInt = Integer::parseInt;
        int result = strToInt.apply("222");
        System.out.println(result);
    }

    @Test
    void TestFunctionComposeAndThen() {
        Function<Integer, Integer> nultiple2 = i ->  i * 2;
        Function<Integer, Integer> plus2 = i ->  i + 2;
        int result1 = plus2.compose(nultiple2).apply(8);
        result1 = plus2.andThen(nultiple2).apply(8);
        System.out.println(result1);
    }

    @Test
    void TestStreamForEach() {
        List<String> arr = List.of("Test","Test2","Test3","Mount","And");
        // loop all
        arr.stream().forEach(System.out::println);
        System.out.println("--------");
        // filter
        List<String> arr2 = arr.stream().filter( a -> a.contains("Test")).toList();
        arr2.stream().forEach(System.out::println);
        System.out.println("--------");
        // sorted
        List<String> arr3 = arr.stream().sorted().toList();
        arr3.stream().forEach(System.out::println);
        System.out.println("--------");
        // anyMatch
        System.out.println(arr.stream().anyMatch(a -> a.startsWith("G")));
    }

    @Test
    void TestParallelStream() {
        List<String> arr = List.of("Test","Test2","Test3","Mount","And");
        arr.parallelStream().forEach(System.out::println);
    }

    @Test
    void TestShortCircuit() {
        System.out.println((10 < 100) ? "True":"False");
    }

}