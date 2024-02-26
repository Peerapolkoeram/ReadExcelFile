package com.TPK.ReadExcel;

import com.TPK.ReadExcel.modal.ColumnExcel;
import com.TPK.ReadExcel.service.FastExcelService;
import com.TPK.ReadExcel.service.JxlsExcelService;
import com.TPK.ReadExcel.service.PoiExcelService;
import lombok.RequiredArgsConstructor;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.util.ArrayList;
import java.util.List;

@RequiredArgsConstructor
@SpringBootApplication
public class ReadExcelApplication implements CommandLineRunner {

	private final PoiExcelService poiExcelService;
	private final JxlsExcelService jxlsExcelService;
	private final FastExcelService fastExcelService;

	public static void main(String[] args) {
		SpringApplication.run(ReadExcelApplication.class, args);
	}

	@Override
	public void run(String... args) throws Exception {

		String path = "FileTest/TestExcel.xlsx";
		List<Integer> column = new ArrayList<>();
		column.add(2);
		column.add(3);
		column.add(4);
		Integer rowStart = 2;
		Integer sheetStart = 0;

		// POI
		List<ColumnExcel> resultPoi = poiExcelService.readExcel(path, column,rowStart, sheetStart);
		resultPoi.forEach(System.out::println);

		// Jxl
		List<ColumnExcel> resultJxl = jxlsExcelService.readExcel(path, column, rowStart, sheetStart);
		resultJxl.forEach(System.out::println);

		// Fast
		List<ColumnExcel> resultFast = fastExcelService.readExcel(path, column, rowStart, sheetStart);
		resultFast.forEach(System.out::println);
	}

}