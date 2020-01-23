package com.mmelo.excel;

import com.mmelo.excel.processor.GenerateExcelWithImageCell;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

import java.time.LocalDate;

@SpringBootTest
class ExcelApplicationTests {

	@Autowired
	private GenerateExcelWithImageCell generateExcelWithImageCell;


	@Test
	void contextLoads() throws Exception {
		generateExcelWithImageCell.create(LocalDate.now(), LocalDate.now());
	}

}
