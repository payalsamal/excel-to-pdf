package com.excel.conversion.controller;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import com.excel.conversion.util.ExcelToPdfUtil;

@RestController
@RequestMapping("v1/conversion")
public class ExcelReadController {
	
	
	@Autowired
	private ExcelToPdfUtil excelUtil;
	
	@PostMapping("/excel-to-pdf")
	public ResponseEntity<?> convertExcelToPdf(@RequestParam MultipartFile file) throws Exception {
		
		return excelUtil.writeExcelDataInPdf(file, null);
	}
	

}
