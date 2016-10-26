package com.excel.poi;

import java.io.File;
import java.io.FileOutputStream;
import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class PoiExportExcel2007 {
	
	private static String[] titels = { "编号", "姓名", "性别" };

	public static void main(String[] args) {
		File file=new File("/Users/lene/Desktop/poi.xlsx");
		PoiExportExcel2007 pee=new PoiExportExcel2007();
		try {
			if(file.exists()){
				file.delete();
			}
			file.createNewFile();
			pee.createExcel(file);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void createExcel(File file) throws Exception {
		// 创建工作簿
		XSSFWorkbook workbook = new XSSFWorkbook();
		// 创建sheet页
		Sheet sheet = workbook.createSheet();
		// 创建第一行
		Row row = sheet.createRow(0);
		// 创建单元格
		Cell cell = null;
		//设置列名
		for (int i = 0; i < titels.length; i++) {
			cell=row.createCell(i);
			cell.setCellValue(titels[i]);
		}
		//追加数据
		for (int i = 1; i < 10; i++) {
			row=sheet.createRow(i);
			cell=row.createCell(0);
			cell.setCellValue(i);
			cell=row.createCell(1);
			cell.setCellValue("张三"+i);
			cell=row.createCell(2);
			cell.setCellValue("男");
		}
		
		//将数据写入文件
		FileOutputStream fos=FileUtils.openOutputStream(file);
		workbook.write(fos);
		//关闭流
		fos.close();
		workbook.close();
	}

}
