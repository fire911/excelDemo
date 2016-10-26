package com.excel.jxl;

import java.io.File;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;

public class JxlReadExcel {

	public static void main(String[] args) {
		JxlReadExcel jre=new JxlReadExcel();
		File file=new File("/Users/lene/Desktop/jxl.xls");
		try {
			jre.readExcel(file);
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	public void readExcel(File file) throws Exception{
		//创建Workbook
		Workbook workbook=Workbook.getWorkbook(file);
		//获取第一个sheet页
		Sheet sheet=workbook.getSheet(0);
		for (int i = 0; i < sheet.getRows(); i++) {
			for (int j = 0; j < sheet.getColumns(); j++) {
				Cell cell=sheet.getCell(j, i);
				System.out.print(cell.getContents()+"  ");
			}
			System.out.println();
		}
		workbook.close();
	}
}
