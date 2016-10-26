package com.excel.poi;

import java.io.File;
import java.io.InputStream;

import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class PoiReadExcel {

	public static void main(String[] args) {
		File file=new File("/Users/lene/Desktop/assessment_template-1.xls");
		PoiReadExcel pre=new PoiReadExcel();
		try {
			pre.readExcel(file);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	/**
	 * excel 2003
	 * @param file
	 * @throws Exception
	 */
	public void readExcel(File file) throws Exception{
		InputStream inputStream= FileUtils.openInputStream(file);
		//获取需要读取的excel
		HSSFWorkbook workbook=new HSSFWorkbook(inputStream);
		//获取sheet页 
		//HSSFSheet sheet=workbook.getSheet(String sheetName);//通过sheet页名称获取
		HSSFSheet sheet=workbook.getSheetAt(0);//通过索引获取
		int lastRowIndex=sheet.getLastRowNum();//获取最后一行的索引
		//循环每一行读取数据 
		for (int i = 0; i <=lastRowIndex; i++) {
			HSSFRow row= sheet.getRow(i);
			int lastColumnIndex=row.getLastCellNum();//获取每一个行的最后一个单元格索引
			for (int j = 0; j < lastColumnIndex; j++) {
				HSSFCell cell= row.getCell(j);
				if(cell.getCellType()==HSSFCell.CELL_TYPE_NUMERIC){//数值类型
					if(HSSFDateUtil.isCellDateFormatted(cell)){
						//  如果是date类型则 ，获取该cell的date值     
				        System.out.print(HSSFDateUtil.getJavaDate(cell.getNumericCellValue()).toString()+" ");
					}else{
						System.out.print(String.valueOf((int)cell.getNumericCellValue())+"  ");
					}
				}else{
					System.out.print(cell.getStringCellValue()+"   ");
				}
			}
			System.out.println();
		}
		inputStream.close();
		workbook.close();
	}
	
	
}
