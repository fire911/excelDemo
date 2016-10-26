package com.excel.poi;

import java.io.File;
import java.io.InputStream;

import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class PoiReadExcel2007 {

	public static void main(String[] args) {
		File file=new File("/Users/lene/Desktop/poi.xlsx");
		PoiReadExcel2007 pre=new PoiReadExcel2007();
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
		XSSFWorkbook workbook=new XSSFWorkbook(inputStream);
		//获取sheet页 
		//HSSFSheet sheet=workbook.getSheet(String sheetName);//通过sheet页名称获取
		Sheet sheet=workbook.getSheetAt(0);//通过索引获取
		int lastRowIndex=sheet.getLastRowNum();//获取最后一行的索引
		//循环每一行读取数据 
		for (int i = 0; i <=lastRowIndex; i++) {
			Row row= sheet.getRow(i);
			int lastColumnIndex=row.getLastCellNum();//获取每一个行的最后一个单元格索引
			for (int j = 0; j < lastColumnIndex; j++) {
				Cell cell= row.getCell(j);
				if(cell.getCellType()==HSSFCell.CELL_TYPE_NUMERIC){//数值类型
					if(HSSFDateUtil.isCellDateFormatted(cell)){
						//  如果是date类型则 ，获取该cell的date值     
				        //value = HSSFDateUtil.getJavaDate(cell.getNumericCellValue()).toString();  
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
