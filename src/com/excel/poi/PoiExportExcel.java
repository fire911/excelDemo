package com.excel.poi;

import java.io.File;
import java.io.FileOutputStream;
import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.DVConstraint;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDataValidation;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.ComparisonOperator;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationConstraint.OperatorType;
import org.apache.poi.ss.usermodel.DataValidationConstraint.ValidationType;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;


public class PoiExportExcel {
	
	private static String[] titels = { "Location", "Assessment date", "AssessmentStatrTime","AssessmentEndTime","Capacity","RegistrationStartTime","RegistrationEndTime"};

	public static void main(String[] args) {
		File file=new File("/Users/lene/Desktop/poi.xls");
		PoiExportExcel pee=new PoiExportExcel();
		try {
			if(file.exists()){
				file.delete();
			}
			file.createNewFile();
			String s="A,B,C,D,E,F,G";
			pee.createExcel(file,s.split(","));
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void createExcel(File file,String[] list) throws Exception {
		// 创建工作簿
		HSSFWorkbook workbook = new HSSFWorkbook();
		// 创建sheet页
		HSSFSheet sheet = workbook.createSheet();
		// 创建第一行
		HSSFRow row = sheet.createRow(0);
		// 创建单元格
		HSSFCell cell = null;
		//设置列名
		for (int i = 0; i < titels.length; i++) {
			cell=row.createCell(i);
			cell.setCellValue(titels[i]);
		}
		DVConstraint dvConstraint = DVConstraint
		        .createExplicitListConstraint(list);
		CellRangeAddressList addressList = null;
	    HSSFDataValidation validation = null;
	    addressList = new CellRangeAddressList(1, 1000, 0, 0);
	    validation = new HSSFDataValidation(addressList, dvConstraint);
	    validation.setShowErrorBox(true);// 取消弹出错误框
	    sheet.addValidationData(validation);
	    //构造constraint对象  
//        DVConstraint constraint=DVConstraint.createCustomFormulaConstraint("B1");  
        //四个参数分别是：起始行、终止行、起始列、终止列  
//        CellRangeAddressList regions=new CellRangeAddressList(1,1000,0,0);  
        //数据有效性对象  
//        HSSFDataValidation data_validation_view = new HSSFDataValidation(regions, constraint); 
//        sheet.addValidationData(data_validation_view);
        //添加日期校验
        HSSFDataValidation validationDate=getDataValidationByDate(1, 1000, 1, 1);
        sheet.addValidationData(validationDate);
        HSSFDataValidation validationStartTime= getDataValidationByTime(1, 1000, 2, 3);
        sheet.addValidationData(validationStartTime);
        HSSFDataValidation validationNumber=getDataValidationByNumber(1, 1000, 4, 4);
        sheet.addValidationData(validationNumber);
        CellRangeAddress c = CellRangeAddress.valueOf("A1:G1");
		sheet.setAutoFilter(c);
		//追加数据
//		for (int i = 1; i < 10; i++) {
//			row=sheet.createRow(i);
//			cell=row.createCell(0);
//			cell.setCellValue(i);
//			cell=row.createCell(1);
//			addressList = new CellRangeAddressList(i, i, 0, 0);
//		    validation = new HSSFDataValidation(addressList, dvConstraint);
//			cell.setCellValue("张三"+i);
//			cell=row.createCell(2);
//			cell.setCellValue("男");
//			
//		}
		
		//将数据写入文件
		FileOutputStream fos=FileUtils.openOutputStream(file);
		workbook.write(fos);
		//关闭流
		fos.close();
		workbook.close();
	}
	/**
	 * 设置导出的excel某区域单元格为下拉选择方式
	 * @param selectList 选择的select数据列表
	 * @param firstRow 起始行
	 * @param lastRow 终止行
	 * @param firstCol 起始列
	 * @param lastCol 终止列
	 * @return
	 */
	public static HSSFDataValidation validationData(String[] selectList,int firstRow,int lastRow,int firstCol,int lastCol){
		DVConstraint dvConstraint = DVConstraint.createExplicitListConstraint(selectList);
		CellRangeAddressList addressList = new CellRangeAddressList(firstRow, lastRow, firstCol, lastCol);
		return new HSSFDataValidation(addressList, dvConstraint);
	}
	
	/**
	 * 校验导出的excel某区域单元格为下拉选择方式所输入数据的有效性
	 * @param selectList 选择的select数据列表
	 * @param firstRow 起始行
	 * @param lastRow 终止行
	 * @param firstCol 起始列
	 * @param lastCol 终止列
	 * @return
	 */
	public static HSSFDataValidation validationData(int firstRow,int lastRow,int firstCol,int lastCol){
		DVConstraint constraint=DVConstraint.createCustomFormulaConstraint("B1");  
		CellRangeAddressList regions = new CellRangeAddressList(firstRow, lastRow, firstCol, lastCol);
		return new HSSFDataValidation(regions, constraint);
	}
	
	/**
	 * 校验导出的excel某区域单元格为日期方式所输入数据的有效性
	 * @param selectList 选择的select数据列表
	 * @param firstRow 起始行
	 * @param lastRow 终止行
	 * @param firstCol 起始列
	 * @param lastCol 终止列
	 * @return
	 */
	public static HSSFDataValidation getDataValidationByDate(int firstRow,int lastRow,int firstCol,int lastCol){
		DVConstraint constraint=DVConstraint.createDateConstraint(DataValidationConstraint.OperatorType.BETWEEN,"1970-01-01","2999-12-31","yyyy-MM-dd"); 
		CellRangeAddressList regions = new CellRangeAddressList(firstRow, lastRow, firstCol, lastCol);
		return new HSSFDataValidation(regions, constraint);
	}
	
	/**
	 * 校验导出的excel某区域单元格为时间方式所输入数据的有效性
	 * @param selectList 选择的select数据列表
	 * @param firstRow 起始行
	 * @param lastRow 终止行
	 * @param firstCol 起始列
	 * @param lastCol 终止列
	 * @return
	 */
	public static HSSFDataValidation getDataValidationByTime(int firstRow,int lastRow,int firstCol,int lastCol){
		DVConstraint constraint=DVConstraint.createTimeConstraint(OperatorType.BETWEEN, "00:00", "23:59"); 
		CellRangeAddressList regions = new CellRangeAddressList(firstRow, lastRow, firstCol, lastCol);
		HSSFDataValidation dataValidation=new HSSFDataValidation(regions, constraint);
		dataValidation.createErrorBox("", "要输入的值必须是一个时间介于00:00和23:59");
		return dataValidation;
	}
	
	/**
	 * 校验导出的excel某区域单元格为数字类型所输入数据的有效性
	 * @param selectList 选择的select数据列表
	 * @param firstRow 起始行
	 * @param lastRow 终止行
	 * @param firstCol 起始列
	 * @param lastCol 终止列
	 * @return
	 */
	public static HSSFDataValidation getDataValidationByNumber(int firstRow,int lastRow,int firstCol,int lastCol){
		DVConstraint constraint=DVConstraint.createNumericConstraint(ValidationType.INTEGER,OperatorType.GREATER_OR_EQUAL,"0", null); 
		CellRangeAddressList regions = new CellRangeAddressList(firstRow, lastRow, firstCol, lastCol);
		HSSFDataValidation dataValidation=new HSSFDataValidation(regions, constraint);
		return dataValidation;
	}
}
