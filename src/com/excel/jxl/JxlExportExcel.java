package com.excel.jxl;

import java.io.File;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

public class JxlExportExcel {
	
	private static String[] titels={"编号","姓名","性别"};

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		File file=new File("/Users/lene/Desktop/jxl.xls");
		JxlExportExcel jee=new JxlExportExcel();
		try {
			if(file.exists()){
				file.delete();
			}
			file.createNewFile();
			jee.createExcel(file);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	/**
	 * 创建EXCEL
	 * @param file
	 * @throws Exception 
	 */
	public void createExcel(File file) throws Exception{
		//创建工作簿
		WritableWorkbook workbook=Workbook.createWorkbook(file);
		//创建sheet页
		WritableSheet sheet= workbook.createSheet("sheet",0);
		//创建label
		Label label=null;
		//设置列名
		for (int i = 0; i < titels.length; i++) {
			label=new Label(i, 0, titels[i]);
			sheet.addCell(label);
		}
		//添加数据
		for (int i = 1; i < 10; i++) {
			label=new Label(0, i,i+"");
			sheet.addCell(label);
			label=new Label(1, i, "张三"+i);
			sheet.addCell(label);
			label=new Label(2, i,"男");
			sheet.addCell(label);
		}
		workbook.write();
		workbook.close();
	}
}
