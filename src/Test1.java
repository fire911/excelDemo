import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.DVConstraint;
import org.apache.poi.hssf.usermodel.HSSFDataValidation;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddressList;

public class Test1 {
	public static void main(String[] args) {
//		File file=new File("/Users/lene/Desktop/poi.xls");
		try {
			dropDownList42003("A,B,C,D,E,F,G", "/Users/lene/Desktop/poi1.xls");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	public static void dropDownList42003(String dataSource, String filePath)
		      throws Exception {
		    HSSFWorkbook workbook = new HSSFWorkbook();
		    HSSFSheet realSheet = workbook.createSheet("下拉列表测试");
		    String[] datas = dataSource.split("\\,");
		    DVConstraint dvConstraint = DVConstraint
		        .createExplicitListConstraint(datas);
		    CellRangeAddressList addressList = null;
		    HSSFDataValidation validation = null;
		    for (int i = 0; i < 100; i++) {
		      addressList = new CellRangeAddressList(i, i, 0, 0);
		      validation = new HSSFDataValidation(addressList, dvConstraint);
		      // 03 默认setSuppressDropDownArrow(false)
		      // validation.setSuppressDropDownArrow(false);
		      // validation.setShowErrorBox(true);
		      validation.setShowErrorBox(false);// 取消弹出错误框
		      realSheet.addValidationData(validation);
		    }
		    FileOutputStream stream = new FileOutputStream(filePath);
		    workbook.write(stream);
		    stream.close();
		    addressList = null;
		    validation = null;
		  }
}
