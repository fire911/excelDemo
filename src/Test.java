import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

import org.apache.poi.hssf.usermodel.DVConstraint;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDataValidation;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Test {
	//主方法

	public static String writeExcel(String str,List mList,String path) throws IOException {
	Date dt = new Date();
	SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmmss");
	String temp_str = "";
	temp_str = sdf.format(dt);
	//System.out.println(InExcel.class.getClassLoader().getResource("/").getPath());
	// String path = request.getSession().getServletContext().getRealPath("/");
	//System.out.println(path+"00000000000000");
	String pathname=path+"upload\\"+str+temp_str+".xls";
	String pathname1="upload\\"+str+temp_str+".xls";
	File file=new File(pathname);
	String fileName = file.getName();
	String extension = fileName.lastIndexOf(".") == -1 ? "" : fileName
	.substring(fileName.lastIndexOf(".") + 1);

	//判断文件类型

	if ("xls".equals(extension)) {
	WriteIn(file,"xls");
	String[][] data=check(str);
	write2003Excel(file.getPath(),mList,data,str);
	return pathname1;

	} else if ("xlsx".equals(extension)) {
	String[][] data=check(str);
	// write2007Excel(file.getPath(),mList,data);

	} else {

	throw new IOException("不支持的文件类型");

	}
	return "error";

	}
	
	//写入excel方法
    public static String WriteIn(File file,String extension) throws IOException{
        //2003xls文件创建模式
          FileOutputStream fileOut = new FileOutputStream(file.getAbsolutePath());
        if ("xls".equals(extension)) {
           HSSFWorkbook wb = new HSSFWorkbook();
           HSSFSheet sheet1 = wb.createSheet("Sheet1");
           HSSFSheet sheet2 = wb.createSheet("Sheet2");
           HSSFSheet sheet3 = wb.createSheet("Sheet3");
           wb.write(fileOut);
        }else if ("xlsx".equals(extension)){
            //2007模式的写法
            XSSFWorkbook wb = new XSSFWorkbook();
               XSSFSheet sheet1 = wb.createSheet("Sheet1");
               XSSFSheet sheet2 = wb.createSheet("Sheet2");
               XSSFSheet sheet3 = wb.createSheet("Sheet3");
               wb.write(fileOut);
       }
           
           fileOut.close();
        return "success";
    }
    
    /**
    * 协查有效性和标题筛选
    * @param str
    * @return
    */
    public static String[][] check(String str){
    if("00201".equals(str)){
    //标题 ： 姓名 性别 身份证号码 死亡时间
    String[] bzbt={"姓名","性别 ","身份证号码","救助业务名称","救助证号","救助金额","救助开始时间 ","救助结束时间"};
    //有效性 ： 0-未知的性别,1-男性,2-女性,9-未说明的性别
    String[] bzdata2={ "0-未知的性别","1-男性","2-女性","9-未说明的性别"};
    //01-救灾救济,02-城市社会救助,03-双拥、优抚、安置,04-社会事务与社会福利,05-慈善事业
    String[] bzdata3={ "01-救灾救济","02-城市社会救助","03-双拥、优抚、安置","04-社会事务与社会福利","05-慈善事业"};
    //行有效性
    String [] bzdatarow={"1","3"};
    String[][] a ={bzbt,bzdata2,bzdata3,bzdatarow};
    return a;
    }

    return null;

    }
    
  //下拉列表元素很多的情况

    private static HSSFDataValidation SetDataValidation(String strFormula,int firstRow,int firstCol,int endRow,int endCol)
    {
    CellRangeAddressList regions = new CellRangeAddressList(firstRow, endRow, firstCol, endCol);
    // CellRangeAddressList regions = new CellRangeAddressList( firstRow,
    // (short) 300, (short) 1, (short) 1);//add 新顺序为 起始行 终止行 起始列 终止列
    DVConstraint constraint = DVConstraint.createFormulaListConstraint(strFormula);//add
    HSSFDataValidation dataValidation = new HSSFDataValidation(regions,constraint);//add

    dataValidation.createErrorBox("Error", "Error");
    dataValidation.createPromptBox("", null);

    return dataValidation;
    }
    //255以内的下拉
    public static DataValidation setDataValidation(Sheet sheet,String[] textList, int firstRow, int endRow, int firstCol, int endCol) {

    DataValidationHelper helper = sheet.getDataValidationHelper();
    // 加载下拉列表内容
    DataValidationConstraint constraint = helper.createExplicitListConstraint(textList);
    // DVConstraint constraint = new DVConstraint();
    constraint.setExplicitListValues(textList);

    // 设置数据有效性加载在哪个单元格上。
    // 四个参数分别是：起始行、终止行、起始列、终止列
    CellRangeAddressList regions = new CellRangeAddressList((short) firstRow, (short) endRow, (short) firstCol, (short) endCol);

    // 数据有效性对象
    DataValidation data_validation = helper.createValidation(constraint, regions);
    //DataValidation data_validation = new DataValidation(regions, constraint);

    return data_validation;
    }
    
    
  //write 2003Excel
    public static void write2003Excel(String filePath,List list,String[][] data,String str) {

       try {
           if(list.size()<=60000){
          //创建excel文件对象   

          HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream(filePath));
           //创建一个张表   
          HSSFSheet sheet;
          //创建行对象
          HSSFRow row = null;

          //创建表格对象

          HSSFCell cell = null;
          
       
                      sheet = wb.getSheetAt(0);
                     
                    //   sheet.addValidationData(setDataValidation(sheet,data[1], 1,list.size(), 1, 1));
                       String[] num=data[data.length-1];
                       System.out.println(num.length);
                       if (data.length>1) {
                            for(int i=0;i<num.length;i++){
                                   int rowdata=Integer.parseInt(num[i]);
                                   sheet.addValidationData(setDataValidation(sheet,data[i+1], 1,list.size(),rowdata , rowdata));
                                   System.out.println("我出现"+rowdata);
                               }
                       }
                       if("00301".equals(str)){
                            String strFormula = "Sheet2!$A$2:$A$59" ;
                            sheet.addValidationData(SetDataValidation(strFormula, 1, 5, list.size(), 5));
                        }
                      
                  row=sheet.createRow(0);
                  for (int i = 0; i < data[0].length; i++) {
                      cell=row.createCell(i);
                      
                      cell.setCellValue(data[0][i]);
               }
                  
                  //循环行
                  for (int i=1; i <=list.size(); i++) {
                     row = sheet.createRow(i);
                     //循环列
                     
                     for (int j=0; j<data[0].length; j++) {
                        cell = row.createCell(j);//创建单元格
                        String m=String.valueOf(list.get(i-1)).replace("[", "").replace("]", "");
                       String[] a=m.split(",");
                        cell.setCellValue(a[j]);//赋值
                      //  cell.setCellFormula("Sheet2!$A$2:$A$59");

                     }
                 }
                  //select Ahd010401,Ahd010405,Ahd010404,Ahd010409,Ahd010410,Ahd010406,Ahd010408,Ahd010416 from res_00301
                  if("00301".equals(str)){
                      String[] bzdata4={"民族","01-汉族","02-蒙古族","03-回族","04-藏族","05-维吾尔族","06-苗族","07-彝族","08-壮族","09-布依族","10-朝鲜族","11-满族","12-侗族","13-瑶族","14-白族","15-土家族","16-哈尼族","17-哈萨克族","18-傣族","19-黎族","20-傈僳族","21-佤族","22-畲族","23-高山族","24-拉祜族","25-水族","26-东乡族","27-纳西族","28-景颇族","29-柯尔克孜族","30-土族","31-达斡尔族","32-仫佬族","33-羌族","34-布朗族","35-撒拉族","36-毛难族","37-仡佬族","38-锡伯族","39-阿昌族","40-普米族","41-塔吉克族","42-怒族","43-乌孜别克族","44-俄罗斯族","45-鄂温克族","46-德昂族","47-保安族","48-裕固族","49-京族","50-塔塔尔族","51-独龙族","52-鄂伦春族","53-赫哲族","54-门巴族","55-珞巴族","56-基诺族","98-外国血统","99-其他"};
                      sheet=wb.getSheetAt(1);
                      for (int i = 0; i < bzdata4.length; i++) {
                       row=sheet.createRow(i);
                       cell=row.createCell(0);
                       cell.setCellValue(bzdata4[i]);
                   }
                  }  
      
          FileOutputStream out = new FileOutputStream(filePath);

          wb.write(out);

          out.close();
            }else {
                throw new IOException("超出excel的可写范围，可写为60000行");
            }

       } catch (Exception e) {

          e.printStackTrace();

       }

    }
}
