package cn.linkey.rulelib.RYDY;

import java.util.*;
import cn.linkey.dao.*;
import cn.linkey.util.*;
import cn.linkey.doc.*;
import cn.linkey.factory.*;
import cn.linkey.wf.*;
import cn.linkey.rule.LinkeyRule;
import cn.linkey.org.LinkeyUser;

//输出流
import java.io.FileOutputStream;
import java.io.OutputStream;

//Excel相关
import org.apache.poi.hssf.usermodel.HSSFCell;  
import org.apache.poi.hssf.usermodel.HSSFCellStyle;  
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;  
import org.apache.poi.hssf.usermodel.HSSFSheet;  
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
/**
 * @RuleName:导出Excel
 * @author  admin
 * @version: 8.0
 * @Created: 2016-09-05 10:15
 */
final public class R_RYDY_B004 implements LinkeyRule {
	@Override
	public String run(HashMap<String, Object> params) throws Exception  {
	    //params为运行本规则时所传入的参数
	    String result="系统提示：Excel文件导出成功！";
		String sql = "select * from rydy_dyjlb where wf_addname <> 'admin' and wf_orunid ";
        sql+="not in('7aea1f8b06b7f041510b1d9034097ef3d967','bf36e9430eecc042e60b4550bb6e7edf6d6d','cfa3e10c045bf04a0c0a9680ddbbe7fb5d3e','a521aa2000e7b044580a19706766faa9f87e','95604fed00af104dbb08b870ab75b8058d41',";
        sql+="'8a6b2a2c0dd70046270a9f30a5d2b7d1d09e','848a24800aa1704736082b304cf7342e4445','2bee9fdc0fa8e0434909902041a4b3377f6d','69b639210c7ed0490a0b2760f67fe3ef3385','cfa3e10c045bf04a0c0a9680ddbbe7fb5d3e','123d537d0b8fc04bc4096540d454a8f3403e')";
        sql+="   order by WF_DOCCREATED desc  ";
// 		if(!"".equals(name)&&name!=null){
// 	        sql +=" and sqr like '%"+name+"%'";
// 	    }
        
        //导出Excel
        try {
            // 第一步，创建一个webbook，对应一个Excel文件  
            HSSFWorkbook wb = new HSSFWorkbook();
            // 第二步，在webbook中添加一个sheet,对应Excel文件中的sheet  
            HSSFSheet sheet = wb.createSheet("sheet1");
            sheet.setColumnWidth((short) 0,4000);
            sheet.setColumnWidth((short) 1,4000);
            sheet.setColumnWidth((short) 2,4000);
            sheet.setColumnWidth((short) 3,8000);
            sheet.setColumnWidth((short) 4,10000);
            sheet.setColumnWidth((short) 5,10000);
            sheet.setColumnWidth((short) 6,10000);
            sheet.setColumnWidth((short) 7,10000);
             sheet.setColumnWidth((short) 8,10000);
              sheet.setColumnWidth((short) 9,30000);
            //   sheet.setColumnWidth((short) 10,4000);
            
            // 第三步，创建表格样式
            HSSFCellStyle mainStyle = wb.createCellStyle();//正文样式
            mainStyle.setAlignment(HSSFCellStyle.ALIGN_RIGHT); 
            HSSFFont mainFont = (HSSFFont) wb.createFont();
            mainFont.setFontName("微软雅黑");
            mainFont.setFontHeightInPoints((short) 10);
            mainStyle.setFont(mainFont);
            mainStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN); 
            mainStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN); 
            mainStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
            mainStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
            
            HSSFCellStyle headerStyle = wb.createCellStyle(); //表头样式
            headerStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER); 
            headerStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            headerStyle.setFillForegroundColor(HSSFColor.PALE_BLUE.index);
            HSSFFont headerFont = (HSSFFont) wb.createFont();
            headerFont.setFontName("微软雅黑");
            headerFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
            headerFont.setFontHeightInPoints((short) 14);
            headerStyle.setFont(headerFont);
            headerStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
            headerStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);  
            headerStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN); 
            headerStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
            headerStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
            // 第四步，在sheet中添加表头第0行,注意老版本poi对Excel的行数列数有限制short  
            HSSFRow row = sheet.createRow((int) 0);
            HSSFCell cell = row.createCell((short) 0);  
            cell.setCellValue("申请时间");  
            cell.setCellStyle(headerStyle); 
            cell = row.createCell((short) 1);
            cell.setCellValue("申请人");  
            cell.setCellStyle(headerStyle);  
            cell = row.createCell((short) 2);
            cell.setCellValue(" 所在大区");  
            cell.setCellStyle(headerStyle);  
            cell = row.createCell((short) 3);  
            cell.setCellValue("被调用人姓名");  
            cell.setCellStyle(headerStyle);
            cell = row.createCell((short) 4);  
            cell.setCellValue("被调用人所在部门");  
            cell.setCellStyle(headerStyle);
            cell = row.createCell((short) 5);  
            cell.setCellValue("调用开始时间");  
            cell.setCellStyle(headerStyle);
            cell = row.createCell((short) 6);
            cell.setCellValue("调用结束时间");
            cell.setCellStyle(headerStyle);
            cell = row.createCell((short) 7);
            cell.setCellValue("调用天数");  
            cell.setCellStyle(headerStyle);
             cell = row.createCell((short) 8);
            cell.setCellValue("调用费用");  
            cell.setCellStyle(headerStyle);
             cell = row.createCell((short) 9);
            cell.setCellValue("调用理由");  
            cell.setCellStyle(headerStyle);
            //       cell = row.createCell((short) 10);
            // cell.setCellValue("调用流程状态");  
            // cell.setCellStyle(headerStyle);
            //第五步, 插入正文
            Document []_result = Rdb.getAllDocumentsBySql(sql);
            int i=0;
            for(Document _doc:_result){
                row = sheet.createRow(i + 1);  
                row.createCell((short) 0).setCellValue(_doc.g("WF_DOCCREATED"));
                row.createCell((short) 1).setCellValue(_doc.g("WF_ADDNAME_CN"));
                row.createCell((short) 2).setCellValue(_doc.g("SZDQ"));
                row.createCell((short) 3).setCellValue(_doc.g("DYRY"));
                row.createCell((short) 4).setCellValue(_doc.g("SZDF"));
                row.createCell((short) 5).setCellValue(_doc.g("KSSJ"));
                row.createCell((short) 6).setCellValue(_doc.g("JSSJ"));
                row.createCell((short) 7).setCellValue(_doc.g("DYTS")); 
                row.createCell((short) 8).setCellValue(_doc.g("DYFY")); 
                row.createCell((short) 9).setCellValue(_doc.g("REASON")); 
                // String flag = "";
                // if(_doc.g("FLAG")=="1"){
                //     flag ="调用完成";
                // }else{
                //     flag="在批";
                // }
                // row.createCell((short) 10).setCellValue(flag); 
                //设置正文单元格格式
                for(int j = 0; j < 10; j++){
                    row.getCell((short)j).setCellStyle(mainStyle);
                }
                i++;
            }
            //导出数据到文件
            OutputStream os = null;
            String title="人员调用记录";
            BeanCtx.getResponse().reset(); // 清空输出流  
            os = BeanCtx.getResponse().getOutputStream(); // 取得输出流  
            BeanCtx.getResponse().setHeader("Content-disposition", "attachment; filename="  
            + new String((title+".xls").getBytes("gb2312"), "ISO-8859-1")); // 设定输出文件头
            BeanCtx.getResponse().setContentType("application/msexcel"); // 定义输出类型  
            os.flush();  
            wb.write(os);
        } catch (Exception e) {  
            result="系统提示：Excel文件导出失败，原因："+ e.toString();  
            e.printStackTrace();  
        }
	    return "";
	}
}