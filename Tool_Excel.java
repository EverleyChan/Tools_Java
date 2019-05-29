package com.vfcc.util;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.util.List;
import java.util.Map;
import java.util.UUID;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;

import jxl.CellView;
import jxl.Workbook;
import jxl.format.UnderlineStyle;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableWorkbook;

/**
 * 用于导出BOSS平台设备列表到Excel
 * 设备名称、设备IP、平台、机房、ISP、SN
 * hostname、ip、platform、pop、isp_name、device_id
 * @author chenwx
 *
 */
public class Tool_Excel {
	private List<String> excelHeader = null;
	private List<String> excelKey = null;
	private List<Map> excelBody = null;
	private String fileName = "";
	private String fileUrl = "";
//	private static  final Tool_Excel excel = new Tool_Excel();
	
	public Tool_Excel(List<String> head,List<String> body,List<Map> list,String name,String fileUrl){
		System.out.println("配置Excel属性：" + fileName);
		this.excelHeader = head;
		this.excelKey = body;
		this.excelBody = list;
		this.fileName = name;
		this.fileUrl = fileUrl;
	}
//	public static Tool_Excel getInstance(){ return excel;}
	
	public void setExcel(List<String> head,List<String> body,List<Map> list,String name){
		System.out.println("配置Excel属性：" + fileName);
		this.excelHeader = head;
		this.excelKey = body;
		this.excelBody = list;
		this.fileName = name;
		
//		File file = new File(this.fileUrl+this.fileName);
//		if(file.exists()){
//			System.out.println(fileUrl+fileName +" 已存在，先删除！");
//			file.delete();
//		}
	}
	

    /**
     * 递归删除目录下的所有文件及子目录下所有文件
     * 最后删除该目录
     * @param dir 将要删除的文件目录
     * @return boolean Returns "true" if all deletions were successful.
     *                 If a deletion fails, the method stops attempting to
     *                 delete and returns "false".
     */
    private static boolean deleteDir(File dir) {
        if (dir.isDirectory()) {
            String[] children = dir.list();
//　　　　　　　递归删除目录中的子目录下
            for (int i=0; i<children.length; i++) {
                boolean success = deleteDir(new File(dir, children[i]));
                if (!success) {
                    return false;
                }
            }
        }
        // 目录此时为空，可以删除
        return dir.delete();
    }
	
	
	
	public String creatExcel(){
		String id_name = "";
		try {
			System.out.println("生成Excel文件：" + fileName);
			//构建Workbook对象, 只读Workbook对象
			System.out.println("原文件名：" + fileName);
			id_name = makeFileName(fileName);
			System.out.println("文件标识名：" + id_name);
			System.out.println("保存原始路径：" + fileUrl);
			String targetUrl = makePath(id_name, fileUrl);
			System.out.println("保存最终路径：" + targetUrl);
//			linux
			String excel_name = targetUrl + "/" + id_name;
//			windows
//			String excel_name = targetUrl + "\\" + id_name;
			jxl.write.WritableWorkbook wwb = Workbook.createWorkbook(new File(excel_name));
			
			//将WritableWorkbook直接写入到输出流
//			ByteArrayOutputStream os = new ByteArrayOutputStream();               
//			WritableWorkbook wwb = Workbook.createWorkbook(os);  
			//创建Excel工作表
		    jxl.write.WritableSheet ws = wwb.createSheet("Sheet 1", 0);
		    /**
		     * 定义单元格样式
		     * 定义格式 字体 下划线 斜体 粗体 颜色
		     */
		     WritableFont wf = new WritableFont(WritableFont.TIMES, 12,WritableFont.BOLD, false, UnderlineStyle.NO_UNDERLINE,jxl.format.Colour.BLACK); 
		     WritableCellFormat wcf = new WritableCellFormat(wf); // 单元格定义
		     wcf.setBackground(jxl.format.Colour.WHITE); // 设置单元格的背景颜色
		     wcf.setAlignment(jxl.format.Alignment.CENTRE); // 设置对齐方式
			  //表头
		     for(int i=0; i<excelHeader.size();i++){
		    	String title_label = excelHeader.get(i)+"";
		    	Label label = new Label(i,0,title_label,wcf);
			    //ws.setColumnView(i, 50); // 设置列(+1)的宽度
			    ws.setRowView(0, 500); // 设置行(+1)的高度
			    ws.addCell(label);
		     }
			    //表内容
		     for (int i=0 ; i < excelBody.size(); i++){
		    	Map map = (Map)excelBody.get(i);
		    	for(int j=0; j < excelKey.size(); j++){
		    		String key_String = (String) excelKey.get(j);
		    		String value = map.get(key_String)+"";
		    		Label label = new Label(j,i+1,value);
		    		CellView cellView = new CellView();  
		    		cellView.setAutosize(true); //设置自动大小  
		    		String tit =(String) excelHeader.get(j);
		    		if(value.length() > tit.length()){
		    			ws.setColumnView(j,cellView); // 设置列(+1)的宽度
		    		}
		    		else{
		    			ws.setColumnView(j,20); // 设置列(+1)的宽度
		    		}
			    	ws.addCell(label);  		    		
		    	}
		     }
		     //写入Exel工作表
		     wwb.write();	
		     //关闭Excel工作薄对象
		     wwb.close();
//		     HttpServletResponse response
//		    response.reset();
//	        response.setContentType("application/msexcel;charset=utf-8");
//	        response.setHeader("Content-disposition", "attachment;filename= "+new String(fileName.getBytes("gb2312"),"iso8859-1")); 
//	        ServletOutputStream out = response.getOutputStream();
//	        os.writeTo(response.getOutputStream());
//	        out.flush();
//	        out.close();
		} catch (Exception e) {
			// TODO: handle exception
			System.out.println("生成Excel失败！");
			e.printStackTrace();
		}
		return id_name;
	}
	
	/**
	 * 生成上传文件的文件名，文件名以：uuid+"_"+文件的原始名称
	 * @param file_name 文件的原始名称
	 * @return uuid+"_"+文件的原始名称
	 */
	private String makeFileName(String file_name){
//		为防止文件覆盖的现象发生，要为上传文件产生一个唯一的文件名
		return UUID.randomUUID().toString() + "_" + file_name;
	}
	private String makePath(String filename,String savePath){
//		得到文件名的hashCode的值，得到的就是filename这个字符串对象在内存中的地址
		int hashcode = filename.hashCode();
		int dir1 = hashcode&0xf;  //0--15
		int dir2 = (hashcode&0xf0)>>4;  //0-15
//		构造新的保存目录,windows
//		String dir = savePath + "\\" + dir1 + "\\" + dir2;  //upload\2\3  upload\3\5
//		构造新的保存目录,linux
		String dir = savePath + "/" + dir1 + "/" + dir2;  //upload/2/3  upload/3/5
//		File既可以代表文件也可以代表目录
		File file = new File(dir);
		if(!file.exists()){
			 file.mkdirs();
		 }
		return dir;
	}
}
