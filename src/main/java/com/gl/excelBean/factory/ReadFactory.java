package com.gl.excelBean.factory;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * 
 * @author GeL
 * 读取文件工厂
 */
public class ReadFactory {

	private ReadFactory(){
		
	}
	
	public static Workbook readExcel(String filePath) throws Exception{
		File file = new File(filePath);
		//文件不存在则创建文件
		if(!file.exists()){
			file.createNewFile();
		}
		//进行excel解析
        InputStream xls = new FileInputStream(file);
        //当前excel
        Workbook wb = null;
        //根据文件后缀创建具体工作蒲
        if (file.getName().endsWith("xls")) {
            wb = new HSSFWorkbook(xls);
        } else if (file.getName().endsWith("xlsx")) {
            wb = new XSSFWorkbook(xls);
        }
		return wb;
	}
	
	
	
	public static OutputStream writeFile(String path) throws Exception{
		File file = new File(path);
		if(!file.exists()){
			file.createNewFile();
		}
		return new FileOutputStream(file);
	}
	
}
