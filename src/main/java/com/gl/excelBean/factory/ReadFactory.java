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
 * ��ȡ�ļ�����
 */
public class ReadFactory {

	private ReadFactory(){
		
	}
	
	public static Workbook readExcel(String filePath) throws Exception{
		File file = new File(filePath);
		//�ļ��������򴴽��ļ�
		if(!file.exists()){
			file.createNewFile();
		}
		//����excel����
        InputStream xls = new FileInputStream(file);
        //��ǰexcel
        Workbook wb = null;
        //�����ļ���׺�������幤����
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
