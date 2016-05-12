package com.gl.excelBean.rwExcel;


import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.gl.excelBean.common.BeanUtil;
import com.gl.excelBean.factory.ReadFactory;
import com.gl.excelBean.vo.SheetVo;

/**
 * 
 * @author GeL
 * 读取excel文件
 */
public class ReadExcel {

	/**
	 * 读取excel文件
	 * @param filePath 文件路径
	 * @param startRow 开始读取的行
	 * @param c 		读取解析对象
	 * @return
	 * @throws Exception
	 */
	@SuppressWarnings("all")
	public static <T> List<SheetVo<T>> readExcel(String filePath,int startRow,Class c) throws Exception{
		//声明返回结果
		List<SheetVo<T>> result = new ArrayList<SheetVo<T>>();
        //当前excel
        Workbook wb = ReadFactory.readExcel(filePath);
        int numerOfSheet = wb.getNumberOfSheets();
        for(int i=0;i<numerOfSheet;i++){
        	SheetVo<T> sheet = new SheetVo<T>();
        	//设置索引
        	sheet.setSheetIndex(i);
        	Sheet s = wb.getSheetAt(i);
        	//设置name
        	sheet.setSheetName(s.getSheetName());
        	List<T> list = ReadExcel.getListFromSheet(s,startRow,c);
        	sheet.setList(list);
        	result.add(sheet);
        }
		return result;
	}
	
	
	
	
	
	@SuppressWarnings("all")
	private static <T> List<T> getListFromSheet(Sheet sheet,int startRow,Class c) throws Exception{
		List<T> result = new ArrayList<T>();
		int allRows = sheet.getPhysicalNumberOfRows();
		for(int i = startRow;i<allRows;i++){
			Row row = sheet.getRow(i);
			T t = (T) c.newInstance();
			result.add((T) BeanUtil.rowToBean(t, row));
		}
		return result;
	}
}
