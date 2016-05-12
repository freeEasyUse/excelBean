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
 * ��ȡexcel�ļ�
 */
public class ReadExcel {

	/**
	 * ��ȡexcel�ļ�
	 * @param filePath �ļ�·��
	 * @param startRow ��ʼ��ȡ����
	 * @param c 		��ȡ��������
	 * @return
	 * @throws Exception
	 */
	@SuppressWarnings("all")
	public static <T> List<SheetVo<T>> readExcel(String filePath,int startRow,Class c) throws Exception{
		//�������ؽ��
		List<SheetVo<T>> result = new ArrayList<SheetVo<T>>();
        //��ǰexcel
        Workbook wb = ReadFactory.readExcel(filePath);
        int numerOfSheet = wb.getNumberOfSheets();
        for(int i=0;i<numerOfSheet;i++){
        	SheetVo<T> sheet = new SheetVo<T>();
        	//��������
        	sheet.setSheetIndex(i);
        	Sheet s = wb.getSheetAt(i);
        	//����name
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
