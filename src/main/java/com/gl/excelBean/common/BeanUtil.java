package com.gl.excelBean.common;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.beanutils.BeanUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import com.gl.excelBean.annotaion.CellConfig;
import com.gl.excelBean.vo.CellVo;

public class BeanUtil {
	
	/**
	 * 将excel行转换层对应的bean
	 * @param t
	 * @param row
	 * @return
	 * @throws Exception 
	 * @throws  
	 */
	public static Object rowToBean(Object t,Row row) throws Exception{
		Field[] fields = t.getClass().getDeclaredFields();
		for(Field f:fields){
			CellConfig cellConfig = f.getAnnotation(CellConfig.class);
			int cellIndex = cellConfig.index();
			Cell cell = row.getCell(cellIndex);
			int cellType = cell.getCellType();
			switch (cellType) {
			case Cell.CELL_TYPE_STRING:
				BeanUtils.setProperty(t, f.getName(), cell.getStringCellValue());
				break;
			case Cell.CELL_TYPE_NUMERIC:
				BeanUtils.setProperty(t, f.getName(), cell.getNumericCellValue());
				break;
			default:
				break;
			}
		}
		return t;
	}
	
	
	public static <T> List<CellVo> beanToCell(T t,Class<T> c) throws Exception{
		List<CellVo> record = new ArrayList<CellVo>();
		//当前记录为空 则是创建heard
		if(t==null){
			Field[] fields = c.getDeclaredFields();
			for(Field f:fields){
				CellConfig cellConfig = f.getAnnotation(CellConfig.class);
				CellVo cellVo = new CellVo(cellConfig.index(),f.getName());
				record.add(cellVo);
			}
		}
		else{
			Field[] fields = t.getClass().getDeclaredFields();
			for(Field f:fields){
				CellConfig cellConfig = f.getAnnotation(CellConfig.class);
				CellVo cellVo = new CellVo(cellConfig.index(),BeanUtils.getProperty(t, f.getName()));
				record.add(cellVo);
			}			
		}
		return record;
	}
	
}
