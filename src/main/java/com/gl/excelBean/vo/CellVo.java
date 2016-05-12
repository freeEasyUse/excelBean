package com.gl.excelBean.vo;

public class CellVo {
	
	public CellVo(){
		
	}
	
	public CellVo(int index,Object value){
		this.index = index;
		this.value = value;
	}
	
	private int index;
	
	private Object value;

	public int getIndex() {
		return index;
	}

	public void setIndex(int index) {
		this.index = index;
	}

	public Object getValue() {
		return value;
	}

	public void setValue(Object value) {
		this.value = value;
	}
}
