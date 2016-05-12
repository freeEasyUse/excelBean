package com.gl.excelBean.vo;

import com.gl.excelBean.annotaion.CellConfig;

public class Person {
	
	public Person(){
		
	}
	
	public Person(int age,String name,Double salary){
		this.age = age;
		this.name = name;
		this.salary = salary;
	}
	
	@CellConfig(index = 1)
	private int age;
	@CellConfig(index = 0)
	private String name;
	@CellConfig(index = 3)
	private Double salary;
	public int getAge() {
		return age;
	}
	public void setAge(int age) {
		this.age = age;
	}
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}
	public Double getSalary() {
		return salary;
	}
	public void setSalary(Double salary) {
		this.salary = salary;
	}
	
	
	
	
}
