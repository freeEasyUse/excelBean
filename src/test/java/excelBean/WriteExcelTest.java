package excelBean;

import java.util.ArrayList;
import java.util.List;

import org.junit.Test;

import com.gl.excelBean.rwExcel.WriteExcel;
import com.gl.excelBean.vo.Person;


public class WriteExcelTest {

	@SuppressWarnings("unchecked")
	@Test
	public void testWrite() throws Exception{
		List<Person> list = new ArrayList<Person>();
		list.add(new Person(12, "nihoa", 23434.9));
		list.add(new Person(23, "nihoa", 234.9));
		list.add(new Person(122, "nihoa", 23434.9));
		WriteExcel.writeExcel("E:/workspace/excelBean/src/test/resources/123.xlsx", Person.class, list,list);
	}
	
	
	@Test
	public void testWriteappend() throws Exception{
		List<Person> list = new ArrayList<Person>();
		list.add(new Person(1223, "nihoaf", 23434.9));
		list.add(new Person(2332, "nihoaf", 234.9));
		list.add(new Person(123222, "nihoaf", 23434.9));
		WriteExcel.appendWrite("E:/workspace/excelBean/src/test/resources/123.xlsx",list,Person.class);
	}
	
}
