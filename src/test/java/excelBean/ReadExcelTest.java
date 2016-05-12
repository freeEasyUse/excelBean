package excelBean;

import java.util.List;

import org.junit.Test;

import com.gl.excelBean.rwExcel.ReadExcel;
import com.gl.excelBean.vo.Person;
import com.gl.excelBean.vo.SheetVo;

public class ReadExcelTest {

	@Test
	public void testRead() throws Exception{
		List<SheetVo<Person>> result = ReadExcel.readExcel("E:\\workspace\\bingExcel-master\\excel\\src\\test\\resources\\person.xls",1,Person.class);
		System.out.println(result);
	}
	
}
