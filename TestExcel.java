import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.Collection;

import com.sargeraswang.util.ExcelUtil.ExcelLogs;

import fencer911.ExcelColumnUtil;

public class TestExcel {

	public static void main(String[] args) throws Exception {
	    File f=new File("src/main/java/test.xls");
	    InputStream inputStream= new FileInputStream(f);
	    ExcelLogs logs =new ExcelLogs();


	    Collection<Model> list = new ExcelColumnUtil().importExcel(Model.class, inputStream,null);

	    for(Model m : list){
	      System.out.println(m);
	    }
	}

}
