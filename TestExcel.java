import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.Collection;



public class TestExcel {

	public static void main(String[] args) throws Exception {
	    File f=new File("src/main/java/test.xls");
	    InputStream inputStream= new FileInputStream(f);



	    Collection<Model> list = new ExcelColumnUtil().importExcel(Model.class, inputStream);

	    for(Model m : list){
	      System.out.println(m);
	    }
	}

}
