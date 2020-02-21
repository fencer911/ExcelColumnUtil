# ExcelColumnUtil
用于导入导出Excel的Util包，基于Java的POI(老版本3.8)，把excel数据和javabean进行转换(基于列名的注解而不是列号)
Copy 代码到直接的项目，适合内网的开发，参考了 SargerasWang/ExcelUtil

```java
public class Model {

    @ExcelColumn("姓名")
    private String a;
    @ExcelColumn("年龄")
    private String b;
}    
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
```
