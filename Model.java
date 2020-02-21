/**
 * @author SargerasWang
 */


import java.util.Date;

import fencer911.ExcelColumn;
import fencer911.ExcelColumn.Valid;

/**

 * @author https://github.com/fencer911/ExcelColumnUtil
 */
public class Model {

    @ExcelColumn("姓名")
    private String a;
    @ExcelColumn("年龄")
    private String b;
    @ExcelColumn("性别")
    private String c;
    @ExcelColumn("出生日期")
    private Date d;
    
    @ExcelColumn(value="薪水",valid= @Valid(gt=222.0))
    private Double salary;
    
    @ExcelColumn("发薪日")
    private Date salaryDay;    
    
    public Date getD() {
        return d;
    }

    public void setD(Date d) {
        this.d = d;
    }

    public Model() {
		super();
		// TODO Auto-generated constructor stub
	}

	public Model(String a, String b, String c,Date d) {
        this.a = a;
        this.b = b;
        this.c = c;
        this.d = d;
    }

    /**
     * @return the a
     */
    public String getA() {
        return a;
    }

    /**
     * @param a
     *            the a to set
     */
    public void setA(String a) {
        this.a = a;
    }

    /**
     * @return the b
     */
    public String getB() {
        return b;
    }

    /**
     * @param b
     *            the b to set
     */
    public void setB(String b) {
        this.b = b;
    }

    /**
     * @return the c
     */
    public String getC() {
        return c;
    }

    /**
     * @param c
     *            the c to set
     */
    public void setC(String c) {
        this.c = c;
    }

	public Double getSalary() {
		return salary;
	}

	public void setSalary(Double salary) {
		this.salary = salary;
	}

	public Model(String a, String b, String c, Date d, Double salary) {
		super();
		this.a = a;
		this.b = b;
		this.c = c;
		this.d = d;
		this.salary = salary;
	}

	public Date getSalaryDay() {
		return salaryDay;
	}

	public void setSalaryDay(Date salaryDay) {
		this.salaryDay = salaryDay;
	}

	@Override
	public String toString() {
		return "Model [a=" + a + ", b=" + b + ", c=" + c + ", d=" + d + ", salary=" + salary + ", salaryDay="
				+ salaryDay + "]";
	}



    
    
}
