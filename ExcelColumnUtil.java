package cn.fencer911;

import java.awt.Point;
import java.awt.Rectangle;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.text.MessageFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;





/**
 * 
 * @author https://github.com/fencer911/ExcelColumnUtil
 *
 */
public class ExcelColumnUtil {

	public int headerRowNum=0;
	private int maxRows=Integer.MAX_VALUE;
	private static Logger logger = LoggerFactory.getLogger(ExcelColumnUtil.class);
    private static Map<Class<?>, Integer[]> validateMap = new HashMap<Class<?>, Integer[]>();
    static {
        validateMap.put(String[].class, new Integer[]{Cell.CELL_TYPE_STRING});
        validateMap.put(Double[].class, new Integer[]{Cell.CELL_TYPE_NUMERIC});
        validateMap.put(String.class, new Integer[]{Cell.CELL_TYPE_STRING});
        validateMap.put(Double.class, new Integer[]{Cell.CELL_TYPE_NUMERIC});
        validateMap.put(Date.class, new Integer[]{Cell.CELL_TYPE_NUMERIC,Cell.CELL_TYPE_STRING});
        validateMap.put(Integer.class, new Integer[]{Cell.CELL_TYPE_NUMERIC});
        validateMap.put(Float.class, new Integer[]{Cell.CELL_TYPE_NUMERIC});
        validateMap.put(Long.class, new Integer[]{Cell.CELL_TYPE_NUMERIC});
        validateMap.put(Boolean.class, new Integer[]{Cell.CELL_TYPE_NUMERIC});
        validateMap.put(BigDecimal.class, new Integer[]{Cell.CELL_TYPE_NUMERIC});
    }	
	//data will init  name is columnName
	private Map<String,Field>   nameToField=null;
	
	private Map<String,Integer> nameToIndex=null;
	private Map<Integer,String> indexToName=null;
	
	public boolean fileCheck=false;
	private List<ExcelLog> logList=new ArrayList<ExcelLog>();
	
    public List<ExcelLog> getLogList() {
		return logList;
	}
    public   <T> Collection<T> importExcel(Class<T> clazz, InputStream is) throws InstantiationException, IllegalAccessException {
    	return importExcel(clazz,is,null);
    }
	public   <T> Collection<T> importExcel(Class<T> clazz, InputStream is,FileCheck check) throws InstantiationException, IllegalAccessException {
		Workbook workBook;
		try {
			workBook = WorkbookFactory.create(is);
			if(check!=null&&!check.check(workBook)) {
				fileCheck=true;
				return null;
			}
		} catch (Exception e) {
			logger.error("load excel file error",e);
			return null;
		}
		Sheet sheet = workBook.getSheetAt(0);
		Collection<T> list =importExcel(clazz,sheet);
		return list;
	}
	public   <T> Collection<T> importExcel(Class<T> clazz, Sheet sheet) {
		List<T> list = new ArrayList<T>();
		try {
		parseHeader(sheet.getRow(headerRowNum));
		boolean isMap=clazz == Map.class;
		if (!isMap) {
			nameToField=parseClass(clazz); 
		}
		//check excel headerColumn,classField
		System.out.println(nameToField.keySet()+" eq "+this.nameToIndex.keySet());
		if(!nameToField.keySet().containsAll(this.nameToIndex.keySet())) {
			logger.error("excel 标题列与实体字段不匹配!");
			return null;
		}
		int rowIndex=headerRowNum+1;
		maxRows=sheet.getLastRowNum();
		
        while (rowIndex<=maxRows) {
            Row row = sheet.getRow(rowIndex);
            rowIndex++;
            if (row==null||isEmptyRow(row.cellIterator())) {
            	logger.warn("Excel row " + rowIndex + " all row value is null!");
                continue;
            }
            if (isMap) {
                list.add((T) parseDataOfMap(nameToIndex,row));
            } else {
                T t = clazz.newInstance();
                parseData(t,row);
                list.add(t);
            }
        }    
		}catch(Throwable e) {
			e.printStackTrace();
		}
		return list;
	}
    public  <T> T parseData(T t,Row row) {
    	StringBuilder log = new StringBuilder();
        for (String columnName : nameToField.keySet()) {
            Field field = nameToField.get(columnName);
            field.setAccessible(true);
            if(!this.nameToIndex.containsKey(columnName)) {
            	this.logger.warn(columnName+" can not find cell");
            	continue;
            }
            int columnIndex=this.nameToIndex.get(columnName);
            Cell cell = row.getCell(columnIndex);
            if("备注".equals(columnName)||"耗材长度".equals(columnName)) {
//            	System.out.println(row.getRowNum()+columnName);
            }
            String errMsg =null;// validateCell(cell, field, columnName);
            if (isBlank(errMsg)) {
            	Object value=convert(field,cell);
            	try {
					field.set(t, value);
				} catch (Exception e) {
					log.append("save "+columnName+" error").append(";"); 
					e.printStackTrace();
				} 
            }
            if (!isBlank(errMsg)) {
                log.append(errMsg).append(";");    
            }
        }
//        if(log.length()>0) 
        {
	        ExcelLog rowLog= new ExcelLog(t, log.toString(), row.getRowNum() + 1);
	        this.logList.add(rowLog);
        }
    	return t;
    }
    private static  Map<String,Field> parseClass(Class<?> clazz) {
        Field[] fieldsArr = clazz.getDeclaredFields();
        List<Field> fields=new ArrayList<Field>();
        fields.addAll(Arrays.asList(fieldsArr));
        
        if(Object.class!= clazz.getSuperclass()) {
        	fieldsArr=clazz.getSuperclass().getDeclaredFields();
        	if(fieldsArr!=null) {
        		fields.addAll(Arrays.asList(fieldsArr));
        	}
        }
        Map<String,Field> nameToField=new HashMap<String,Field>();
        for (Field field : fields) {
            ExcelColumn ec = field.getAnnotation(ExcelColumn.class);
            if (ec == null) {
                continue;
            }
            String indexName= ec.value();
            nameToField.put(indexName, field);
            field.setAccessible(true);
        }
        return nameToField;
    }
     public static Map<String, Object> parseDataOfMap(Map<String, Integer> titleMap,Row dataRow) {
            Map<String, Object> map = new HashMap<String, Object>();
            for (String k : titleMap.keySet()) {
                Integer index = titleMap.get(k);
                Cell cell = dataRow.getCell(index);
                // 判空
                if (cell == null) {
                    map.put(k, null);
                } else {
                    cell.setCellType(Cell.CELL_TYPE_STRING);
                    String value = cell.getStringCellValue();
                    map.put(k, value);
                }
            }  
            return map;
    }
    public static boolean isEmptyRow(Iterator<Cell> cellIterator) {
        // 整行都空，就跳过
        boolean allRowIsNull = true;
        while (cellIterator.hasNext()) {
            Object cellValue = getCellValue(cellIterator.next());
            if (cellValue != null) {
                allRowIsNull = false;
                break;
            }
        }
        return allRowIsNull;
    }
    public void parseHeader(Row headerRow){
        // 解析map用的key,就是excel标题行
    	this.nameToIndex =new HashMap<String,Integer>();
    	this.indexToName =new HashMap<Integer,String>();
        Iterator<Cell> cellIterator = headerRow.cellIterator();
        Integer index = 0;
        while (cellIterator.hasNext()) {
        	Cell cell=cellIterator.next();
            String columnName = cell.getStringCellValue();
            if(!isBlank(columnName)) {
            	nameToIndex.put(columnName.trim(), cell.getColumnIndex());
            	indexToName.put( cell.getColumnIndex(),columnName.trim());
            }
        }
    }
    public static boolean isBlank(String str){
        if(str == null){
            return true;
        }
        return str.length() == 0?true:str.trim().length()==0?true:false;
    }
    public static  String getCellValueAsString(Cell cell){
        if(cell == null) return "";
        cell.setCellType(Cell.CELL_TYPE_STRING);
        return cell.getStringCellValue();
    }
    public static Object getCellValue(Cell cell) {
    	
        if (cell == null|| (cell.getCellType() == cell.CELL_TYPE_STRING && isBlank(cell.getStringCellValue()))) {
            return null;
        }
        int cellType=cell.getCellType();
            if(cellType == cell.CELL_TYPE_BLANK)
                return null;
            else if(cellType == cell.CELL_TYPE_BOOLEAN)
                return cell.getBooleanCellValue();
            else if(cellType == cell.CELL_TYPE_ERROR)
                return cell.getErrorCellValue();
            else if(cellType == cell.CELL_TYPE_FORMULA) {
                try {
                    if (HSSFDateUtil.isCellDateFormatted(cell)) {
                        return cell.getDateCellValue();
                    } else {
                        return cell.getNumericCellValue();
                    }
                } catch (IllegalStateException e) {
                    return cell.getRichStringCellValue();
                }
            }
            else if(cellType == cell.CELL_TYPE_NUMERIC){
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue();
                } else {
                    return cell.getNumericCellValue();
                }
            }
            else if(cellType == cell.CELL_TYPE_STRING)
                return cell.getStringCellValue();
            else
                return null;
    }
    public Object convert(Field field, Cell cell) {
    	Object value = null; 
    	if (field.getType().equals(Date.class)&& cell.getCellType() == cell.CELL_TYPE_STRING){
    		 Object strDate = getCellValue(cell);
             try {
            	 if(strDate!=null) {
            		 if(strDate.toString().toString().length()<=10) {
            			 value = new SimpleDateFormat("yyyy/MM/dd").parse(strDate.toString());
            		 }else {
            			 value = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss").parse(strDate.toString());
            		 }
            	 }
             } catch (Exception e) {
            	 logger.error(MessageFormat.format("the cell [{0},{1}] can not be converted to a date ", cell.getRowIndex(),cell.getColumnIndex()),e);
             }
    	 }else {
             value = getCellValue(cell);
             // 处理特殊情况,excel的value为String,且bean中为其他,且defaultValue不为空,那就=defaultValue
             ExcelColumn annoCell = field.getAnnotation(ExcelColumn.class);
             if (value!=null&&field.getType().equals(String.class)) {
                value=value.toString();
             }
             if(value instanceof Number &&field.getType().equals(BigDecimal.class)) {
            	 value=new BigDecimal(value.toString());
             }
    	 }
    	 return value;
    }
    public static <T> T convert(Class<T> cls, Object obj) {
    	if(obj!=null) {
    		if(cls==Integer.class) {
    			if(obj instanceof Double)
    				return (T) new Integer(((Double)obj).intValue());
    			if(obj instanceof String)
    				return (T) new Integer(Integer.parseInt(obj.toString())); 			
    		}else {
    			return (T) obj;
    		}
    	}
    	return null;
    }
    private static String validateCell(Cell cell, Field field, String columnName) {
        String result = null;
        Integer[] cellTypeArr = validateMap.get(field.getType());
        if (cellTypeArr == null) {
            result = MessageFormat.format("Unsupported type [{0}]", field.getType().getSimpleName());
            return result;
        }
        if("FROM".equals(columnName)) {
        	System.out.println("");
        }
        //空着处理
        ExcelColumn annoCell = field.getAnnotation(ExcelColumn.class);
        if(annoCell.valid().allowNull()) {
        	return null;
        }
        
        int cellType0=cell.getCellType();
        if (cell!=null&&(cellType0 == Cell.CELL_TYPE_STRING) ) {
        	if(isBlank(cell.getStringCellValue())){
                if (annoCell.valid().allowNull() == false) {
                    return MessageFormat.format("the cell [{0}] can not empty", columnName);
                }
        	}else {
        		return valid(cell,annoCell,columnName);
        	}
        } else if (cellType0 == Cell.CELL_TYPE_BLANK && !annoCell.valid().allowNull()) {
            return MessageFormat.format("the cell [{0}] can not empty", columnName);
        } else if(cellType0 == Cell.CELL_TYPE_BLANK && annoCell.valid().allowNull()) {
        		return null;
        }else {
            List<Integer> cellTypes = Arrays.asList(cellTypeArr);

            // 如果類型不在指定範圍內,並且沒有默認值
            if (!(cellTypes.contains(cell.getCellType()))|| !isBlank(annoCell.defaultValue())
                    && cell.getCellType() == Cell.CELL_TYPE_STRING) {
                StringBuilder strType = new StringBuilder();
                for (int i = 0; i < cellTypes.size(); i++) {
                     int cellType = cellTypes.get(i);
                    strType.append(getCellTypeDesc(cellType));
                    if (i != cellTypes.size() - 1) {
                        strType.append(",");
                    }
                }
                result=MessageFormat.format("the cell [{0}] type must [{1}]", columnName, strType.toString());
            } else {
                // 类型符合验证,但值不在要求范围内的
                // String in
                if (annoCell.valid().in().length != 0 && cell.getCellType() == Cell.CELL_TYPE_STRING) {
                    String[] in = annoCell.valid().in();
                    String cellValue = cell.getStringCellValue();
                    boolean isIn = false;
                    for (String str : in) {
                        if (str.equals(cellValue)) {
                            isIn = true;
                        }
                    }
                    if (!isIn) {
                        result = MessageFormat.format("the cell [{0}] value must in {1}", columnName, in);
                    }
                }
  
            
                if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                    double cellValue = cell.getNumericCellValue();
                    // 小于
                    if (!Double.isNaN(annoCell.valid().lt())) {
                        if (!(cellValue < annoCell.valid().lt())) {
                            result =MessageFormat.format("the cell [{0}] value must less than [{1}]", columnName,annoCell.valid().lt());
                        }
                    }
                    // 大于
                    if (!Double.isNaN(annoCell.valid().gt())) {
                        if (!(cellValue > annoCell.valid().gt())) {
                            result = MessageFormat.format("the cell [{0}] value must greater than [{1}]", columnName,annoCell.valid().gt());
                        }
                    }
                    // 小于等于
                    if (!Double.isNaN(annoCell.valid().le())) {
                        if (!(cellValue <= annoCell.valid().le())) {
                            result =MessageFormat.format("the cell [{0}] value must less than or equal [{1}]",columnName, annoCell.valid().le());
                        }
                    }
                    // 大于等于
                    if (!Double.isNaN(annoCell.valid().ge())) {
                        if (!(cellValue >= annoCell.valid().ge())) {
                            result =MessageFormat.format("the cell [{0}] value must greater than or equal [{1}]",columnName, annoCell.valid().ge());
                        }
                    }
                }
            }
        }
        return result;
    }
    private static String valid(Cell cell, ExcelColumn annoCell, String columnName) {
    	String result=null;
        if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
        	String cellValue = cell.getStringCellValue();
            // 小于
            if (!Double.isNaN(annoCell.valid().lt())) {
                if (!(cellValue.length() < annoCell.valid().lt())) {
                    result =MessageFormat.format("the cell [{0}] length must less than [{1}]", columnName,annoCell.valid().lt());
                }
            }
            // 大于
            if (!Double.isNaN(annoCell.valid().gt())) {
                if (!(cellValue.length() > annoCell.valid().gt())) {
                    result = MessageFormat.format("the cell [{0}] length must greater than [{1}]", columnName,annoCell.valid().gt());
                }
            }
            // 小于等于
            if (!Double.isNaN(annoCell.valid().le())) {
                if (!(cellValue.length() <= annoCell.valid().le())) {
                    result =MessageFormat.format("the cell [{0}] length must less than or equal [{1}]",columnName, annoCell.valid().le());
                }
            }
            // 大于等于
            if (!Double.isNaN(annoCell.valid().ge())) {
                if (!(cellValue.length() >= annoCell.valid().ge())) {
                    result =MessageFormat.format("the cell [{0}] length must greater than or equal [{1}]",columnName, annoCell.valid().ge());
                }
            }
        }        	
    	return result;
    }
    private static String getCellTypeDesc(int cellType) {
        if(cellType == Cell.CELL_TYPE_BLANK)
            return "Null type";
        else if(cellType == Cell.CELL_TYPE_BOOLEAN)
            return "Boolean type";
        else if(cellType == Cell.CELL_TYPE_ERROR)
            return "Error type";
        else if(cellType == Cell.CELL_TYPE_FORMULA)
            return "Formula type";
        else if(cellType == Cell.CELL_TYPE_NUMERIC)
            return "Numeric type";
        else if(cellType == Cell.CELL_TYPE_STRING)
            return "String type";
        else
            return "Unknown type";
    }

    public interface FileCheck {
    	public boolean check(Workbook workBook);
    }
    private Map<String,List<Rectangle>> mergedRegionsMap=new HashMap<String,List<Rectangle>>();
    
    public static List<Rectangle>  getMergedRegionList(Sheet sheet){
        int sheetMergeCount = sheet.getNumMergedRegions();
        List<Rectangle> mergedRegions=new ArrayList<Rectangle>(); 
        for(int i = 0 ; i < sheetMergeCount ; i++){
            CellRangeAddress ca = sheet.getMergedRegion(i);
            
            int firstColumn = ca.getFirstColumn();
            int firstRow = ca.getFirstRow();
            
            int lastColumn = ca.getLastColumn();
            int lastRow = ca.getLastRow();
            Rectangle r=new Rectangle(firstColumn,firstRow,lastColumn-firstColumn+1,lastRow-firstRow+1);
            
            mergedRegions.add(r);
        }
        return mergedRegions;
        //mergedRegionsMap.put(sheet.getSheetName(),mergedRegions);
    }
    public static Rectangle getMergedRegion(List<Rectangle> mergedRegions,int row, int column) {
    	if(mergedRegions!=null) {
    		Point p=new Point(column,row);
    		for(Rectangle r:mergedRegions) {
    			if(r.contains(p)) {
    				return r;
    			}
    		}
    	}
    	return null;
    }
    
    public Cell getCell(Sheet sheet,int row, int column) {
    	return sheet.getRow(row).getCell(column);
    }
    
    public static void printCells(Sheet sheet) {
        System.out.println("--print sheet");
        int rowNum = sheet.getLastRowNum();;
        for(int i=0;i<rowNum;i++) {
        	Row r=sheet.getRow(i);
        	int columnNum=r.getPhysicalNumberOfCells();
        	System.out.print((i+1)+"="+columnNum+"[");
        	
        	for(int j=0;j<columnNum;j++) {
        		Cell c=r.getCell(j);
        		System.out.print(j+"="+getCellValue(c)+",");
        	}
        	System.out.println("]");
        }
    }
    public  static void printBlock(Sheet sheet) {
    	List<Rectangle> mergedRegions=getMergedRegionList(sheet);
        System.out.println("--print sheet");
        int rowNum = sheet.getLastRowNum();;
        for(int i=0;i<rowNum;i++) {
        	Row r=sheet.getRow(i);
        	int columnNum=r.getPhysicalNumberOfCells();
        	
        	for(int j=0;j<columnNum;j++) {
        		Rectangle rct=getMergedRegion(mergedRegions,i,j);
        		if(rct==null) {
        			continue;
        		}
        		Cell c=r.getCell(j);
        		String val=getCellValueAsString(c);
        		if(!isBlank(val)) {
//        			System.out.print("["+i+","+j+"]="+getCellValue(c)+",");
        			System.out.print(rct.toString().replace("java.awt.Rectangle", "")+"="+getCellValue(c)+",");
        		}
        		
        	}
        	System.out.println("");
        }
    }
    
	public static <T> void exportExcel(List<String> headers,Class<T> clazz, Collection<T> dataset, OutputStream out) {
		Map<String, Field> nameToField = parseClass(clazz);
		if(headers==null) {
			headers=new ArrayList<String>(nameToField.keySet());
		}
		// 声明一个工作薄
		HSSFWorkbook workbook = new HSSFWorkbook();
		// 生成一个表格
		HSSFSheet sheet = workbook.createSheet();
		
		
		writeHeader(sheet,headers);
		List<Field> fields=new ArrayList<Field>(headers.size());
		for(String key:headers){
			Field f=nameToField.get(key);
			if(f!=null) {
				fields.add(f);
			}
		}
		writeData(sheet,fields, dataset, "");
		setColumnWidth(sheet,headers.size());
		try {
			workbook.write(out);
		} catch (IOException e) {
			logger.error(e.toString(), e);
		}
	}

	private static void writeHeader(HSSFSheet sheet,List<String> headers){
		//标题格式
		HSSFCellStyle titleStyle = sheet.getWorkbook().createCellStyle();
		titleStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		titleStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
		setCellBorderStyle(titleStyle);
		setBackgroundStyle(titleStyle, HSSFColor.SKY_BLUE.index);
		setSimpleCellFontStyle(sheet.getWorkbook(), titleStyle, (short)13, HSSFColor.BLACK.index);
		// 产生表格标题行
		HSSFRow row = sheet.createRow(0);
		// 标题行转中文
		int c = 0; // 标题列号
		for(String key:headers){
				HSSFCell cell = row.createCell(c);
				cell.setCellStyle(titleStyle);
				HSSFRichTextString text = new HSSFRichTextString(key);
				cell.setCellValue(text);
				c++;
		}
	}
	private static  HSSFCellStyle setCellBorderStyle(HSSFCellStyle cellStyle){
		cellStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		cellStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
		cellStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
		cellStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		return cellStyle;
	}
	/**
	 * 设置字体
	 * @param workbook
	 * @param cellStyle
	 * @param size
	 * @param color
	 * @return
	 */
	private  static HSSFCellStyle setSimpleCellFontStyle(HSSFWorkbook workbook,HSSFCellStyle cellStyle, short size, short color){
		HSSFFont font = workbook.createFont();
		font.setFontHeightInPoints(size);
		font.setColor(color);
		cellStyle.setFont(font);
		return cellStyle;
	}
	/**
	 * 设置背景色
	 * @param cellStyle
	 * @param color
	 * @return
	 */
	private static HSSFCellStyle setBackgroundStyle(HSSFCellStyle cellStyle, short color){
		cellStyle.setFillForegroundColor(color);
		cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		return cellStyle;
	}	
	// 让列宽随着导出的列长自动适应
	public static void setColumnWidth(HSSFSheet sheet,int columnNum){
		for (int colNum = 0; colNum < columnNum; colNum++) {
			int columnWidth = sheet.getColumnWidth(colNum) / 256;
			for (int rowNum = 0; rowNum < sheet.getLastRowNum(); rowNum++) {
				HSSFRow currentRow;
				if (sheet.getRow(rowNum) == null) {
					currentRow = sheet.createRow(rowNum);
				} else {
					currentRow = sheet.getRow(rowNum);
				}
				if (currentRow.getCell(colNum) != null) {
					HSSFCell currentCell = currentRow.getCell(colNum);
					if (currentCell.getCellType() == HSSFCell.CELL_TYPE_STRING) {
						if(currentCell.getStringCellValue() != null){
							int length = currentCell.getStringCellValue()
									.getBytes().length;
							if (columnWidth < length) {
								columnWidth = length;
							}
						}
					}
				}
			}
			if(columnWidth > 200){
				columnWidth = 200;
			}
			if (colNum == 0) {
				sheet.setColumnWidth(colNum, (columnWidth - 2) * 256);
			} else {
				sheet.setColumnWidth(colNum, (columnWidth + 4) * 256);
			}
		}
	}	
	
	private static <T> void writeData(HSSFSheet sheet, List<Field> fields, Collection<T> dataset,String pattern) {
		// 时间格式默认"yyyy-MM-dd"
		if (isBlank(pattern)) {
			pattern = "yyyy-MM-dd";
		}
		// 遍历集合数据，产生数据行
		Iterator<T> it = dataset.iterator();
		int index = 0;
		while (it.hasNext()) {
			index++;
			HSSFRow row  = sheet.createRow(index);
			T t = it.next();
			try {
					int cellNum = 0;
					for (int i = 0; i < fields.size(); i++) {
						HSSFCell cell = row.createCell(cellNum);
						Field field = fields.get(i);
						field.setAccessible(true);
						Object value = field.get(t);

						cellNum = setCellValue(cell, value, pattern, cellNum, field, row);

						cellNum++;
					}
			} catch (Exception e) {
				logger.error(e.toString(), e);
			}
		}
		// 设定自动宽度
		for (int i = 0; i < fields.size(); i++) {
			sheet.autoSizeColumn(i);
		}
	}
	
    private static int setCellValue(HSSFCell cell,Object value,String pattern,int cellNum,Field field,HSSFRow row){
        String textValue = null;
        if (value instanceof Integer) {
            int intValue = (Integer) value;
            cell.setCellValue(intValue);
        } else if (value instanceof Float) {
            float fValue = (Float) value;
            cell.setCellValue(fValue);
        } else if (value instanceof Double) {
            double dValue = (Double) value;
            cell.setCellValue(dValue);
        } else if (value instanceof Long) {
            long longValue = (Long) value;
            cell.setCellValue(longValue);
        } else if (value instanceof Boolean) {
            boolean bValue = (Boolean) value;
            cell.setCellValue(bValue);
        } else if (value instanceof Date) {
            Date date = (Date) value;
            SimpleDateFormat sdf = new SimpleDateFormat(pattern);
            textValue = sdf.format(date);
        } else if (value instanceof String[]) {
            String[] strArr = (String[]) value;
            for (int j = 0; j < strArr.length; j++) {
                String str = strArr[j];
                cell.setCellValue(str);
                if (j != strArr.length - 1) {
                    cellNum++;
                    cell = row.createCell(cellNum);
                }
            }
        } else if (value instanceof Double[]) {
            Double[] douArr = (Double[]) value;
            for (int j = 0; j < douArr.length; j++) {
                Double val = douArr[j];
                // 值不为空则set Value
                if (val != null) {
                    cell.setCellValue(val);
                }

                if (j != douArr.length - 1) {
                    cellNum++;
                    cell = row.createCell(cellNum);
                }
            }
        } else {
            // 其它数据类型都当作字符串简单处理
            String empty = "";
            if(field != null) {
                ExcelColumn anno = field.getAnnotation(ExcelColumn.class);
                if (anno != null) {
                    empty = anno.defaultValue();
                }
            }
            textValue = value == null ? empty : value.toString();
        }
        if (textValue != null) {
            HSSFRichTextString richString = new HSSFRichTextString(textValue);
            cell.setCellValue(richString);
        }
        return cellNum;
    }	
}
