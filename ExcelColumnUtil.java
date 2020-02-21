

import java.io.InputStream;
import java.lang.reflect.Field;
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

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;

import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.sargeraswang.util.ExcelUtil.ExcelLogs;
import com.sargeraswang.util.ExcelUtil.ExcelUtil;



/**
 * 
 * @author https://github.com/fencer911/ExcelColumnUtil
 *
 */
public class ExcelColumnUtil {

	private int headerRowNum=0;
	private int maxRows=Integer.MAX_VALUE;
	private static Logger logger = LoggerFactory.getLogger(ExcelUtil.class);
    private static Map<Class<?>, Integer[]> validateMap = new HashMap<>();
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
    }	
	//data will init  name is columnName
	private Map<String,Field>   nameToField=null;
	
	private Map<String,Integer> nameToIndex=null;
	private Map<Integer,String> indexToName=null;
	
	//
	StringBuilder logList = new StringBuilder();
    public   <T> Collection<T> importExcel(Class<T> clazz, InputStream is, ExcelLogs logs) throws InstantiationException, IllegalAccessException {
		Workbook workBook;
		try {
			workBook = WorkbookFactory.create(is);
		} catch (Exception e) {
			logger.error("load excel file error",e);
			return null;
		}
		Sheet sheet = workBook.getSheetAt(0);
		parseHeader(sheet.getRow(headerRowNum));
		boolean isMap=clazz == Map.class;
		if (!isMap) {
			nameToField=parseClass(clazz); 
		}
		
		int rowIndex=headerRowNum+1;
		maxRows=sheet.getLastRowNum();
		List<T> list = new ArrayList<>();
        while (rowIndex<=maxRows) {
            Row row = sheet.getRow(rowIndex);
            rowIndex++;
            if (isEmptyRow(row.cellIterator())) {
            	logger.warn("Excel row " + row.getRowNum() + " all row value is null!");
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
        logger.debug(this.logList.toString());
		return list;
	}
    public  <T> T parseData(T t,Row row) {
        for (String columnName : nameToField.keySet()) {
            Field field = nameToField.get(columnName);
            field.setAccessible(true);
            if(!this.nameToIndex.containsKey(columnName)) {
            	this.logger.warn(columnName+" can not find cell");
            	continue;
            }
            int columnIndex=this.nameToIndex.get(columnName);
            Cell cell = row.getCell(columnIndex);
            String errMsg = validateCell(cell, field, columnName);
            if (isBlank(errMsg)) {
            	Object value=convert(field,cell);
            	try {
					field.set(t, value);
				} catch (Exception e) {
					e.printStackTrace();
				} 
            }
            if (!isBlank(errMsg)) {
            	logList.append("error at "+row.getRowNum() + 1+" msg:"+errMsg+".\n");
            }	 
        }
    	return t;
    }
    private static  Map<String,Field> parseClass(Class<?> clazz) {
        Field[] fieldsArr = clazz.getDeclaredFields();
        Map<String,Field> nameToField=new HashMap<String,Field>();
        for (Field field : fieldsArr) {
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
            Map<String, Object> map = new HashMap<>();
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
    private static boolean isBlank(String str){
        if(str == null){
            return true;
        }
        return str.length() == 0?true:str.trim().length()==0?true:false;
    }
    private static Object getCellValue(Cell cell) {
    	
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
                 value = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss").parse(strDate.toString());
             } catch (Exception e) {
            	 logger.error(MessageFormat.format("the cell [{0},{1}] can not be converted to a date ", cell.getRowIndex(),cell.getColumnIndex()),e);
             }
    	 }else {
             value = getCellValue(cell);
             // 处理特殊情况,excel的value为String,且bean中为其他,且defaultValue不为空,那就=defaultValue
             ExcelColumn annoCell = field.getAnnotation(ExcelColumn.class);
             if (value instanceof String && !field.getType().equals(String.class)) {
                //nothing do
             }
    	 }
    	 return value;
    }
    private static String validateCell(Cell cell, Field field, String columnName) {
        String result = null;
        Integer[] cellTypeArr = validateMap.get(field.getType());
        if (cellTypeArr == null) {
            result = MessageFormat.format("Unsupported type [{0}]", field.getType().getSimpleName());
            return result;
        }
        ExcelColumn annoCell = field.getAnnotation(ExcelColumn.class);
        if (cell == null|| (cell.getCellType() == Cell.CELL_TYPE_STRING && isBlank(cell.getStringCellValue()))) {
            if (annoCell != null && annoCell.valid().allowNull() == false) {
                result = MessageFormat.format("the cell [{0}] can not null", columnName);
            }
        } else if (cell.getCellType() == Cell.CELL_TYPE_BLANK && annoCell.valid().allowNull()) {
            return result;
        } else {
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
                // 数字型
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

}
