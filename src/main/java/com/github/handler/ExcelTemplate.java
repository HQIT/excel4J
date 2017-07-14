package com.github.handler;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import com.github.data.Position;
import com.github.sink.IExcelSink;
import com.github.source.IExcelSource;

import java.io.*;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * <p></p></br>
 * author : wbhe2</br>
 * date  : 2017/6/15  17:37</br>
 */
public class ExcelTemplate {

    /**
     * 当前工作簿
     */
    private Workbook workbook;
    /**
     * 当前工作sheet表
     */
    private Sheet sheet;
    /**
     * 当前表编号
     */
    private int sheetIndex;
    
    /**
     * 当前行
     */
    private Row currentRow;
    /**
     * 当前单元格样式
     */
    private CellStyle currentStyle;
    /**
     * 当前列数
     */
    private int currentColumnIndex;
    /**
     * 当前行数
     */
    private int currentRowIndex;
    /**
     * 默认样式
     */
    private CellStyle defaultStyle;
    /**
     * 指定行样式
     */
    private Map<Integer, CellStyle> appointLineStyle = new HashMap<>();
    /**
     * 分类样式模板
     */
    private Map<String, CellStyle> classifyStyle = new HashMap<>();
    /**
     * 单数行样式
     */
    private CellStyle singleLineStyle;
    /**
     * 双数行样式
     */
    private CellStyle doubleLineStyle;
    /**
     * 数据的初始化列数
     */
    private int initColumnIndex;
    /**
     * 数据的初始化行数
     */
    private int initRowIndex;

    /**
     * 最后一行的数据
     */
    private int lastRowIndex;
    /**
     * 默认行高
     */
    private float rowHeight;
    /**
     * 序号坐标点
     */
    private int serialNumberColumnIndex = -1;
    /**
     * 当前序号
     */
    private int serialNumber;

    private ExcelTemplate() {
    }

    public static ExcelTemplate getInstance(IExcelSource excelSource, int sheetIndex) {
        ExcelTemplate template = new ExcelTemplate();
        template.sheetIndex = sheetIndex;
        try {
            template.loadTemplate(excelSource);
        } catch (Exception e) {
            e.printStackTrace();
        }
        return template;
    }
    
    
    /***********************************初始化模板开始***********************************/

    private ExcelTemplate loadTemplate(IExcelSource excelSource) throws Exception {
        this.workbook = excelSource.getWorkBook();
        this.sheet = this.workbook.getSheetAt(this.sheetIndex);
       // workbook.setSheetName(sheetIndex, sheetName);
        initModuleConfig();
        this.currentRowIndex = this.initRowIndex;
        this.currentColumnIndex = this.initColumnIndex;
        this.lastRowIndex = this.sheet.getLastRowNum();
        return this;
    }

    /**
     * 初始化数据信息
     */
    private void initModuleConfig() {

        for (Row row : sheet) {
            for (Cell c : row) {
                if (c.getCellType() != Cell.CELL_TYPE_STRING)
                    continue;
                String str = c.getStringCellValue().trim();
                // 寻找序号列
                if (str.equals(HanderConstant.SERIAL_NUMBER)) {
                    this.serialNumberColumnIndex = c.getColumnIndex();
                }
                // 寻找数据列
                if (str.equals(HanderConstant.DATA_INIT_INDEX)) {
                    this.initColumnIndex = c.getColumnIndex();
                    this.initRowIndex = row.getRowNum();
                    this.rowHeight = row.getHeightInPoints();
                }
                // 初始化自定义模板样式
                initStyles(c, str);
            }
        }
    }

    /**
     * 初始化样式信息
     */
    private void initStyles(Cell cell, String moduleContext) {

        if (HanderConstant.DEFAULT_STYLE.equals(moduleContext)) {
            this.defaultStyle = cell.getCellStyle();
            clearCell(cell);
        }
        if (null != moduleContext && moduleContext.startsWith("&")) {
            this.classifyStyle.put(moduleContext.substring(1), cell.getCellStyle());
            clearCell(cell);
        }
        if (HanderConstant.APPOINT_LINE_STYLE.equals(moduleContext)) {
            this.appointLineStyle.put(cell.getRowIndex(), cell.getCellStyle());
            clearCell(cell);
        }
        if (HanderConstant.SINGLE_LINE_STYLE.equals(moduleContext)) {
            this.singleLineStyle = cell.getCellStyle();
            clearCell(cell);
        }
        if (HanderConstant.DOUBLE_LINE_STYLE.equals(moduleContext)) {
            this.doubleLineStyle = cell.getCellStyle();
            clearCell(cell);
        }
    }

    private void clearCell(Cell cell) {
        cell.setCellStyle(null);
        cell.setCellValue("");
    }

    /***********************************初始化模板结束***********************************/


    /*************************************数据填充开始***********************************/

    
    /**
     * 根据map替换相应的常量，通过Map中的值来替换#开头的值
     *
     * @param data 替换映射
     */
    public void extendData(Map<String, String> data) {
        if (data == null)
            return;
        for (Row row : this.sheet) {
            for (Cell c : row) {
                if (c.getCellType() != Cell.CELL_TYPE_STRING)
                    continue;
                String str = c.getStringCellValue().trim();
                if (str.startsWith("#")) {
                    if (data.containsKey(str.substring(1))) {
                        c.setCellValue(data.get(str.substring(1)));
                    }
                }
            }
        }
    }
    
    public CellStyle getDefaultStyle(){
    	return defaultStyle;
    }
    public CellStyle getDoubleLineStyle(){
    	return doubleLineStyle;
    }
    public CellStyle getSingleLineStyle(){
    	return singleLineStyle;
    }
    
    public CellStyle getCellStyle( String styleKey) {
    	CellStyle cellstyle = null;
    	if (null != styleKey && null != this.classifyStyle.get(styleKey)) {
    		cellstyle = this.classifyStyle.get(styleKey);
        }
    	else
    		cellstyle = this.defaultStyle;
        return cellstyle;
    }
    
    public Sheet getSheet(){
    	return this.sheet;
    }
    
    /*
    public static void setRegionStyle(XSSFSheet sheet, Region region, XSSFCellStyle cs) {
    	  for (int i = region.getRowFrom(); i <= region.getRowTo(); i++) {
    	   HSSFRow row = HSSFCellUtil.getRow(i, sheet);
    	   for (int j = region.getColumnFrom(); j <= region.getColumnTo(); j++) {
    	    HSSFCell cell = HSSFCellUtil.getCell(row, (short) j);
    	    cell.setCellStyle(cs);
    	   }
    	  }
    	 }
    
    */
    
    public Map<String, Cell> getExtendData(Map<String, String> data){
    	Map<String, Cell> results = new HashMap<>();
    	if (data == null)
            return null;
        for (Row row : this.sheet) {
            for (Cell c : row) {
                if (c.getCellType() != Cell.CELL_TYPE_STRING)
                    continue;
                String str = c.getStringCellValue().trim();
                if (str.startsWith("#")) {
                    if (data.containsKey(str.substring(1))) {
                    	results.put(str.substring(1), c);
                    }
                }
            }
        }
        
        return results;
    }

    
    public String[] getExtendDataList(Map<String, String> extendMap){
	    String[] datalist  = new String[extendMap.size()];
    	String[] keySet = new String[extendMap.keySet().size()];
    	extendMap.keySet().toArray(keySet); 
    	for (int c = 0; c < extendMap.keySet().size(); c++) 
    	{
    		datalist[c] = extendMap.get(keySet[c]);
    	}
		return datalist;
    }
    
    public Cell[] getExtendDataCell(Map<String, String> extendMap, Map<String, Cell> extendData){
    	
	    Cell[] cells = new Cell[extendMap.size()];
	    String[] keySet = new String[extendMap.keySet().size()];
	    extendMap.keySet().toArray(keySet); 
	    for (int c = 0; c < extendMap.keySet().size(); c++) 
	    {
	    	cells[c] = extendData.get(keySet[c]);
	    }
		return cells;
    }
    
    public  String[] getCellNames(int rowIndex) throws IOException {
    	String[] testArray = null;
        Row row = this.sheet.getRow(rowIndex);
        testArray = new String[row.getPhysicalNumberOfCells()];
        for(int column=0;column<row.getPhysicalNumberOfCells();column++){
        	testArray[column] = row.getCell(column).getStringCellValue();
        }
        return testArray;
    }
    public  int  getSerialNumberColumnIndex() {
		return serialNumberColumnIndex;
    }
    
    public Position getInitPosition(){
    	return new Position(this.initRowIndex, this.initColumnIndex);
    }
    
    /**
     * 创建新行，在使用时只要添加完一行，需要调用该方法创建
     */
    public void createNewRow() {
        if (this.lastRowIndex > this.currentRowIndex && this.currentRowIndex != this.initRowIndex) {
            this.sheet.shiftRows(this.currentRowIndex, this.lastRowIndex, 1, true, true);
            this.lastRowIndex++;
        }
        this.currentRow = this.sheet.createRow(this.currentRowIndex);
        this.currentRow.setHeightInPoints(this.rowHeight);
        this.currentRowIndex++;
        this.currentColumnIndex = this.initColumnIndex;
    }

    /**
     * 插入序号，会自动找相应的序号标示的位置完成插入
     */
    public void insertSerial(String styleKey) {
        if (this.serialNumberColumnIndex < 0)
            return;
        this.serialNumber++;
        Cell c = this.currentRow.createCell(this.serialNumberColumnIndex);
        setCellStyle(c, styleKey);
        c.setCellValue(this.serialNumber);
    }

    /**
     * <p>设置Excel元素样式及内容</p></br>
     */
    public void createCell(Object value, String styleKey) {
        Cell cell = this.currentRow.createCell(currentColumnIndex);
        setCellStyle(cell, styleKey);
        if (null == value || "".equals(value)) {
            this.currentColumnIndex++;
            return;
        }

        if (String.class == value.getClass()) {
            cell.setCellValue((String) value);
            this.currentColumnIndex++;
            return;
        }

        if (int.class == value.getClass()) {
            cell.setCellValue((int) value);
            this.currentColumnIndex++;
            return;
        }

        if (Integer.class == value.getClass()) {
            cell.setCellValue((Integer) value);
            this.currentColumnIndex++;
            return;
        }

        if (double.class == value.getClass()) {
            cell.setCellValue((double) value);
            this.currentColumnIndex++;
            return;
        }

        if (Double.class == value.getClass()) {
            cell.setCellValue((Double) value);
            this.currentColumnIndex++;
            return;
        }

        if (Date.class == value.getClass()) {
            cell.setCellValue((Date) value);
            this.currentColumnIndex++;
            return;
        }

        if (boolean.class == value.getClass()) {
            cell.setCellValue((boolean) value);
            this.currentColumnIndex++;
            return;
        }
        if (Boolean.class == value.getClass()) {
            cell.setCellValue((Boolean) value);
            this.currentColumnIndex++;
            return;
        }
        if (Calendar.class == value.getClass()) {
            cell.setCellValue((Calendar) value);
            this.currentColumnIndex++;
            return;
        }
        this.currentColumnIndex++;
    }

    /**
     * 设置某个元素的样式
     *
     * @param cell cell元素
     */
    private void setCellStyle(Cell cell, String styleKey) {
        if (null != styleKey && null != this.classifyStyle.get(styleKey)) {
            cell.setCellStyle(this.classifyStyle.get(styleKey));
            return;
        }

        if (null != this.appointLineStyle && this.appointLineStyle.containsKey(cell.getRowIndex())) {
            cell.setCellStyle(this.appointLineStyle.get(cell.getRowIndex()));
            return;
        }
        if (null != this.singleLineStyle && (cell.getRowIndex() % 2 != 0)) {
            cell.setCellStyle(this.singleLineStyle);
            return;
        }
        if (null != this.doubleLineStyle && (cell.getRowIndex() % 2 == 0)) {
            cell.setCellStyle(this.doubleLineStyle);
            return;
        }
        if (null != this.defaultStyle)
            cell.setCellStyle(this.defaultStyle);
    }
    /*************************************数据填充结束***********************************/

    /*************************************写出数据开始***********************************/

    /**
     * 将文件写到ExcelSink中
     *
     * @param excelSink
     */
    public void write(IExcelSink excelSink) {
        try {
            this.workbook.write(excelSink.getSink());
            excelSink.onCompleted();
        } catch (IOException e) {
            e.printStackTrace();
            excelSink.close();
            throw new RuntimeException("写入数据失败:" + e.getMessage());
        }
    }

	public CellStyle getCurrentStyle() {
		return currentStyle;
	}

	public void setCurrentStyle(CellStyle currentStyle) {
		this.currentStyle = currentStyle;
	}

    /*************************************写出数据结束***********************************/

}
