package com.github;

import java.lang.reflect.Array;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;
import java.util.Map;

import org.apache.commons.beanutils.BeanUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.github.data.Position;
import com.github.handler.ExcelHeader;
import com.github.handler.ExcelTemplate;
import com.github.sink.IExcelSink;
import com.github.source.IExcelSource;
import com.github.utils.Utils;

import moudles.Student1;

public class ExcelUtils {

    static private ExcelUtils excelUtils = new ExcelUtils();

    private ExcelUtils() {
    }

    public static ExcelUtils getInstance() {
        return excelUtils;
    }

    /*----------------------------------------读取Excel操作基于注解映射---------------------------------------------*/
    /*  一. 操作流程 ：                                                                                            */
    /*      1) 读取表头信息,与给出的Class类注解匹配                                                                  */
    /*      2) 读取表头下面的数据内容, 按行读取, 并映射至java对象                                                      */
    /*  二. 参数说明                                                                                               */
    /*      *) excelPath        =>      目标Excel路径                                                              */
    /*      *) excelSource      =>      用于获取Workbook, 支持继承, 参考ExcelFileSource和ExcelIostreamSource         */
    /*      *) clazz            =>      java映射对象                                                               */
    /*      *) offsetLine       =>      开始读取行坐标(默认0)                                                       */
    /*      *) limitLine        =>      最大读取行数(默认表尾)                                                      */
    /*      *) sheetIndex       =>      Sheet索引(默认0)                                                           */

    public <T> List<T> readExcel2Objects(IExcelSource excelSource, Class<T> clazz, int offsetLine, int limitLine, int
            sheetIndex) throws Exception {
        Workbook workbook = excelSource.getWorkBook();
        return readExcel2ObjectsHandler(workbook, clazz, offsetLine, limitLine, sheetIndex);
    }

    public <T> List<T> readExcel2Objects(IExcelSource excelSource, Class<T> clazz, int sheetIndex)
            throws Exception {
        return readExcel2Objects(excelSource, clazz, 0, Integer.MAX_VALUE, sheetIndex);
    }

    public <T> List<T> readExcel2Objects(IExcelSource excelSource, Class<T> clazz)
            throws Exception {
        return readExcel2Objects(excelSource, clazz, 0, Integer.MAX_VALUE, 0);
    }

    private <T> List<T> readExcel2ObjectsHandler(Workbook workbook, Class<T> clazz, int offsetLine, int limitLine,
                                                 int sheetIndex) throws Exception {
        Sheet sheet = workbook.getSheetAt(sheetIndex);
        Row row = sheet.getRow(offsetLine);
        List<T> list = new ArrayList<>();
        Map<Integer, ExcelHeader> maps = Utils.getHeaderMap(row, clazz);
        if (maps == null || maps.size() <= 0)
            throw new RuntimeException("要读取的Excel的格式不正确，检查是否设定了合适的行");
        int maxLine = sheet.getLastRowNum() > (offsetLine + limitLine) ? (offsetLine + limitLine) : sheet
                .getLastRowNum();
        for (int i = offsetLine + 1; i <= maxLine; i++) {
            row = sheet.getRow(i);
            T obj = clazz.newInstance();
            for (Cell cell : row) {
                int ci = cell.getColumnIndex();
                ExcelHeader header = maps.get(ci);
                if (null == header)
                    continue;
                String filed = header.getFiled();
                Utils.fixCellType(cell, header.getFiledClazz());
                String val = Utils.getCellValue(cell);
                Object value = Utils.str2TargetClass(val, header.getFiledClazz());
                BeanUtils.copyProperty(obj, filed, value);
            }
            list.add(obj);
        }
        return list;
    }

    /*----------------------------------------读取Excel操作无映射--------------------------------------------------*/
    /*  一. 操作流程 ：                                                                                            */
    /*      *) 按行读取Excel文件,存储形式为  Cell->String => Row->List<Cell> => Excel->List<Row>                    */
    /*  二. 参数说明                                                                                               */
    /*      *) excelPath        =>      目标Excel路径                                                              */
    /*      *) excelSource      =>      用于获取Workbook, 支持继承, 参考ExcelFileSource和ExcelIostreamSource         */
    /*      *) offsetLine       =>      开始读取行坐标(默认0)                                                       */
    /*      *) limitLine        =>      最大读取行数(默认表尾)                                                      */
    /*      *) sheetIndex       =>      Sheet索引(默认0)                                                           */

    public List<List<String>> readExcel2List(IExcelSource excelSource, int offsetLine, int limitLine, int sheetIndex)
            throws Exception {

        Workbook workbook = excelSource.getWorkBook();
        return readExcel2ObjectsHandler(workbook, offsetLine, limitLine, sheetIndex);
    }

    public List<List<String>> readExcel2List(IExcelSource excelSource, int offsetLine)
            throws Exception {

        Workbook workbook = excelSource.getWorkBook();
        return readExcel2ObjectsHandler(workbook, offsetLine, Integer.MAX_VALUE, 0);
    }

    public List<List<String>> readExcel2List(IExcelSource excelSource)
            throws Exception {

        Workbook workbook = excelSource.getWorkBook();
        return readExcel2ObjectsHandler(workbook, 0, Integer.MAX_VALUE, 0);
    }

    private List<List<String>> readExcel2ObjectsHandler(Workbook workbook, int offsetLine, int limitLine, int
            sheetIndex)
            throws Exception {

        List<List<String>> list = new ArrayList<>();
        Sheet sheet = workbook.getSheetAt(sheetIndex);
        int maxLine = sheet.getLastRowNum() > (offsetLine + limitLine) ? (offsetLine + limitLine) : sheet
                .getLastRowNum();
        for (int i = offsetLine; i <= maxLine; i++) {
            List<String> rows = new ArrayList<>();
            Row row = sheet.getRow(i);
            for (Cell cell : row) {
                String val = Utils.getCellValue(cell);
                rows.add(val);
            }
            list.add(rows);
        }
        return list;
    }


    /*--------------------------------------------基于模板、注解导出excel-------------------------------------------*/
    /*  一. 操作流程 ：                                                                                            */
    /*      1) 初始化模板                                                                                          */
    /*      2) 根据Java对象映射表头                                                                                 */
    /*      3) 写入数据内容                                                                                        */
    /*  二. 参数说明                                                                                               */
    /*      *) templateExcelSource =>      模板路径                                                                   */
    /*      *) sheetIndex       =>      Sheet索引(默认0)                                                           */
    /*      *) data             =>      导出内容List集合                                                            */
    /*      *) extendMap        =>      扩展内容Map(具体就是key匹配替换模板#key内容)                                  */
    /*      *) clazz            =>      映射对象Class                                                              */
    /*      *) isWriteHeader    =>      是否写入表头                                                               */
    /*      *) targetPath       =>      导出文件路径                                                               */
    /*      *) os               =>      导出文件流                                                                 */

    public void exportObjects2Excel(IExcelSource templateExcelSource, int sheetIndex,String[] sheetNames, List<?> data, Map<String, String> extendMap,
                                    Class<?> clazz, boolean isWriteHeader, IExcelSink excelSink) throws Exception {

        exportExcelByModuleHandler(templateExcelSource, sheetIndex, sheetNames, data, extendMap, clazz, isWriteHeader)
                .write(excelSink.getSink());
    }

    public void exportObjects2Excel(IExcelSource templateExcelSource, List<?> data, Map<String, String> extendMap, Class<?> clazz,
                                    boolean isWriteHeader, IExcelSink excelSink) throws Exception {

        exportObjects2Excel(templateExcelSource, 0,null, data, extendMap, clazz, isWriteHeader, excelSink);
    }

    public void exportObjects2Excel(IExcelSource templateExcelSource, List<?> data, Map<String, String> extendMap, Class<?> clazz,
    		IExcelSink excelSink) throws Exception {

        exportObjects2Excel(templateExcelSource, 0,null, data, extendMap, clazz, false, excelSink);
    }

    public void exportObjects2Excel(IExcelSource templateExcelSource, List<?> data, Class<?> clazz, IExcelSink excelSink)
            throws Exception {

        exportObjects2Excel(templateExcelSource, 0, null, data,null,  clazz, false, excelSink);
    }
    private Workbook exportExcelByModuleHandler(IExcelSource templateExcelSource, int sheetIndex, String[] sheetNames, List<?> data,
                                                     Map<String, String> extendMap, Class<?> clazz, boolean isWriteHeader)
            throws Exception {
    		ExcelTemplate templates = ExcelTemplate.getInstance(templateExcelSource, sheetIndex);
    		Workbook workbook = new XSSFWorkbook();
    		for(int i = 0; i < sheetNames.length; i++){
    			Sheet sheet = workbook.createSheet(sheetNames[i]);
        		CellStyle style = workbook.createCellStyle();
        		Map<String, Cell> extendData = templates.getExtendData(extendMap);
        		//获取extendMap里的数据
        	    Row[] rows = new Row[extendMap.size()];
        		Cell[] cells = templates.getExtendDataCell(extendMap, extendData);
        		String[] datalist  = templates.getExtendDataList(extendMap);
        		List<ExcelHeader> headers = Utils.getHeaderList(clazz);
        		//套用模板的前两行格式
        		Cell[] newCells = new Cell[extendMap.size()];
        		for(int h = 0; h < extendMap.size(); h++){
        			rows[h] = sheet.createRow(cells[h].getRow().getRowNum());
        			newCells[h] = rows[h].createCell(cells[h].getColumnIndex());
        		}
        		//合并单元格，并为合并后的单元格赋值
        		int sheetMergeCount = templates.getSheet().getNumMergedRegions();
        		for (int n = 0; n < sheetMergeCount; n++) {
        			// 得出具体的合并单元格
        			CellRangeAddress ca = templates.getSheet().getMergedRegion(n);
        			// 得到合并单元格的起始行, 结束行, 起始列, 结束列            
        			int firstCol = ca.getFirstColumn();            
        			int lastCol = ca.getLastColumn();        
        			int firstRow = ca.getFirstRow();      
        			int lastRow = ca.getLastRow();
        			sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
        			//设置列宽
        			for(int t = 0; t < templates.getSheet().getRow(n).getPhysicalNumberOfCells(); t++){
            			int a = templates.getSheet().getColumnWidth(t);
            			sheet.setColumnWidth(t, a);
            		}
        			//设置行高
        			for(int t = 0; t < templates.getSheet().getNumMergedRegions(); t++){
            			int a = templates.getSheet().getRow(t).getHeight();
            			sheet.getRow(t).setHeight((short) a);
            		}
        			Cell cell = null;
        			style = workbook.createCellStyle();
        			style.setBorderBottom(cells[n].getCellStyle().getBorderBottom());
    				style.setBorderLeft(cells[n].getCellStyle().getBorderLeft());
    				style.setBorderRight(cells[n].getCellStyle().getBorderRight());
    				style.setBorderTop(cells[n].getCellStyle().getBorderTop());
            		style.cloneStyleFrom((XSSFCellStyle) ((XSSFCellStyle) cells[n].getCellStyle()).clone());
        			for (int k = firstCol; k <= lastCol; k++) {
                        cell=rows[n].createCell(k);
                		cell.setCellValue("");  
                        cell.setCellStyle(style);
                       } 
        			
            		newCells[n].setCellStyle(style);
        			//为合并后的单元格赋值
        			newCells[n].setCellValue(datalist[n]);
        		}
        		Cell cell = null;
        		Row row = null;
        		//获取列名
        		String[] cellNames = new String[templates.getSheet().getRow(sheetMergeCount).getPhysicalNumberOfCells()];
        		cellNames = templates.getCellNames(sheetMergeCount);
        		row = sheet.createRow(sheetMergeCount);
        		//根据模板的格式在sheet中添加列名
    			for (int m = 0; m < headers.size()+1; m++) {
        			cell = row.createCell(m);
        			style = workbook.createCellStyle();
        			

      				XSSFCellStyle src = (XSSFCellStyle) (templates.getSheet().getRow(sheetMergeCount).getCell(m).getCellStyle());
      				style.setFillForegroundColor(src.getFillForegroundColor());
      				style.setFillBackgroundColor(src.getFillBackgroundColor());
      				style.cloneStyleFrom((CellStyle) src.clone());
      				style.setBorderBottom(src.getBorderBottom());
      				style.setBorderLeft(src.getBorderLeft());
      				style.setBorderRight(src.getBorderRight());
      				style.setBorderTop(src.getBorderTop());
            		cell.setCellStyle(style);
        			cell.setCellValue(cellNames[m]);
    			}
        		if (isWriteHeader) {
                    // 写标题
        			row = sheet.createRow(1);
        			for (int m = 0; m < headers.size()+1; m++) {
            			cell = row.createCell(m);
            			cell.setCellValue(1);
        			}
                }
        		//获取序号起始列
        		int serialNumberColumnIndex = templates.getSerialNumberColumnIndex();
        		//获取数据起始列和行
        		Position position = templates.getInitPosition();
        		int columnIndex = position.getColumn();
        		int rowIndex = position.getRow();
        		//按照模板的数据起始列和行等格式添加数据
        		Object _data;
            	for (int k = 0; k < data.size(); k++) {
            		row = sheet.createRow(k+rowIndex);
            		style = workbook.createCellStyle();
            		_data = data.get(k);
            		if(k % 2 == 0){
        				style.cloneStyleFrom((XSSFCellStyle) ((XSSFCellStyle) templates.getSingleLineStyle()).clone());
        			}
        			else{
        				style.cloneStyleFrom((XSSFCellStyle) ((XSSFCellStyle) templates.getDoubleLineStyle()).clone());
        			}
            		for (int m = 0; m < headers.size()+1; m++) {
            			cell = row.createCell(m);
            			cell.setCellStyle(style);
            			if(m == serialNumberColumnIndex)
            				cell.setCellValue(k+1);
            			else
            				cell.setCellValue(BeanUtils.getProperty(_data, headers.get(m-columnIndex).getFiled()));
            		}
            	}
    		}
        return workbook;
    	
    }

    /*---------------------------------------基于模板、注解导出Map数据----------------------------------------------*/
    /*  一. 操作流程 ：                                                                                            */
    /*      1) 初始化模板                                                                                          */
    /*      2) 根据Java对象映射表头                                                                                */
    /*      3) 写入数据内容                                                                                        */
    /*  二. 参数说明                                                                                               */
    /*      *) templateExcelSource =>      模板路径                                                                  */
    /*      *) sheetIndex       =>      Sheet索引(默认0)                                                          */
    /*      *) data             =>      导出内容Map集合                                                            */
    /*      *) extendMap        =>      扩展内容Map(具体就是key匹配替换模板#key内容)                                 */
    /*      *) clazz            =>      映射对象Class                                                             */
    /*      *) isWriteHeader    =>      是否写入表头                                                              */
    /*      *) targetPath       =>      导出文件路径                                                              */
    /*      *) os               =>      导出文件流                                                                */
    public void exportObject2Excel(IExcelSource templateExcelSource, int sheetIndex, String[] sheetNames, Map<String, List<?>> data,
                                   Map<String, String> extendMap, Class<?> clazz, boolean isWriteHeader, IExcelSink excelSink)
            throws Exception {

        exportExcelByModuleHandler(templateExcelSource, sheetIndex, sheetNames, data, extendMap, clazz, isWriteHeader)
                .write(excelSink.getSink());
    }

    public void exportObject2Excel(IExcelSource templateExcelSource, Map<String, List<?>> data, Map<String, String> extendMap,
                                   Class<?> clazz, IExcelSink excelSink) throws Exception {

        exportExcelByModuleHandler(templateExcelSource, 0, null, data, extendMap, clazz, false)
                .write(excelSink.getSink());
    }

    private Workbook exportExcelByModuleHandler(IExcelSource templateExcelSource, int sheetIndex, String[] sheetNames, Map<String, List<?>> data,
                                                     Map<String, String> extendMap, Class<?> clazz, boolean isWriteHeader)
            throws Exception {
    	ExcelTemplate templates = ExcelTemplate.getInstance(templateExcelSource, sheetIndex);
		Workbook workbook = new XSSFWorkbook();
		for(int i = 0; i < sheetNames.length; i++){
			Sheet sheet = workbook.createSheet(sheetNames[i]);
    		CellStyle style = workbook.createCellStyle();
    		Map<String, Cell> extendData = templates.getExtendData(extendMap);
    		//获取extendMap里的数据
    	    Row[] rows = new Row[extendMap.size()];
    		Cell[] cells = templates.getExtendDataCell(extendMap, extendData);
    		String[] datalist  = templates.getExtendDataList(extendMap);
    		//套用模板的前两行格式
    		Cell[] newCells = new Cell[extendMap.size()];
    		for(int h = 0; h < extendMap.size(); h++){
    			rows[h] = sheet.createRow(cells[h].getRow().getRowNum());
    			newCells[h] = rows[h].createCell(cells[h].getColumnIndex());
    		}
    		//合并单元格，并为合并后的单元格赋值
    		int sheetMergeCount = templates.getSheet().getNumMergedRegions();
    		for (int n = 0; n < sheetMergeCount; n++) {
    			// 得出具体的合并单元格
    			CellRangeAddress ca = templates.getSheet().getMergedRegion(n);
    			// 得到合并单元格的起始行, 结束行, 起始列, 结束列            
    			int firstCol = ca.getFirstColumn();            
    			int lastCol = ca.getLastColumn();        
    			int firstRow = ca.getFirstRow();      
    			int lastRow = ca.getLastRow();
    			sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
    			//设置列宽
    			for(int t = 0; t < templates.getSheet().getRow(n).getPhysicalNumberOfCells(); t++){
        			int a = templates.getSheet().getColumnWidth(t);
        			sheet.setColumnWidth(t, a);
        		}
    			//设置行高
    			for(int t = 0; t < templates.getSheet().getNumMergedRegions(); t++){
        			int a = templates.getSheet().getRow(t).getHeight();
        			sheet.getRow(t).setHeight((short) a);
        		}
    			Cell cell = null;
    			style = workbook.createCellStyle();
    			style.setBorderBottom(cells[n].getCellStyle().getBorderBottom());
				style.setBorderLeft(cells[n].getCellStyle().getBorderLeft());
				style.setBorderRight(cells[n].getCellStyle().getBorderRight());
				style.setBorderTop(cells[n].getCellStyle().getBorderTop());
        		style.cloneStyleFrom((XSSFCellStyle) ((XSSFCellStyle) cells[n].getCellStyle()).clone());
    			for (int k = firstCol; k <= lastCol; k++) {
                    cell=rows[n].createCell(k);
            		cell.setCellValue("");  
                    cell.setCellStyle(style);
                   } 
        		newCells[n].setCellStyle(style);
    			//为合并后的单元格赋值
    			newCells[n].setCellValue(datalist[n]);
    		}
    		Cell cell = null;
    		Row row = null;
    		//获取列名
    		String[] cellNames = new String[templates.getSheet().getRow(sheetMergeCount).getPhysicalNumberOfCells()];
    		cellNames = templates.getCellNames(sheetMergeCount);
    		row = sheet.createRow(sheetMergeCount);
    		List<ExcelHeader> headers = Utils.getHeaderList(clazz);
    		//根据模板的格式在sheet中添加列名
			for (int m = 0; m < headers.size()+1; m++) {
    			cell = row.createCell(m);
    			style = workbook.createCellStyle();
    			style.cloneStyleFrom((XSSFCellStyle) ((XSSFCellStyle) templates.getSheet().getRow(sheetMergeCount).getCell(m).getCellStyle()).clone());
    			style.setBorderBottom(templates.getSheet().getRow(sheetMergeCount).getCell(m).getCellStyle().getBorderBottom());
        		style.setBorderLeft(templates.getSheet().getRow(sheetMergeCount).getCell(m).getCellStyle().getBorderLeft());
        		style.setBorderRight(templates.getSheet().getRow(sheetMergeCount).getCell(m).getCellStyle().getBorderRight());
        		style.setBorderTop(templates.getSheet().getRow(sheetMergeCount).getCell(m).getCellStyle().getBorderTop());
        		cell.setCellStyle(style);
        		cell.setCellValue(cellNames[m]);
			}
    		if (isWriteHeader) {
                // 写标题
    			row = sheet.createRow(1);
    			for (int m = 0; m < headers.size()+1; m++) {
        			cell = row.createCell(m);
        			cell.setCellValue(1);
    			}
            }
    		//获取序号起始列
    		int serialNumberColumnIndex = templates.getSerialNumberColumnIndex();
    		//获取数据起始列和行
    		Position position = templates.getInitPosition();
    		int columnIndex = position.getColumn();
    		int rowIndex = position.getRow();
    		//按照模板的数据起始列和行等格式添加数据
    		int currentColumnIndex = serialNumberColumnIndex+1;
    		for (Map.Entry<String, List<?>> entry : data.entrySet()) {
   			 for (Object object : entry.getValue()) {
   				 row = sheet.createRow(rowIndex++);
   				 style = workbook.createCellStyle();
   				 XSSFCellStyle src = (XSSFCellStyle) templates.getCellStyle(entry.getKey());
   				 style.setFillForegroundColor(src.getFillForegroundColor());
   				 style.setFillBackgroundColor(src.getFillBackgroundColor());
   				 style.cloneStyleFrom((CellStyle) src.clone());
   				 style.setBorderBottom(src.getBorderBottom());
				 style.setBorderLeft(src.getBorderLeft());
				 style.setBorderRight(src.getBorderRight());
				 style.setBorderTop(src.getBorderTop());
   				 for (int m = 0; m < headers.size()+1; m++) {
   					 cell = row.createCell(m);
   					 cell.setCellStyle(style);
   					 if(m == serialNumberColumnIndex)
   						 cell.setCellValue(currentColumnIndex);
   					 else
   						 cell.setCellValue(BeanUtils.getProperty(object, headers.get(m-columnIndex).getFiled()));
   				 }
   				currentColumnIndex++;
   			 }
   		 }
		}
    return workbook;
    }

    /*----------------------------------------无模板基于注解导出---------------------------------------------------*/
    /*  一. 操作流程 ：                                                                                            */
    /*      1) 根据Java对象映射表头                                                                     */
    /*      2) 写入数据内容                                                                                    */
    /*  二. 参数说明                                                                                                 */
    /*      *) data             =>      导出内容List集合                         */
    /*      *) isWriteHeader    =>      是否写入表头                                */
    /*      *) sheetName        =>      Sheet索引名(默认0)          */
    /*      *) clazz            =>      映射对象Class               */
    /*      *) isXSSF           =>      是否Excel2007以上                      */
    /*      *) targetPath       =>      导出文件路径                                 */
    /*      *) os               =>      导出文件流                                     */
    public void exportObjects2Excel(List<?> data, Class<?> clazz, boolean isWriteHeader, String[] sheetNames, boolean isXSSF,
    		IExcelSink excelSink) throws Exception {
        exportExcelNoModuleHandler(data,clazz, isWriteHeader, sheetNames, isXSSF).write(excelSink.getSink());
        excelSink.onCompleted().close();
    }

    public void exportObjects2Excel(List<?> data, Class<?> clazz, boolean isWriteHeader, IExcelSink excelSink)
            throws Exception {

        exportExcelNoModuleHandler(data, clazz, isWriteHeader, null, true).write(excelSink.getSink());
        excelSink.onCompleted().close();
    }

    private Workbook exportExcelNoModuleHandler(List<?> data,Class<?> clazz, boolean isWriteHeader, String[] sheetNames,
                                                boolean isXSSF) throws Exception {

        Workbook workbook;
        if (isXSSF) {
            workbook = new XSSFWorkbook();
        } else {
            workbook = new HSSFWorkbook();
        }
        
        Sheet sheet;
        for(int i = 0; i < sheetNames.length; i++){
        	if (null != sheetNames[i] && !"".equals(sheetNames[i])) {
                sheet = workbook.createSheet(sheetNames[i]);
                Row row = sheet.createRow(0);
                List<ExcelHeader> headers = Utils.getHeaderList(clazz);
                if (isWriteHeader) {
                    // 写标题
                    for (int j = 0; j < headers.size(); j++) {
                        row.createCell(j).setCellValue(headers.get(j).getTitle());
                    }
                }
             // 写数据
                List<Student1> list = new ArrayList<>();
                list = (List<Student1>) data.get(i);
                Object _data;
                	for (int k = 0; k < list.size(); k++) {
                		row = sheet.createRow(k + 1);
                		_data = list.get(k);
                		for (int m = 0; m < headers.size(); m++) {
                			Cell cell = row.createCell(m);
                			
                			XSSFCellStyle ztStyle = (XSSFCellStyle) workbook.createCellStyle();     
                            // 创建字体对象     
                            cell.setCellValue(BeanUtils.getProperty(_data, headers.get(m).getFiled()));
                			
                		}
                    }
                
            } else {
                sheet = workbook.createSheet();
            }
        }
        return workbook;
    }

    /*-----------------------------------------无模板无注解导出----------------------------------------------------*/
    /*  一. 操作流程 ：                                                                                           */
    /*      1) 写入表头内容(可选)                                                                                  */
    /*      2) 写入数据内容                                                                                       */
    /*  二. 参数说明                                                                                              */
    /*      *) data             =>      导出内容List集合                                                          */
    /*      *) header           =>      表头集合,有则写,无则不写                                                   */
    /*      *) sheetName        =>      Sheet索引名(默认0)                                                        */
    
    /*      *) isXSSF           =>      是否Excel2007以上                                                         */
    /*      *) targetPath       =>      导出文件路径                                                              */
    /*      *) os               =>      导出文件流                                                                */

    public void exportObjects2Excel(List<?> data, List<?> header, String[] sheetNames, boolean isXSSF, IExcelSink excelSink) throws Exception {
        exportExcelNoModuleHandler(data, header, sheetNames, isXSSF).write(excelSink.getSink());
		excelSink.onCompleted().close();
    }

    public void exportObjects2Excel(List<?> data, List<?> header, IExcelSink excelSink) throws Exception {

        exportExcelNoModuleHandler(data, header,null, true)
                .write(excelSink.getSink());
        excelSink.onCompleted().close();
    }

    public void exportObjects2Excel(List<?> data, IExcelSink excelSink) throws Exception {

        exportExcelNoModuleHandler(data, null,null, true)
                .write(excelSink.getSink());
        excelSink.onCompleted().close();
    }

    private Workbook exportExcelNoModuleHandler(List<?> data, List<?> header, String[] sheetNames, boolean isXSSF)
            throws Exception {
    	 Workbook workbook;
         if (isXSSF) {
             workbook = new XSSFWorkbook();
         } else {
             workbook = new HSSFWorkbook();
         }
         Sheet sheet;
         for(int k = 0; k < sheetNames.length; k++){
         	if (null != sheetNames[k] && !"".equals(sheetNames[k])) {
                 sheet = workbook.createSheet(sheetNames[k]);
             } else {
                 sheet = workbook.createSheet();
             }
         	int rowIndex = 0;
         	List<String> head = new ArrayList<>();
         	head = (List<String>) header.get(k);
             if (null != head && head.size() > 0) {
                 // 写标题
                 Row row = sheet.createRow(rowIndex);
                 for (int i = 0; i < head.size(); i++) {
                     row.createCell(i, Cell.CELL_TYPE_STRING).setCellValue(head.get(i));
                 }
                 rowIndex++;
             }
            List<List<String>> datalist = new ArrayList<>();
            datalist =  (List<List<String>>) data.get(k);
             for (Object object : datalist) {
                 Row row = sheet.createRow(rowIndex);
                 if (object.getClass().isArray()) {
                     for (int j = 0; j < Array.getLength(object); j++) {
                         row.createCell(j, Cell.CELL_TYPE_STRING).setCellValue(Array.get(object, j).toString());
                     }
                 } else if (object instanceof Collection) {
                     Collection<?> items = (Collection<?>) object;
                     int m = 0;
                     for (Object item : items) {
                         row.createCell(m, Cell.CELL_TYPE_STRING).setCellValue(item.toString());
                         m++;
                     }
                 } else {
                     row.createCell(0, Cell.CELL_TYPE_STRING).setCellValue(object.toString());
                 }
                 rowIndex++;
             }
         }
         return workbook;
    }
}
