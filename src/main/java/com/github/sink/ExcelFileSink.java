package com.github.sink;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

public class ExcelFileSink implements IExcelSink {
	
	@Override
	public IExcelSink onCompleted() {
		
		System.out.println(outputStream);
		
		return IExcelSink.super.onCompleted();
	}

	private String path;
	
	private OutputStream outputStream;
	
	private ExcelFileSink(){}

	/**
	 * 创建
	 * @param excelPath
	 * @return
	 * @throws Exception
	 */
	public static ExcelFileSink create(String excelPath) throws Exception{
		ExcelFileSink excelFileSink = new ExcelFileSink();
		excelFileSink.path = excelPath;
		return excelFileSink;
	}

	@Override
	public OutputStream getSink() {
		if(outputStream != null){
			return outputStream;
		}
		
		try {
			outputStream = new FileOutputStream(path);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
			close();
			throw new RuntimeException("写入的文件不存在");
		}
		return outputStream;
	}

	@Override
	public void close() {
		try {
            if (outputStream != null)
            	outputStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
	}
}