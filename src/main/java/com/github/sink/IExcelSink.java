package com.github.sink;

import java.io.OutputStream;

public interface IExcelSink {
	
	/**
	 * 获取到输出的OutputStream
	 * @return
	 */
	OutputStream getSink();
	
	/**
	 * 写完毕后进行调用, 关闭之前
	 */
	default IExcelSink onCompleted(){return this;};
	
	/**
	 * 关闭Sink
	 */
	void close();
}