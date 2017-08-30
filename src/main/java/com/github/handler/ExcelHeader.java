package com.github.handler;

/**
 * 功能说明: 用来存储Excel标题的对象，通过该对象可以获取标题和方法的对应关系
 */
public class ExcelHeader implements Comparable<ExcelHeader> {
    /**
     * excel的标题名称
     */
    private String title;
    /**
     * 每一个标题的顺序
     */
    private int order;
    /**
     * 注解域
     */
    private String filed;
    /**
     * 属性类型
     */
    private Class<?> filedClazz;
    /**
     * 该标题是否必须存在
     */
    private boolean required = false;

    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public int getOrder() {
        return order;
    }

    public void setOrder(int order) {
        this.order = order;
    }

    public String getFiled() {
        return filed;
    }

    public void setFiled(String filed) {
        this.filed = filed;
    }

    public Class<?> getFiledClazz() {
        return filedClazz;
    }

    public void setFiledClazz(Class<?> filedClazz) {
        this.filedClazz = filedClazz;
    }

    public int compareTo(ExcelHeader o) {
        return order - o.order;
    }

    public ExcelHeader(String title, int order, String filed, Class<?> filedClazz) {
        super();
        this.title = title;
        this.order = order;
        this.filed = filed;
        this.filedClazz = filedClazz;
    }
    
    public ExcelHeader(String title, int order, boolean required, String filed, Class<?> filedClazz) {
        super();
        this.title = title;
        this.order = order;
        this.filed = filed;
        this.filedClazz = filedClazz;
        this.required = required;
    }

	public boolean isRequired() {
		return required;
	}

	public void setRequired(boolean required) {
		this.required = required;
	}

}
