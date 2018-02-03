package com.github.data;

public class SheetInfo {

    public SheetInfo(){}

    public SheetInfo(int index){
        this.index = index;
    }

    public SheetInfo(int offsetLine, int limitLine, int index){
        this.offsetLine = offsetLine;
        this.limitLine = limitLine;
        this.index = index;
    }

    /**
     * 当前Sheet的名字
     */
    private String name;

    /**
     * Sheet的Index
     * 默认为: 0
     */
    private int index = 0;

    /**
     * Sheet的数量
     * 默认为: 1
     */
    private int number;

    /**
     * 偏移行
     * 默认为: 0
     */
    private int offsetLine = 0;

    /**
     * 限制行
     * 默认为: Integer::MAX_VALUE
     */
    private int limitLine = Integer.MAX_VALUE;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public int getIndex() {
        return index;
    }

    public void setIndex(int index) {
        this.index = index;
    }

    public int getNumber() {
        return number;
    }

    public void setNumber(int number) {
        this.number = number;
    }

    public int getOffsetLine() {
        return offsetLine;
    }

    public void setOffsetLine(int offsetLine) {
        this.offsetLine = offsetLine;
    }

    public int getLimitLine() {
        return limitLine;
    }

    public void setLimitLine(int limitLine) {
        this.limitLine = limitLine;
    }
}
