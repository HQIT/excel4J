package com.github.data;

public class Position {
	
	public Position(){}
	
	public Position(int row, int column){
		this.row = row;
		this.column = column;
	}
	
	private int column;
	private int row;
	
	public int getColumn() {
		return column;
	}

	public void setColumn(int column) {
		this.column = column;
	}

	public int getRow() {
		return row;
	}

	public void setRow(int row) {
		this.row = row;
	}

}
