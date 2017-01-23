package com.infy.gcoe.vo;

public class ExcelSheetVO {
	
	private String name;
	private int sheetSequence;
	private long rowCount;
	private long colCount;
	
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}	
	public int getSheetSequence() {
		return sheetSequence;
	}
	public void setSheetSequence(int sheetSequence) {
		this.sheetSequence = sheetSequence;
	}
	public long getRowCount() {
		return rowCount;
	}
	public void setRowCount(long rowCount) {
		this.rowCount = rowCount;
	}
	public long getColCount() {
		return colCount;
	}
	public void setColCount(long colCount) {
		this.colCount = colCount;
	}
	@Override
	public String toString() {
		StringBuilder builder = new StringBuilder();
		builder.append("\n\t");
		builder.append("ExcelSheetVO: [");
		builder.append("name=");
		builder.append(name);
		builder.append(", sheetSequence=");
		builder.append(sheetSequence);
		builder.append(", rowCount=");
		builder.append(rowCount);
		builder.append(", colCount=");
		builder.append(colCount);
		builder.append("]");
		
		return builder.toString();
	}
		
}
