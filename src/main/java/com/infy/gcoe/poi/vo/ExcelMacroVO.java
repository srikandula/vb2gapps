package com.infy.gcoe.poi.vo;

public class ExcelMacroVO {
	
	private String name;
	private String content;
	private long lineCount;
	
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}
	public String getContent() {
		return content;
	}
	public void setContent(String content) {
		this.content = content;
	}
	public long getLineCount() {
		return lineCount;
	}
	public void setLineCount(long lineCount) {
		this.lineCount = lineCount;
	}
	
	
	@Override
	public String toString() {
		StringBuilder builder = new StringBuilder();
		builder.append("\n\t");
		builder.append("ExcelMacroVO [name=");
		builder.append(name);
		//builder.append(", content=");
		//builder.append(content);
		builder.append(", lineCount=");
		builder.append(lineCount);
		builder.append("]");
		
		return builder.toString();
	}
		
}
