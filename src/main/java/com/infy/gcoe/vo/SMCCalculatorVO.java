package com.infy.gcoe.vo;

public class SMCCalculatorVO {
	
	private String name;
	private String complexity;
	private long lineCount;
	
	
	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public String getComplexity() {
		return complexity;
	}

	public void setComplexity(String complexity) {
		this.complexity = complexity;
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
		builder.append("SMCCalculatorVO : [FileName=");
		builder.append(name);
		builder.append(", complexity=");
		builder.append(complexity);
		//builder.append(", content=");
		//builder.append(content);
		builder.append(", lineCount=");
		builder.append(lineCount);
		builder.append("]");
		
		return builder.toString();
	}
		
}
