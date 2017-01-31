package com.infy.gcoe.vo;

public class SMCCalculatorVO {
	
	private String name;
	private String complexity;
	private long lineCount;
	private long hitCount;
	private long missedCount;
	private long customFunctionsCount;
	
	public long getCustomFunctionsCount() {
		return customFunctionsCount;
	}

	public void setCustomFunctionsCount(long customFunctionsCount) {
		this.customFunctionsCount = customFunctionsCount;
	}

	public long getHitCount() {
		return hitCount;
	}

	public void setHitCount(long hitCount) {
		this.hitCount = hitCount;
	}

	public long getMissedCount() {
		return missedCount;
	}

	public void setMissedCount(long missedCount) {
		this.missedCount = missedCount;
	}

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
		builder.append(", lineCount=");
		builder.append(lineCount);
		builder.append(", hitCount=");
		builder.append(hitCount);
		builder.append(", missedCount=");
		builder.append(missedCount);
		builder.append(", customFunctionsCount=");
		builder.append(customFunctionsCount);
		builder.append("]");
		
		return builder.toString();
	}
		
}
