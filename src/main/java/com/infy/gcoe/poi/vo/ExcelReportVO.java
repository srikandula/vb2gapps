package com.infy.gcoe.poi.vo;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

public class ExcelReportVO {
	
	private File file;
	
	private String fileName;
	private String absolutePath;
	private String createdBy;
	private String lastModifiedBy;
	
	private long size;
	
	public ExcelReportVO(){
		
	}
	
	public ExcelReportVO(File file, String fileName, String absolutePath, long size) {
		super();
		this.file = file;
		this.fileName = fileName;
		this.absolutePath = absolutePath;
		this.size = size;
	}

	private List<ExcelMacroVO> macroList = new ArrayList<>();

	public File getFile() {
		return file;
	}

	public void setFile(File file) {
		this.file = file;
	}
	
	public String getFileName() {
		return fileName;
	}

	public void setFileName(String fileName) {
		this.fileName = fileName;
	}

	public String getAbsolutePath() {
		return absolutePath;
	}

	public void setAbsolutePath(String absolutePath) {
		this.absolutePath = absolutePath;
	}

	public String getCreatedBy() {
		return createdBy;
	}

	public void setCreatedBy(String createdBy) {
		this.createdBy = createdBy;
	}

	public String getLastModifiedBy() {
		return lastModifiedBy;
	}

	public void setLastModifiedBy(String lastModifiedBy) {
		this.lastModifiedBy = lastModifiedBy;
	}

	public long getSize() {
		return size;
	}

	public void setSize(long size) {
		this.size = size;
	}

	public List<ExcelMacroVO> getMacroList() {
		return macroList;
	}

	public void setMacroList(List<ExcelMacroVO> macroList) {
		this.macroList = macroList;
	}

	@Override
	public String toString() {
		StringBuilder builder = new StringBuilder();
		builder.append("ExcelReportVO [fileName=");
		builder.append(fileName);
		builder.append(", absolutePath=");
		builder.append(absolutePath);
		builder.append(", createdBy=");
		builder.append(createdBy);
		builder.append(", lastModifiedBy=");
		builder.append(lastModifiedBy);
		builder.append(", size=");
		builder.append(size);
		builder.append("\n");
		if(macroList != null){
			builder.append(",[ macroList=");
			builder.append(macroList.toString());
			builder.append(" ]");
		}
		builder.append("]");
		builder.append("\n");
		return builder.toString();
	}
	
	
}
