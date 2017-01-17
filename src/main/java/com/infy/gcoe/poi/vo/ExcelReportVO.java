package com.infy.gcoe.poi.vo;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.infy.gcoe.poi.GenerateReport;

public class ExcelReportVO {
	
	private static Logger logger = LoggerFactory.getLogger(ExcelReportVO.class);
	
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
		builder.append("\n");
		builder.append("ExcelReportVO [fileName=");
		builder.append(fileName);		
		builder.append(", createdBy=");
		builder.append(createdBy);
		builder.append(", lastModifiedBy=");
		builder.append(lastModifiedBy);
		builder.append(", size=");
		builder.append(size);
		builder.append(", No Of Macros=");
		builder.append(macroList.size());
		
		if(logger.isDebugEnabled()){
			builder.append(", absolutePath=");
			builder.append(absolutePath);
		}
				
		if(logger.isDebugEnabled() && macroList != null){
			builder.append("\n");
			builder.append(",[ macroList=");
			builder.append(macroList.toString());
			builder.append(" ]");
			builder.append("\n");
		}
		builder.append("]");
		builder.append("\n");
		return builder.toString();
	}
	
	
}
