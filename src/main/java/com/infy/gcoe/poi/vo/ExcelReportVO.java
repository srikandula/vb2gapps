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
	
	private boolean isExcelFile;
	private boolean isOldFormat;
	private String fileExtension;
	
	private boolean hasMacros;
	
	private long size;
	
	private List<ExcelSheetVO> sheetList = new ArrayList<>();
	private List<ExcelMacroVO> macroList = new ArrayList<>();
	
	public ExcelReportVO(){
		
	}
	
	public ExcelReportVO(File file, String fileName, String absolutePath, long size) {
		super();
		this.file = file;
		this.fileName = fileName;
		this.absolutePath = absolutePath;
		this.size = size;
	}

	

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

	public List<ExcelSheetVO> getSheetList() {
		return sheetList;
	}

	public void setSheetList(List<ExcelSheetVO> sheetList) {
		this.sheetList = sheetList;
	}
	
	public void addSheet(ExcelSheetVO sheet) {
		if(this.sheetList == null){
			sheetList = new ArrayList<>();
		}
		this.sheetList.add(sheet);
	}

	public List<ExcelMacroVO> getMacroList() {
		return macroList;
	}

	public void setMacroList(List<ExcelMacroVO> macroList) {
		this.macroList = macroList;
	}
	
	public boolean isExcelFile() {
		return isExcelFile;
	}

	public void isExcelFile(boolean isExcelFile) {
		this.isExcelFile = isExcelFile;
	}
	
	public boolean isOldFormat() {
		return isOldFormat;
	}

	public void setOldFormat(boolean isOldFormat) {
		this.isOldFormat = isOldFormat;
	}

	public String getFileExtension() {
		return fileExtension;
	}

	public void setFileExtension(String fileExtension) {
		this.fileExtension = fileExtension;
	}

	public boolean hasMacros() {
		return hasMacros;
	}

	public void setHasMacros(boolean hasMacros) {
		this.hasMacros = hasMacros;
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
		builder.append(", fileExtension=");
		builder.append(fileExtension);
		builder.append(", isOldFormat=");
		builder.append(isOldFormat);
		builder.append(", hasMacros=");
		builder.append(hasMacros);
		builder.append(", isExcelFile=");
		builder.append(isExcelFile);
		builder.append(", size=");
		builder.append(size);
		builder.append(", No Of Sheets=");
		builder.append(sheetList.size());
		builder.append(", No Of Macros=");
		builder.append(macroList.size());
		
		if(logger.isDebugEnabled()){
			builder.append(", absolutePath=");
			builder.append(absolutePath);
		}
				
		if(sheetList != null){
			builder.append("\n");
			builder.append(",[ sheetList=");
			builder.append(sheetList.toString());
			builder.append(" ]");
			builder.append("\n");
		}
		
		if(macroList != null){
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
