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
	private long totalRowCount           = 0l;
	private long numberOfSheets          = 0l;
	
	private List<ExcelMacroVO> macroList = new ArrayList<>();
	private long macroLinesOfCode        = 0l;
	private long numberOfMacros          = 0l;
	
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
		incrementNumberOfSheets();
		incrementTotalRowCount(sheet.getRowCount());
	}

	public List<ExcelMacroVO> getMacroList() {
		return macroList;
	}

	public void setMacroList(List<ExcelMacroVO> macroList) {
		this.macroList = macroList;
	}
	
	public void addMacro(ExcelMacroVO macro) {
		if(this.macroList == null){
			macroList = new ArrayList<>();
		}
		this.macroList.add(macro);
		incrementNumberOfMacros();
		incrementMacroLinesOfCode(macro.getLineCount());
	}
	
	public long getTotalRowCount() {
		return totalRowCount;
	}

	public void incrementTotalRowCount(long rowCount) {
		this.totalRowCount += rowCount;
	}

	public long getNumberOfSheets() {
		return numberOfSheets;
	}

	public void incrementNumberOfSheets() {
		this.numberOfSheets += 1;
	}

	public long getMacroLinesOfCode() {
		return macroLinesOfCode;
	}

	public void incrementMacroLinesOfCode(long macroLinesOfCode) {
		this.macroLinesOfCode += macroLinesOfCode;
	}

	public long getNumberOfMacros() {
		return numberOfMacros;
	}

	public void setNumberOfMacros(long numberOfMacros) {
		this.numberOfMacros = numberOfMacros;
	}
	
	public void incrementNumberOfMacros() {
		this.numberOfMacros += 1;
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
		builder.append("ReportVO [fileName=");
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
		builder.append(numberOfSheets);
		builder.append(", Total Rows=");
		builder.append(totalRowCount);
		builder.append(", No Of Macros=");
		builder.append(numberOfMacros);
		builder.append(", Macros LOC=");
		builder.append(macroLinesOfCode);
		
		if(logger.isDebugEnabled()){
			builder.append(", absolutePath=");
			builder.append(absolutePath);
		}
				
		if(logger.isDebugEnabled() && sheetList != null){
			builder.append("\n");
			builder.append(",[ sheetList=");
			builder.append(sheetList.toString());
			builder.append(" ]");
			builder.append("\n");
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
