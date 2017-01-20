package com.infy.gcoe.poi.vo;

import static com.infy.gcoe.poi.base.ReportConstants.ACTION_RESPONSE;
import static com.infy.gcoe.poi.base.ReportConstants.ACTION_RESP_NO_ACTION;
import static com.infy.gcoe.poi.base.ReportConstants.ACTION_TYPE;
import static com.infy.gcoe.poi.base.ReportConstants.CREATED_BY;
import static com.infy.gcoe.poi.base.ReportConstants.FILE_EXTENSION;
import static com.infy.gcoe.poi.base.ReportConstants.FILE_NAME;
import static com.infy.gcoe.poi.base.ReportConstants.FILE_SIZE;
import static com.infy.gcoe.poi.base.ReportConstants.FULL_FILE_NAME;
import static com.infy.gcoe.poi.base.ReportConstants.IS_2003_FORMAT;
import static com.infy.gcoe.poi.base.ReportConstants.IS_EXCEL_FILE;
import static com.infy.gcoe.poi.base.ReportConstants.LAST_MODIFIED_BY;
import static com.infy.gcoe.poi.base.ReportConstants.MACROS;
import static com.infy.gcoe.poi.base.ReportConstants.MACROS_LOC;
import static com.infy.gcoe.poi.base.ReportConstants.NO_OF_EMBEDDS;
import static com.infy.gcoe.poi.base.ReportConstants.NO_OF_MACROS;
import static com.infy.gcoe.poi.base.ReportConstants.NO_OF_PICTURES;
import static com.infy.gcoe.poi.base.ReportConstants.NO_OF_SHEETS;
import static com.infy.gcoe.poi.base.ReportConstants.PIVOT_COUNT;
import static com.infy.gcoe.poi.base.ReportConstants.TOTAL_ROWS;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.infy.gcoe.poi.base.ReportConstants;
import com.infy.gcoe.poi.base.ReportConstants.*;
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
	
	private long noOfPivotTables;
	private long noOfEmbedds;
	private long noOfPictures;
	
	
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
		this.setHasMacros(true);
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
	

	public long getNoOfPivotTables() {
		return noOfPivotTables;
	}

	public void setNoOfPivotTables(long noOfPivotTables) {
		this.noOfPivotTables = noOfPivotTables;
	}

	public long getNoOfEmbedds() {
		return noOfEmbedds;
	}

	public void setNoOfEmbedds(long noOfEmbedds) {
		this.noOfEmbedds = noOfEmbedds;
	}

	public long getNoOfPictures() {
		return noOfPictures;
	}

	public void setNoOfPictures(long noOfPictures) {
		this.noOfPictures = noOfPictures;
	}

	public boolean isHasMacros() {
		return hasMacros;
	}

	public void setExcelFile(boolean isExcelFile) {
		this.isExcelFile = isExcelFile;
	}

	public void setTotalRowCount(long totalRowCount) {
		this.totalRowCount = totalRowCount;
	}

	public void setNumberOfSheets(long numberOfSheets) {
		this.numberOfSheets = numberOfSheets;
	}

	public void setMacroLinesOfCode(long macroLinesOfCode) {
		this.macroLinesOfCode = macroLinesOfCode;
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
		builder.append(", noOfPivotTables=");
		builder.append(noOfPivotTables);
		builder.append(", noOfEmbedds=");
		builder.append(noOfEmbedds);
		builder.append(", noOfPictures=");
		builder.append(noOfPictures);
		
		if(!logger.isDebugEnabled()){
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
	
	/**
	 * Prepare 2x2 matrix to populate the excel sheet
	 * 
	 * @return
	 */
	public Object[][] getData() {
		Object header[][] = new Object[][] { 
				{ FILE_NAME, getFileName() }, 
				{ FULL_FILE_NAME, getAbsolutePath() },
				{ CREATED_BY, getCreatedBy() },
				{ LAST_MODIFIED_BY, getLastModifiedBy() }, 
				{ FILE_EXTENSION, getFileExtension() },
				{ IS_2003_FORMAT,  isOldFormat() }, 
				{ MACROS, isHasMacros() },
				{ IS_EXCEL_FILE,  isExcelFile() }, 
				{ FILE_SIZE,  getSize() },
				{ NO_OF_SHEETS,  getNumberOfSheets() }, 
				{ TOTAL_ROWS,  getTotalRowCount() },
				{ NO_OF_MACROS, getNumberOfMacros() }, 
				{ MACROS_LOC,  getMacroLinesOfCode() },
				{ PIVOT_COUNT,  getNoOfPivotTables() }, 
				{ NO_OF_EMBEDDS,  getNoOfEmbedds() },
				{ NO_OF_PICTURES,  getNoOfPictures() } ,
				{ ACTION_TYPE,  ACTION_RESP_NO_ACTION } 
		};

		return header;
	}
	
	/**
	 * Recreate the excel VO
	 * @param headerName
	 * @param value
	 * @return
	 */
	public ExcelReportVO setData(String headerName, String value) {
		
		switch(headerName){
		case FILE_NAME:
			setFileName(value);
			break;
		case FULL_FILE_NAME:
			setAbsolutePath(value);
			break;
		case CREATED_BY:
			setCreatedBy(value);
			break;
		case LAST_MODIFIED_BY:
			setLastModifiedBy(value);
			break;
		case FILE_EXTENSION:
			setFileExtension(value);
			break;
		}		
		return this;
	}
	
	/**
	 * Recreate the excel VO
	 * @param headerName
	 * @param value
	 * @return
	 */
	public ExcelReportVO setData(String headerName, Boolean value) {
		
		switch(headerName){		
		case IS_2003_FORMAT:
			setOldFormat(value);
			break;
		case MACROS:
			setHasMacros(value);
			break;
		case IS_EXCEL_FILE:
			setExcelFile(value);
			break;
		}		
		return this;
	}
	
	/**
	 * Recreate the excel VO
	 * @param headerName
	 * @param value
	 * @return
	 */
	public ExcelReportVO setData(String headerName, Long value) {
		
		switch(headerName){		
		case FILE_SIZE:
			setSize(value);
			break;
		case NO_OF_SHEETS:
			setNumberOfSheets(value);
			break;
		case TOTAL_ROWS:
			setTotalRowCount(value);
			break;
		case NO_OF_MACROS:
			setNumberOfMacros(value);
			break;
		case MACROS_LOC:
			setMacroLinesOfCode(value);
			break;
		case PIVOT_COUNT:
			setNoOfPivotTables(value);
			break;
		case NO_OF_EMBEDDS:
			setNoOfEmbedds(value);
			break;
		case NO_OF_PICTURES:
			setNoOfPictures(value);
			break;
		}
		
		return this;
	}
	
	
}
