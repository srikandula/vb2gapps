package com.infy.gcoe.poi.base;

import static com.infy.gcoe.util.ReportConstants.SUMMARY_REPORT;

import java.awt.Color;
import java.io.File;
import java.io.FileOutputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.apache.poi.xssf.usermodel.XSSFDataValidationConstraint;
import org.apache.poi.xssf.usermodel.XSSFDataValidationHelper;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import com.infy.gcoe.util.ReportConstants;
import com.infy.gcoe.vo.ExcelReportVO;

/**
 * Creates a summary spread sheet with findings
 * 
 * @author srinivas.kandula
 *
 */
@Service
public class SummaryReportBuilder {
	
	private static Logger logger = LoggerFactory.getLogger(SummaryReportBuilder.class);

	public List<ExcelReportVO> update(List<ExcelReportVO> reportList) throws Exception {
		try{
			

			Date date = new Date();
			DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd-HH-mm");
			String archiveFile = "Archive-Summary-" + dateFormat.format(date) +".xlsx";
			
			File oldFile = new File(SUMMARY_REPORT);
			if(oldFile.exists()){
				logger.debug("Archive existing file as : " + archiveFile);
				oldFile.renameTo(new File(archiveFile));
			}
			
			
			XSSFWorkbook xssfWorkBook = new XSSFWorkbook();
			FileOutputStream fileOut = new FileOutputStream(new File(SUMMARY_REPORT));
			XSSFSheet summarySheet = xssfWorkBook.createSheet("summary");
			
			XSSFRow headerRow = summarySheet.createRow((short)0);			
			ExcelReportVO headerReport = new ExcelReportVO();
			Object headers[][] = headerReport.getData();
			
			XSSFCellStyle style = xssfWorkBook.createCellStyle();
			style.setFillBackgroundColor(new XSSFColor(Color.LIGHT_GRAY));
			style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			style.setFillForegroundColor(new XSSFColor(Color.GRAY));
			
			XSSFCell cell = headerRow.createCell(0);
			cell.setCellValue("#");
			cell.setCellStyle(style);
			for(int i=0;i<headers.length;i++){
				cell = headerRow.createCell(i+1);
				cell.setCellStyle(style);
				cell.setCellValue((String)headers[i][0]);				
			}
			
			int rowCounter = 0;
			XSSFRow row = null;
			int colCounter = 0;
			for(ExcelReportVO reportVO : reportList){
				row = summarySheet.createRow((short)(++rowCounter));
				
				Object data[][] = reportVO.getData();
				Object cellData = null;
				colCounter = data.length;
				row.createCell(0).setCellValue(rowCounter);
				for(int i=0;i<colCounter;i++){				
					cellData = data[i][1];
					if(cellData instanceof String){
						row.createCell(i+1).setCellValue((String)cellData);
					}else if(cellData instanceof Long){
						row.createCell(i+1).setCellValue((Long)cellData);
					}else if(cellData instanceof Boolean){
						row.createCell(i+1).setCellValue((Boolean)cellData);
					}else {
						row.createCell(i+1).setCellValue("-");
					}
				}
				
			}
			
			//Imposing constraint for last cell
			XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper(summarySheet);
			XSSFDataValidationConstraint dvConstraint = (XSSFDataValidationConstraint)dvHelper.createExplicitListConstraint(ReportConstants.ACTION_RESPONSE);
			CellRangeAddressList addressList = new CellRangeAddressList(1, rowCounter, colCounter, colCounter);
			XSSFDataValidation validation = (XSSFDataValidation)dvHelper.createValidation(dvConstraint, addressList);
			validation.setShowErrorBox(true);
			summarySheet.addValidationData(validation);
			
			xssfWorkBook.write(fileOut);
			xssfWorkBook.close();
			fileOut.close();
		}catch(Exception exception){
			logger.error("Error in creating summary",exception);
		}
		
		return reportList;
	}

}
