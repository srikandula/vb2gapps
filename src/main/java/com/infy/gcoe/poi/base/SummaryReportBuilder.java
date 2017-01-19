package com.infy.gcoe.poi.base;

import java.awt.Color;
import java.io.File;
import java.io.FileOutputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import com.infy.gcoe.poi.vo.ExcelReportVO;

/**
 * Creates a summary spread sheet with findings
 * 
 * @author srinivas.kandula
 *
 */
@Service
public class SummaryReportBuilder implements IReportBuilder {
	
	private static Logger logger = LoggerFactory.getLogger(SummaryReportBuilder.class);

	@Override
	public List<ExcelReportVO> update(List<ExcelReportVO> reportList) throws Exception {
		try{
			logger.debug("Read folders " + reportList);
			
			Date date = new Date();
			DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd-HH-mm");
			String summaryReportFileName = "summary-" + dateFormat.format(date) +".xlsx";
			
			XSSFWorkbook wb = new XSSFWorkbook();
			FileOutputStream fileOut = new FileOutputStream(new File(summaryReportFileName));
			XSSFSheet summarySheet = wb.createSheet("summary");
			
			XSSFRow headerRow = summarySheet.createRow((short)0);			
			ExcelReportVO headerReport = new ExcelReportVO();
			Object headers[][] = headerReport.getData();
			headerRow.createCell(0).setCellValue("#");
			XSSFCellStyle style = wb.createCellStyle();
			style.setFillBackgroundColor(new XSSFColor(Color.GRAY));
			for(int i=0;i<headers.length;i++){
				XSSFCell cell = headerRow.createCell(i+1);
				cell.setCellStyle(style);
				cell.setCellValue((String)headers[i][0]);				
			}
			
			int rowCounter = 0;
			XSSFRow row = null;
			for(ExcelReportVO reportVO : reportList){
				row = summarySheet.createRow((short)(++rowCounter));
				
				Object data[][] = reportVO.getData();
				Object cellData = null;
				row.createCell(0).setCellValue(rowCounter);
				for(int i=0;i<data.length;i++){				
					cellData = data[i][1];
					if(cellData instanceof String){
						row.createCell(i+1).setCellValue((String)cellData);
					}else if(cellData instanceof Long){
						row.createCell(i+1).setCellValue((Long)cellData);
					}else if(cellData instanceof Boolean){
						row.createCell(i+1).setCellValue((Boolean)cellData);
					}
				}
				
			}
			
			
			wb.write(fileOut);
			fileOut.close();
		}catch(Exception exception){
			logger.error("Error in creating summary",exception);
		}
		
		return reportList;
	}

}
