package com.infy.gcoe.transform.core;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import com.infy.gcoe.poi.vo.ExcelReportVO;

/**
 * Step 1: Re-Create the Summary Report
 *
 * @author srinivas.kandula
 *
 */
@Service
public class ReadSummaryReportBuilder implements ITransformBuilder {

	private static Logger logger = LoggerFactory.getLogger(ReadSummaryReportBuilder.class);
	
	private String summaryReportFileName = null;
	
	public void setSummaryReportFileName(String summaryReportFileName){
		this.summaryReportFileName = summaryReportFileName;
	}

	@Override
	public List<ExcelReportVO> run(List<ExcelReportVO> reportList) throws Exception {

		XSSFWorkbook wb = new XSSFWorkbook(new File(summaryReportFileName));
		XSSFSheet summarySheet = wb.getSheet("summary");
		int counter = 0 ;
		
		Map<Integer,String> headerMap = new HashMap<>();
		XSSFRow row = null;
		Iterator<Cell> cellItr = null;
		
		//Understand header columns
		row = summarySheet.getRow(0);		
		cellItr = row.cellIterator();
		while(cellItr.hasNext()){
			String headerVal = cellItr.next().getStringCellValue();
			headerMap.put(counter,headerVal);
			counter++;
		}
		
		//Reading each row and preparing the excel report vo
		for(int i=1;i<summarySheet.getLastRowNum();i++){
			row = summarySheet.getRow((short)i);
			cellItr = row.cellIterator();
			
			ExcelReportVO reportVO = new ExcelReportVO();
			counter = 0 ; //As first row is row counter, skip this
			cellItr = row.cellIterator();
			while(cellItr.hasNext()){
				Cell cellValue = cellItr.next();
				
				if(cellValue.getCellTypeEnum() == CellType.STRING){
					reportVO.setData(
							headerMap.get(counter),                 //Identify Header
							cellValue.getStringCellValue()            //Pick the Value
					);
				}else if(cellValue.getCellTypeEnum() == CellType.BOOLEAN){
					reportVO.setData(
							headerMap.get(counter),                 //Identify Header
							cellValue.getBooleanCellValue()           //Pick the Value
					);
				}else if(cellValue.getCellTypeEnum() == CellType.NUMERIC){
					reportVO.setData(
							headerMap.get(counter),                  //Identify Header
							(long)cellValue.getNumericCellValue()      //Pick the Value
					);
				}else{
					//Ignore cell value
				}
				counter++;				
			}
			reportList.add(reportVO);
		}
		

		return reportList;
	}

	
}
