package com.infy.gcoe.poi.base;

import java.io.BufferedInputStream;
import java.io.FileInputStream;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFObjectData;
import org.apache.poi.hssf.usermodel.HSSFPictureData;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.xssf.usermodel.XSSFPictureData;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import com.infy.gcoe.vo.ExcelReportVO;

@Service
public class AdvanceStatsReportBuilder implements IReportBuilder {

	private static Logger logger = LoggerFactory.getLogger(AdvanceStatsReportBuilder.class);
	
	@Override
	public ExcelReportVO update(ExcelReportVO report) throws Exception {

		logger.debug("Reporting advance features of " + report.getFileName());
		
		if(report.isOldFormat()){			
			HSSFWorkbook workbook = new HSSFWorkbook(new BufferedInputStream( new FileInputStream(report.getFile())));
			
			List<HSSFObjectData> embeddedList = workbook.getAllEmbeddedObjects();
			List<HSSFPictureData> pictureList = workbook.getAllPictures();
			
			if(embeddedList != null){
				report.setNoOfEmbedds(embeddedList.size());
			}
			if(pictureList != null){
				report.setNoOfPictures(pictureList.size());
			}
			
			workbook.close();
			
		}else{
			
			XSSFWorkbook workbook = new XSSFWorkbook(new BufferedInputStream( new FileInputStream(report.getFile())));
			
			List<PackagePart> embeddedList = workbook.getAllEmbedds();
			List<XSSFPictureData> pictureList = workbook.getAllPictures();
			List<XSSFPivotTable> pivotList = workbook.getPivotTables();
			
			if(embeddedList != null){
				report.setNoOfEmbedds(embeddedList.size());
			}
			if(pictureList != null){
				report.setNoOfPictures(pictureList.size());
			}
			if(pivotList != null){
				report.setNoOfPivotTables(pivotList.size());
			}		
			
			workbook.close();
		}

		return report;
	}

}
