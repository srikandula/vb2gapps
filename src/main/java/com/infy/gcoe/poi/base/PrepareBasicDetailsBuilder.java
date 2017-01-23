package com.infy.gcoe.poi.base;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

import org.apache.poi.POIXMLProperties;
import org.apache.poi.POIXMLProperties.CoreProperties;
import org.apache.poi.hpsf.SummaryInformation;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import com.infy.gcoe.poi.GenerateReport;
import com.infy.gcoe.vo.ExcelReportVO;
import com.infy.gcoe.vo.ExcelSheetVO;

/**
 * Step 2: Identify the basic details like number of sheets, number of rows/columns per sheet
 *
 * @author srinivas.kandula
 *
 */
@Service
public class PrepareBasicDetailsBuilder implements IReportBuilder {

	private static Logger logger = LoggerFactory.getLogger(GenerateReport.class);

	@Override
	public ExcelReportVO update(ExcelReportVO report) throws Exception {

		File file             = report.getFile();
		String fileName       = report.getFileName();
		String fileExtension  = fileName.substring(fileName.lastIndexOf('.') + 1);
		report.setFileExtension(fileExtension);
		
		logger.debug("Reading file properties of " + fileName);

		if(fileExtension.equalsIgnoreCase("xls")                        //Legacy Excel worksheets; officially designated "Microsoft Excel 97-2003 Worksheet"
				|| fileExtension.equalsIgnoreCase("xlt")                //Legacy Excel templates; officially designated "Microsoft Excel 97-2003 Template"
				|| fileExtension.equalsIgnoreCase("xlm")){              //Legacy Excel macro

			HSSFWorkbook wb_hssf = new HSSFWorkbook(new BufferedInputStream( new FileInputStream(file)));

		    Iterator<Sheet> sheetItr = wb_hssf.sheetIterator();
		    int sheetCnt = 0;
		    while(sheetItr.hasNext()){
		    	parseSheet(report, sheetItr.next(),++sheetCnt);
		    }
		    report.isExcelFile(true);
		    report.setOldFormat(true);
		    SummaryInformation info = wb_hssf.getSummaryInformation();
		    report.setCreatedBy(info.getAuthor());
		    report.setLastModifiedBy(info.getLastAuthor());
		    report.setLastModifiedDate(info.getLastSaveDateTime() != null ? info.getLastSaveDateTime().toString() : "-");
		    
		    wb_hssf.close();
		    
		}else if (fileExtension.equalsIgnoreCase("xlsx")                //Excel workbook
				|| fileExtension.equalsIgnoreCase("xlsm")               //Excel macro-enabled workbook; same as xlsx but may contain macros and scripts
				|| fileExtension.equalsIgnoreCase("xltx")               //Excel template
				|| fileExtension.equalsIgnoreCase("xltm")){             //Excel macro-enabled template; same as xltx but may contain macros and scripts

			XSSFWorkbook wb_xssf = new XSSFWorkbook(new BufferedInputStream( new FileInputStream(file)));

		    Iterator<Sheet> sheetItr = wb_xssf.sheetIterator();
		    int sheetCnt = 0;
		    while(sheetItr.hasNext()){
		    	parseSheet(report, sheetItr.next(),++sheetCnt);
		    }
		    report.isExcelFile(true);
		    report.setOldFormat(false);
		    POIXMLProperties poiProp = wb_xssf.getProperties();
		    CoreProperties coreProp  = poiProp.getCoreProperties();
		    report.setCreatedBy(coreProp.getCreator());
		    report.setLastModifiedBy(coreProp.getLastModifiedByUser());
		    report.setLastModifiedDate(coreProp.getModified() != null ? coreProp.getModified().toString() : "-");
		    
		    wb_xssf.close();

		}else if(fileExtension.equalsIgnoreCase("xlsb")                //Excel binary worksheet (BIFF12)
				|| fileExtension.equalsIgnoreCase("xla")               //Excel add-in or macro
				|| fileExtension.equalsIgnoreCase("xlam")              //Excel add-in
				|| fileExtension.equalsIgnoreCase("xll")               //Excel XLL add-in; a form of DLL-based add-in[1]
				|| fileExtension.equalsIgnoreCase("xlw")){             //Excel work space; previously known as "workbook"
			
			throw new Exception("Files with extension :" + fileExtension + ", are not yet supported");
			
		}else{
			throw new Exception("Invalid Files with extension :" + fileExtension);
		}

		return report;
	}

	/**
	 * Collect sheet information and updates the report
	 *
	 * @param reportList
	 * @param sheet
	 * @param sequence
	 */
	private void parseSheet(ExcelReportVO reportVO, Sheet sheet, int sequence){
		ExcelSheetVO sheetVO = new ExcelSheetVO();
		sheetVO.setName(sheet.getSheetName());
		sheetVO.setRowCount(sheet.getLastRowNum());
		sheetVO.setSheetSequence(sequence);

		if(sheet.getRow(0) != null){
			sheetVO.setColCount(sheet.getRow(0).getLastCellNum());
		}
		reportVO.addSheet(sheetVO);
	}

}
