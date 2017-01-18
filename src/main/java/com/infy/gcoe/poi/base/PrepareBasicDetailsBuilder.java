package com.infy.gcoe.poi.base;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.eventfilesystem.POIFSReader;
import org.apache.poi.poifs.eventfilesystem.POIFSReaderEvent;
import org.apache.poi.poifs.eventfilesystem.POIFSReaderListener;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import com.infy.gcoe.poi.GenerateReport;
import com.infy.gcoe.poi.vo.ExcelReportVO;
import com.infy.gcoe.poi.vo.ExcelSheetVO;

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
	public List<ExcelReportVO> update(List<ExcelReportVO> reportList) throws Exception {
		
		
		for(ExcelReportVO report : reportList){
			
			Workbook wb_xssf      = null; //Declare XSSF WorkBook 
			Workbook wb_hssf      = null; //Declare HSSF WorkBook 
			Sheet sheet           = null; //sheet can be used as common for XSSF and HSSF WorkBook 
			
			File file             = report.getFile();
			String fileName       = report.getFileName();
			String fileExtension  = fileName.substring(fileName.lastIndexOf('.'));
			report.setFileExtension(fileExtension);
			
			if(fileExtension.endsWith("xls")                        //Legacy Excel worksheets; officially designated "Microsoft Excel 97-2003 Worksheet"
					|| fileExtension.endsWith("xlt")                //Legacy Excel templates; officially designated "Microsoft Excel 97-2003 Template"
					|| fileExtension.endsWith("xlm")){              //Legacy Excel macro
				
			    wb_hssf = new HSSFWorkbook(new BufferedInputStream( new FileInputStream(file)));
			    
			    Iterator<Sheet> sheetItr = wb_hssf.sheetIterator();
			    int sheetCnt = 0;
			    while(sheetItr.hasNext()){
			    	parseSheet(report, sheetItr.next(),++sheetCnt);
			    }			    
			    report.setHasMacros(isMacroDetected(report.getFile()));
			    report.isExcelFile(true);
			    report.setOldFormat(true);
			    
			}else if (fileExtension.endsWith("xlsx")                //Excel workbook
					|| fileExtension.endsWith("xlsm")               //Excel macro-enabled workbook; same as xlsx but may contain macros and scripts
					|| fileExtension.endsWith("xltx")               //Excel template
					|| fileExtension.endsWith("xltm")){             //Excel macro-enabled template; same as xltx but may contain macros and scripts
				
			    wb_xssf = new XSSFWorkbook(new BufferedInputStream( new FileInputStream(file)));  
			    
			    Iterator<Sheet> sheetItr = wb_xssf.sheetIterator();
			    int sheetCnt = 0;
			    while(sheetItr.hasNext()){
			    	parseSheet(report, sheetItr.next(),++sheetCnt);
			    }		
			    report.setHasMacros(isMacroDetected(report.getFile()));
			    report.isExcelFile(true);
			    report.setOldFormat(false);
			    
			}else if(fileExtension.endsWith("xlsb")                //Excel binary worksheet (BIFF12)
					|| fileExtension.endsWith("xla")               //Excel add-in or macro
					|| fileExtension.endsWith("xlam")              //Excel add-in
					|| fileExtension.endsWith("xll")               //Excel XLL add-in; a form of DLL-based add-in[1]
					|| fileExtension.endsWith("xlw")){             //Excel work space; previously known as "workbook"
				throw new Exception("Files with extension :" + fileExtension + ", are not yet supported");
			}else{
				throw new Exception("Invalid Files with extension :" + fileExtension);
			}
			
		}
		
		
		return reportList;
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
	
	
	private boolean isMacroDetected(File file){
		try{
			POIFSReader r = new POIFSReader();
		    MacroListener ml = new MacroListener();
		    r.registerListener(ml);
		    FileInputStream fis = new FileInputStream(file);
		    r.read(fis);
		    return ml.isMacroDetected();
		}catch(Exception ex){
			logger.error("Error in detecting macro");
		}
		return false;
	}
	
	public class MacroListener implements POIFSReaderListener {

		  boolean macroDetected;

		  public boolean isMacroDetected() {
		    return macroDetected;
		  }

		  public void processPOIFSReaderEvent(POIFSReaderEvent event) {
		    if(event.getPath().toString().startsWith("\\Macros")
		          || event.getPath().toString().startsWith("\\_VBA")) {
		      macroDetected = true;
		    }

		  }
		}

}
