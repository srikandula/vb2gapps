package com.infy.gcoe.poi.base;

import java.util.Map;

import org.apache.poi.poifs.macros.VBAMacroReader;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import com.infy.gcoe.poi.GenerateReport;
import com.infy.gcoe.util.ReportConstants;
import com.infy.gcoe.vo.ExcelMacroVO;
import com.infy.gcoe.vo.ExcelReportVO;

/**
 * Step 3: Prepare the list of embedded macros present in the passed file
 * 
 * @author srinivas.kandula
 *
 */
@Service
public class MacroReporterBuilder implements IReportBuilder {

	private static Logger logger = LoggerFactory.getLogger(GenerateReport.class);
	
	public ExcelReportVO update (ExcelReportVO report) throws Exception {		
		//Extracting the files
		VBAMacroReader macroReader      = null;
		Map<String,String> macroMap     = null;
		ExcelMacroVO macroVO            = null;
		String macroData                = null;
		
		try{
			//Using POI API to read file
			macroReader = new VBAMacroReader(report.getFile());
			//Extracting macros from the file
			macroMap = macroReader.readMacros();
			
			for(String macroName : macroMap.keySet()){
				//Reading macro data
				macroData       = macroMap.get(macroName);
				int lineCount   = macroData.split(ReportConstants.LINE_SEPERATOR).length;
				
				//Collating required information to a VO
				macroVO         = new ExcelMacroVO();
				macroVO.setName(macroName);
				macroVO.setLineCount(lineCount);
				macroVO.setContent(macroData);
				
				//Adding to file List to track macros
				report.addMacro(macroVO);				
			}
			
		}catch(Exception ex){
			logger.debug("Error in reading macro, this may occur if passed XLS document doesn't contain macros");				
		}
		
		return report;
	}
}
