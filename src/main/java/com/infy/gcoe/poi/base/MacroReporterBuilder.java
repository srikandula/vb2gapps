package com.infy.gcoe.poi.base;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.apache.poi.poifs.macros.VBAMacroReader;
import org.springframework.stereotype.Service;

import com.infy.gcoe.poi.vo.ExcelMacroVO;
import com.infy.gcoe.poi.vo.ExcelReportVO;

/**
 * Step 3: Prepare the list of embedded macros present in the passed file
 * 
 * @author srinivas.kandula
 *
 */
@Service
public class MacroReporterBuilder implements IReportBuilder {

	public List<ExcelReportVO> update (List<ExcelReportVO> reportList) throws Exception {		
		//Extracting the files
		VBAMacroReader macroReader      = null;
		Map<String,String> macroMap     = null;
		List<ExcelMacroVO> macroList    = null;
		ExcelMacroVO macroVO            = null;
		String macroData                = null;
		
		for(ExcelReportVO report : reportList){
			if(!report.hasMacros()){
				continue;
			}
			//Using POI API to read file
			macroReader = new VBAMacroReader(report.getFile());
			//Extracting macros from the file
			macroMap = macroReader.readMacros();
			//Creating VO object to hold data
			macroList = new ArrayList<>();
			
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
				macroList.add(macroVO);
			}
			//Add the list to file
			report.setMacroList(macroList);
		}
		
		return reportList;
	}
}
