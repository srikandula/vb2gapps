package com.infy.gcoe.poi;

import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.apache.poi.poifs.macros.VBAMacroExtractor;
import org.apache.poi.poifs.macros.VBAMacroReader;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.ApplicationArguments;
import org.springframework.boot.CommandLineRunner;
import org.springframework.context.annotation.Profile;
import org.springframework.stereotype.Component;

import com.infy.gcoe.poi.base.ReadFolder;
import com.infy.gcoe.poi.vo.ExcelMacroVO;
import com.infy.gcoe.poi.vo.ExcelReportVO;
/**
 * 
 * Run Command from Command prompt
 *     mvn spring-boot:run -Drun.arguments="--spring.profiles.active=GenerateReport,--report.source=./excel,--report.dest=./report" -Preport
 * @author srinivas.kandula
 *
 */
@Component
@Profile(value="GenerateReport")
public class GenerateReport implements CommandLineRunner {

	 private static Logger logger = LoggerFactory.getLogger(GenerateReport.class);
	 
	private List<String> source = null;
	private List<String> dest = null;
	private String LINE_SEPERATOR = null; 

	@Autowired
	ReadFolder readFolder;
	
	public GenerateReport(ApplicationArguments args){
		LINE_SEPERATOR = System.getProperty("line.separator");
		
		if(args.containsOption("report.source")){
			source = args.getOptionValues("report.source");
		}else{
			source = new ArrayList<>();
			source.add("./excel");
		}
		
		if(args.containsOption("report.dest")){
			dest = args.getOptionValues("report.dest");
		}else{
			dest = new ArrayList<>();
			dest.add("./report");
		}
	}

	@Override
	public void run(String[] args) throws Exception {
		
		List<ExcelReportVO> reportList = new ArrayList<>();
		for(String fileName : source){
			readFolder.read(new File(fileName), reportList);
		}

		//Extracting the files
		VBAMacroReader macroReader      = null;
		Map<String,String> macroMap     = null;
		List<ExcelMacroVO> macroList    = null;
		ExcelMacroVO macroVO            = null;
		String macroData                = null;
		
		for(ExcelReportVO report : reportList){
			//Using POI API to read file
			macroReader = new VBAMacroReader(report.getFile());
			//Extracting macros from the file
			macroMap = macroReader.readMacros();
			//Creating VO object to hold data
			macroList = new ArrayList<>();
			
			for(String macroName : macroMap.keySet()){
				//Reading macro data
				macroData       = macroMap.get(macroName);
				int lineCount   = macroData.split(LINE_SEPERATOR).length;
				
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
		
		logger.info("Read folders " + reportList);
	}
}
