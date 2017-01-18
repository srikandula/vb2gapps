package com.infy.gcoe.poi;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.ApplicationArguments;
import org.springframework.boot.CommandLineRunner;
import org.springframework.context.annotation.Profile;
import org.springframework.stereotype.Component;

import com.infy.gcoe.poi.base.MacroReporterBuilder;
import com.infy.gcoe.poi.base.PrepareFileListBuilder;
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

	@Autowired
	PrepareFileListBuilder fileListBuilder;


	@Autowired
	MacroReporterBuilder macroReportBuilder;

	public GenerateReport(ApplicationArguments args){
		
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
		
		//Read Microsoft files from share folders
		for(String fileName : source){
			fileListBuilder.updateFileDetails(new File(fileName), reportList);
		}
		
		//Generate Macro Report
		macroReportBuilder.updateMacroDetails(reportList);
		
		logger.info("Read folders " + reportList);
	}
	
	
}
