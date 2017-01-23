package com.infy.gcoe.transform;

import static com.infy.gcoe.poi.base.ReportConstants.ACTION_RESP_VB_2_APPS;
import static com.infy.gcoe.poi.base.ReportConstants.SUMMARY_REPORT;

import java.util.ArrayList;
import java.util.List;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.ApplicationArguments;
import org.springframework.boot.CommandLineRunner;
import org.springframework.context.annotation.Profile;
import org.springframework.stereotype.Component;

import com.infy.gcoe.poi.vo.ExcelReportVO;
import com.infy.gcoe.transform.core.ConvertVB2JSBuilder;
import com.infy.gcoe.transform.core.CreateMacroFilesBuilder;
import com.infy.gcoe.transform.core.ReadSummaryReportBuilder;

/**
 * 
 * mvn spring-boot:run -Drun.arguments="--spring.profiles.active=TransformToAppScript,--report.source=./excel,--report.dest=./report" -Ptransform
 * 
 * @author srinivas.kandula
 *
 */
@Component
@Profile(value="TransformToAppScript")
public class TransformToAppScript implements CommandLineRunner {
	
	private static Logger logger = LoggerFactory.getLogger(TransformToAppScript.class);
	
	private List<String> summary = null;
	private List<String> source = null;
	private List<String> dest = null;
	
	@Autowired
	ReadSummaryReportBuilder summaryReportBuilder;
	
	@Autowired
	CreateMacroFilesBuilder macroFileBuilder; 
	
	@Autowired
	ConvertVB2JSBuilder convertVB2JSBuilder;
	
	public TransformToAppScript(ApplicationArguments args){
		
		if(args.containsOption("report.summary")){
			summary = args.getOptionValues("report.summary");
		}else{
			summary = new ArrayList<>();
			summary.add("./" + SUMMARY_REPORT);
		}
		
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
		
		//Step 1 : Read Microsoft files from share folders
		for(String fileName : summary){
			summaryReportBuilder.setSummaryReportFileName(fileName);
			summaryReportBuilder.run(reportList);
		}
	
		
		for(ExcelReportVO report : reportList){
			if(report.isExcelFile() && report.hasMacros()){
				//Step 2 : Generate separate files for each macro
				macroFileBuilder.setReportPath(dest.get(0));
				macroFileBuilder.run(report);
				
				//Step 3 : Convert files to app script
				if(ACTION_RESP_VB_2_APPS.equalsIgnoreCase(report.getUserIntention())){
					convertVB2JSBuilder.run(report);
				}
			}
		}
		
	}
}
