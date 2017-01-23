package com.infy.gcoe.upload;

import static com.infy.gcoe.util.ReportConstants.ACTION_RESP_AUTO_UPLOAD;
import static com.infy.gcoe.util.ReportConstants.SUMMARY_REPORT;

import java.util.ArrayList;
import java.util.List;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.ApplicationArguments;
import org.springframework.boot.CommandLineRunner;
import org.springframework.context.annotation.Profile;
import org.springframework.stereotype.Component;

import com.google.api.services.drive.Drive;
import com.infy.gcoe.upload.core.CreateDriveManager;
import com.infy.gcoe.upload.core.UploadFileToDrive;
import com.infy.gcoe.util.ReadSummaryReportBuilder;
import com.infy.gcoe.vo.ExcelReportVO;

/**
 * 
 * mvn spring-boot:run -Drun.arguments="--spring.profiles.active=TransformToAppScript,--report.source=./excel,--report.dest=./report" -Pupload
 * 
 * @author srinivas.kandula
 *
 */
@Component
@Profile(value="UploadExcelFile")
public class UploadExcelFile implements CommandLineRunner {

	private static Logger logger = LoggerFactory.getLogger(UploadExcelFile.class);
	
	private List<String> summary = null;
	private List<String> source = null;
	private List<String> dest = null;
	
	@Autowired
	ReadSummaryReportBuilder summaryReportBuilder;
	
	@Autowired
	CreateDriveManager driveManager;
	
	@Autowired
	UploadFileToDrive uploadFileToDrive;
	
	public UploadExcelFile(ApplicationArguments args){
		
		logger.info("About to run UploadExcelFile");
		
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
	
		Drive drive = null;
		for(ExcelReportVO report : reportList){
			if(report.isExcelFile() && ACTION_RESP_AUTO_UPLOAD.equalsIgnoreCase(report.getUserIntention())){
				
				logger.info("Uploading file {}", report.getFileName());
				
				//Step 2: Create Drive Service		
				if(drive == null){
					drive = driveManager.getDriveService();
				}
				
				//Step 3: Upload the files
				uploadFileToDrive.setDrive(drive);			
				uploadFileToDrive.run(report);
			}
		}
	}
}
