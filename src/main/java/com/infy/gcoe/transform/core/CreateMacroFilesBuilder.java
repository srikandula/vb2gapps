package com.infy.gcoe.transform.core;

import java.io.File;

import org.apache.poi.poifs.macros.VBAMacroExtractor;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import com.infy.gcoe.poi.vo.ExcelReportVO;

/**
 * Step 2: Identify the basic details like number of sheets, number of rows/columns per sheet
 *
 * @author srinivas.kandula
 *
 */
@Service
public class CreateMacroFilesBuilder implements ITransformBuilder {

	private static Logger logger = LoggerFactory.getLogger(CreateMacroFilesBuilder.class);

	private String fileOutout = null;
	
	public void setReportPath(String fileOutout){
		this.fileOutout = fileOutout;
	}
	
	@Override
	public ExcelReportVO run(ExcelReportVO report) throws Exception {

		VBAMacroExtractor macroExt = new VBAMacroExtractor();
		
		File outDir = new File(fileOutout + "/" + report.getFileName() + "/");
		
		logger.info("Generating Macros for : " + report.getAbsolutePath());
		macroExt.extract(new File(report.getAbsolutePath()), outDir);
		
		//Post creation read the macro created from the output directory as the above api directly writes to folder
		report.setVbMacroFiles(outDir.listFiles());
		
		return report;
	}

	

}
