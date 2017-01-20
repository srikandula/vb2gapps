package com.infy.gcoe.transform.core;

import java.io.File;
import java.util.List;

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
	public List<ExcelReportVO> run(List<ExcelReportVO> reportList) throws Exception {


		for(ExcelReportVO report : reportList){
			VBAMacroExtractor macroExt = new VBAMacroExtractor();
			
			if(report.hasMacros()){
				logger.info("Generating Macros for : " + report.getAbsolutePath());
				macroExt.extract(new File(report.getAbsolutePath()), new File(fileOutout + "/" + report.getFileName() + "/"));
			}
		}


		return reportList;
	}

	

}
