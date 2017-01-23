package com.infy.gcoe.transform.core;

import java.io.File;
import java.io.InputStream;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import com.infy.gcoe.transform.util.VB2JavaScriptUtil;
import com.infy.gcoe.vo.ExcelReportVO;

/**
 * Step 3: Converts the VB Script files to AppScript Files
 *
 * @author srinivas.kandula
 *
 */
@Service
public class ConvertVB2JSBuilder implements ITransformBuilder {

	private static Logger logger = LoggerFactory.getLogger(ConvertVB2JSBuilder.class);
	
	@Override
	public ExcelReportVO run(ExcelReportVO report) throws Exception {
		
		ClassLoader loader           = Thread.currentThread().getContextClassLoader();
		InputStream resourceStream   = loader.getResourceAsStream("./vb2js/map.json");
		VB2JavaScriptUtil util       = new VB2JavaScriptUtil(resourceStream);
		File vbScrpts[]              = report.getVbMacroFiles();	
		resourceStream.close();
		
		String macroFile       = null;
		String jsFile          = null;		
		
		for(int i=0;i<vbScrpts.length;i++){
			macroFile = vbScrpts[i].getAbsolutePath();
			jsFile = macroFile.substring(0,macroFile.lastIndexOf('.')) + ".js";
			
			logger.info("Converting {} to {}",macroFile,jsFile);
			util.convert(macroFile, jsFile);
		}	

		
		return report;
	}

	

}
