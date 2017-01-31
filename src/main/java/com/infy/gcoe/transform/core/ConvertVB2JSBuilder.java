package com.infy.gcoe.transform.core;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.List;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import com.infy.gcoe.transform.util.VB2JavaScriptUtil;
import com.infy.gcoe.vo.ExcelReportVO;
import com.infy.gcoe.vo.SMCCalculatorVO;

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
		
		ClassLoader loader          		 = Thread.currentThread().getContextClassLoader();
		InputStream resourceStream   		 = loader.getResourceAsStream("./vb2js/map.json");
		InputStream loadBaseFunctionStream   = loader.getResourceAsStream("./vb2js/BaseFunctions.gs");
		VB2JavaScriptUtil util       		 = new VB2JavaScriptUtil(resourceStream);
		File vbScrpts[]              		 = report.getVbMacroFiles();	
		resourceStream.close();
		
		List<String> baseFunctionNamesList = listBaseFunctionNames(loadBaseFunctionStream);
	    logger.info("baseFunctionNamesList "+baseFunctionNamesList);

	    String macroFile       = null;
		String jsFile          = null;		
		List<SMCCalculatorVO> SMCCalculatorList = new ArrayList<>();
		
		for(int i=0;i<vbScrpts.length;i++){
			macroFile = vbScrpts[i].getAbsolutePath();
			jsFile = macroFile.substring(0,macroFile.lastIndexOf('.')) + ".gs";
			SMCCalculatorVO SMCCalculatorVO = new SMCCalculatorVO();
			SMCCalculatorVO =util.convert(macroFile, jsFile,SMCCalculatorVO,baseFunctionNamesList);
			if(SMCCalculatorVO.getLineCount()<=25 ){
				SMCCalculatorVO.setComplexity("Simple");
			}else if(SMCCalculatorVO.getLineCount()>25 || SMCCalculatorVO.getLineCount()<=50){
				SMCCalculatorVO.setComplexity("Medium");
			}else{
				SMCCalculatorVO.setComplexity("Complex");
			}
			 
			 SMCCalculatorList.add(SMCCalculatorVO);
			
		}	

		logger.info("Complexity classification "+SMCCalculatorList);
		return report;
	}

	/**
	 * @param loadBaseFunctionStream
	 * @return
	 */
	private List<String> listBaseFunctionNames(
			InputStream loadBaseFunctionStream) {
		BufferedReader reader = null;
		 List<String> baseFunctionNamesList = new ArrayList<>();
		   try {
	            reader = new BufferedReader(new InputStreamReader(loadBaseFunctionStream));
	            String line = reader.readLine();
	            while(line != null){
	               if(line.startsWith("function")) {
	            	   if(line.indexOf('(')>=0){
	            		   baseFunctionNamesList.add(line.substring(7, line.indexOf('(')).trim());
	            	   }
	               }
	                line = reader.readLine();
	            }           
	          
	        } catch (FileNotFoundException ex) {
	        	logger.error("File Not Found exception "+ex);
	        } catch (IOException ex) {
	        	logger.error("IOException occured while reading the file "+ex);
	          
	        } finally {
	            try {
	                reader.close();
	                loadBaseFunctionStream.close();
	            } catch (IOException ex) {
	            	logger.error("IOException occured while closing the file "+ex);
	            }
	        }
		return baseFunctionNamesList;
	}
	

}
