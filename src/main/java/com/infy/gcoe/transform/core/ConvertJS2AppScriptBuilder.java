package com.infy.gcoe.transform.core;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.commons.io.FileUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import com.infy.gcoe.transform.util.VB2JavaScriptUtil;
import com.infy.gcoe.vo.ExcelReportVO;
import com.infy.gcoe.vo.SMCCalculatorVO;

@Service
public class ConvertJS2AppScriptBuilder implements ITransformBuilder {

	private static Logger logger = LoggerFactory.getLogger(ConvertJS2AppScriptBuilder.class);
	
	@Override
	public ExcelReportVO run(ExcelReportVO report) throws Exception {
		
		ClassLoader loader          		 = Thread.currentThread().getContextClassLoader();
		
		InputStream loadBaseFunctionStream   = loader.getResourceAsStream("./vb2js/BaseFunctions.gs");
		
		File vbScrpts[]              		 = report.getVbMacroFiles();	
		
		List<String> baseFunctionNamesList = listBaseFunctionNames(loadBaseFunctionStream);
	    logger.info("baseFunctionNamesList "+baseFunctionNamesList);
		
		
		for(int i=0;i<vbScrpts.length;i++){
			
			String name = vbScrpts[i].getName();
			if(name.lastIndexOf("vba")>=0){
				name = name.replace("vba", "js");
			}
			String filePath = vbScrpts[i].getAbsolutePath();
			filePath = filePath.substring(0, filePath.lastIndexOf("\\"));
			
			File sourceFile = new File(filePath+"\\"+name);
			if(name.lastIndexOf("js")>=0){
				name = name.replace("js", "gs");
			}
		
		
			File targetFile = new File(filePath+"\\"+name);
			copyFileUsingApacheCommonsIO(sourceFile,targetFile);
			
			convertToAppscript(targetFile);
		}	


		return report;
	}

	
	private void convertToAppscript(File targetFile) {
		StringBuilder outputBuilder = new StringBuilder();
		try{
			String line = null;
			BufferedReader inputReader = new BufferedReader(new FileReader(targetFile));
			while((line =inputReader.readLine())!= null){
				 String pattern = "(for\\(\\s)(.\\D+).\\D.+";
			     Pattern r = Pattern.compile(pattern);
			     Matcher m = null;
			     String variable = ""; 
				 String intialize = "";
				 String conditionOperand ="";
				 
				 
				String newForString = "for(<variable> =<intialize>; <variable> <= <conditionoperand>;<variable>++) {";
				m = r.matcher(line);
				if (m.find()) {
			    	 
			    	 String[] splitStr = split(line);
			    	 
			    	 for(String s : splitStr){
			    		 
			    		 if(s.indexOf("r(")>0){
			    			 variable = s.substring(s.indexOf("(")+1,s.length()-1);
			    		 }else if(s.indexOf("To")>=1){
			    			 intialize = s.substring(0,s.indexOf("To")-1);
			    			 conditionOperand = s.substring(s.indexOf("To")+2,s.lastIndexOf(")"));
			    		 }
			    	 }
			    	 
			    	 System.out.println(" variable "+variable);
			    	 System.out.println(" intialize "+intialize);
			    	 System.out.println(" conditionOperand "+conditionOperand);
			    	 
			    	
			    	 newForString = newForString.replace("<variable>", variable);
			    	 newForString = newForString.replace("<intialize>", intialize); 
			         newForString = newForString.replace(" <conditionoperand>", conditionOperand);
			    	 
			         line = newForString;
			         
			     }else {
			       // System.out.println("NO MATCH");
			     }
				
				  outputBuilder.append(line);
				  outputBuilder.append("\n");
				  
			}
			inputReader.close();
			//Create parse content to file
			writeAppScriptToFile(outputBuilder.toString(),targetFile);
				
			
			
			
		}catch(Exception ex){
			logger.error("Error in conversion ", ex);
		}
		 
	     
		
	}


	/**
	 * @param loadBaseFunctionStream
	 * @return
	 */
	private List<String> listBaseFunctionNames(InputStream loadBaseFunctionStream) {
		BufferedReader reader = null;
		 List<String> baseFunctionNamesList = new ArrayList<>();
		   try {
	            reader = new BufferedReader(new InputStreamReader(loadBaseFunctionStream));
	            String line = reader.readLine();
	            while(line != null){
	               if(line.startsWith("function")) {
	            	   if(line.indexOf('(')>=0){
	            		   baseFunctionNamesList.add(line.substring(8, line.indexOf('(')).trim());
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
	
	private static void copyFileUsingApacheCommonsIO(File source, File dest) throws IOException {
	    FileUtils.copyFile(source, dest);
	}
	

private static String[] split(String line){
	if(line != null){
		return line.split("\\=");
	}
	return new String[0];
}



/**
 * 
 * @param content
 * @param outputFile
 * @throws Exception
 */
private void writeAppScriptToFile(String content, File outputFile) throws Exception {
	BufferedWriter writer = new BufferedWriter(new FileWriter(outputFile));
	
	writer.write(content);
	writer.flush();
	
	writer.close();
}
}
