package com.infy.gcoe.transform.util;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.FileInputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.FileSystems;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.commons.io.IOUtils;
import org.apache.commons.lang3.StringUtils;
import org.json.JSONObject;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;





import com.infy.gcoe.vo.SMCCalculatorVO;

public class VB2JavaScriptUtil {
	
	private JSONObject wordMap = null;
	private static Logger logger = LoggerFactory.getLogger(VB2JavaScriptUtil.class);
	private static String templateStringFor = "for(<variable> =<intialize>; <variable> <= <conditionoperand>;<variable>++) {";
	
	public VB2JavaScriptUtil(InputStream rulesFile) {
		super();
		try{
            String rules = IOUtils.toString(rulesFile);
            wordMap = new JSONObject(rules); 
		}catch(Exception ex){
			logger.error("Error in loading maps rules ", ex);
		}
		 
	}
	
	public SMCCalculatorVO convert(String vbScriptPath, String outputJsScriptPath, SMCCalculatorVO smcCalculatorVo, List<String> baseFunctionNamesList){
		StringBuilder outputBuilder = new StringBuilder();
		
		Path path = FileSystems.getDefault().getPath(vbScriptPath);
		   
		   
		   try {
			String contents = new String(Files.readAllBytes(path), StandardCharsets.UTF_8);
//			logger.info("File contents read as string "+contents);
			
			contents = contents.replace("As Range", "");
			contents = contents.replace("Private", "");
			contents = contents.replace("Public", "");
			contents = contents.replace("Dim", "var");
			contents = contents.replace("Const", "var");
			contents = contents.replace("Set", "");
			contents = contents.replace("False", "false");
			contents = contents.replace("True", "true");
			contents = contents.replace("&", "+");
			contents = contents.replace("AND", "&&");
			contents = contents.replace("And", "&&");
			contents = contents.replace("Or", "||");
			contents = contents.replace("<>", "!=");
			contents = contents.replace("In", "in");
			contents = contents.replace("As integer", "");
			contents = contents.replace("As Variant", "");
			contents = contents.replace("As Long", "");
			contents = contents.replace("As String", "");
			contents = contents.replace("As Boolean", "");
			contents = contents.replace("As Name", "");
			contents = contents.replace("As Worksheet", "");
			contents = contents.replace("As Double", "");
			contents = contents.replace("Global", "");
			contents = contents.replace("Resume Next", "continue;");
			contents = contents.replace("Now()", "new Date()");
			contents = contents.replace("Now", "new Date()");
			contents = contents.replace("As CustomProperty", "");
			contents = contents.replace("Range(", "SpreadsheetApp.getActiveSpreadsheet().getRange(");
			contents = contents.replace(".Range", "SpreadsheetApp.getActiveSpreadsheet().getRange");
			
			contents = contents.replace("ActiveWorkbook", "ActiveWorkbook()");
			contents = contents.replace("As Workbook", "");
			contents = contents.replace("Select Case", "switch(");
			contents = contents.replace("End Select", "}");
			contents = contents.replace("Case ", "case ");
			contents = contents.replace("'", "//");
			contents = contents.replace("case Else", "case ''");
			contents = contents.replace(".Select", ".activate()");
			contents = contents.replace("End If", "}");
			contents = contents.replace("If Not", "if(!");
			contents = contents.replace("ElseIf", "}else if(");
			
			//Do while loop in VB script is converted to while loop in appscript- start
			contents = contents.replace("Do While", "while(");  
			contents = contents.replace("Exit Do", "break;");
			contents = contents.replace("Loop", "}");
			//Do while loop in VB script is converted to while loop appscript- end
			
			contents = contents.replace("Else", "} else {");
			contents = contents.replace("Then", "){");
			contents = contents.replace("Not", "!");
			contents = contents.replace("If", "if(");
			contents = contents.replace("End Sub", "}");
			contents = contents.replace("Exit Sub", "exit; ");
			contents = contents.replace("Exit Function", "exit; //");
			contents = contents.replace("Sub", "function");
			contents = contents.replace("Exit For", "break;");
			contents = contents.replace("Next", "} //");
			contents = contents.replace("Function", "function");
			contents = contents.replace("ThisWorkbook.Worksheets", "SpreadsheetApp.getActiveSpreadsheet().getSpreadSheetByName");
			contents = contents.replace("End function", "}");
			contents = contents.replace("Option Explicit", "");
			contents = contents.replace("Attribute", "var");
			contents = contents.replace("With ", "with(");
			contents = contents.replace("While ", "while(");
			contents = contents.replace("End With", "}");
			contents = contents.replace(":=", "=");
			contents = contents.replace(".Activate", ".activate()");
			contents = contents.replace("Wend", "}");
			contents = contents.replace("For Each", "for( var ");
			contents = contents.replace("ActiveWorkbook.Names", "SpreadsheetApp.getActiveSpreadsheet().getSheets()");
			contents = contents.replace("ByVal", "");
			contents = contents.replace("UCase$", "UCase");
			contents = contents.replace("Left$", "Left");
			contents = contents.replace("Right$", "Right");
			contents = contents.replace("Optional", "");
			contents = contents.replace("$", "");

			//Need to identify replacements for the following - start 
			contents = contents.replace("GoTo", "GoTo"); 
			contents = contents.replace("On Error", "On Error"); 
			contents = contents.replace("Application.DisplayAlerts", "//Application.DisplayAlerts");
			contents = contents.replace("Application.ScreenUpdating", "//Application.ScreenUpdating");
			contents = contents.replace("Application.StatusBar", "//Application.StatusBar");
			contents = contents.replace("Application.FileDialog", "//Application.FileDialog");
			contents = contents.replace("Application.Evaluate", "//Application.Evaluate");
			contents = contents.replace("Application.WindowState", "//Application.WindowState");
			contents = contents.replace("ActiveWindow.DisplayGridlines", "//ActiveWindow.DisplayGridlines");
			
			//Need to identify replacements for the following - end
			
		
			writeAppScriptToFile(contents,outputJsScriptPath);
			convertToAppscript(outputJsScriptPath);
			
			
		} catch (IOException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}
		   
		
		
		
		
		/*
		
		try{
			String line = null;
			BufferedReader inputReader = new BufferedReader(new FileReader(vbScriptPath));
			
			int lineCounter = 0;
			long hitCounter = 0;
			long missedCounter = 0;
			long customFucntionsCount = 0;
			while((line =inputReader.readLine())!= null){
				lineCounter++;
				logger.trace("{} : Line = {}",lineCounter, line);
				int wordCounter = 0;
				boolean keyWordFoundFlag = false;
				
				
				String words[] = split(line);
				
				
				
				List<String> parsedWords = new ArrayList<>();
				String lineBefore = "";
				String lineAfter = "";	
				for(String str : words){
					wordCounter++;
					
					// count the matched base functions for each line
					if(baseFunctionNamesList.contains(str)){
						customFucntionsCount++;
					}
					
					//Search for conditions like END SUB
					logger.trace("{}, finding string : {}",wordCounter, str);
					JSONObject jsonObj = get(str,null,null);						
					String rule = null;
					if(jsonObj != null){
						keyWordFoundFlag = true;
						Iterator<String> keysItr = jsonObj.keys();
						String key = null;
						while(keysItr.hasNext()){
							key = keysItr.next();
							if(key.startsWith("on") && Arrays.asList(words).contains(key.substring(2))){
								rule = key;
							}
						}
					}
					
					JSONObject data = get(str, rule, null);
					
					if(data == null){
						parsedWords.add(str);
						continue;
					}
					
					StringBuffer jsword = new StringBuffer();	
					
					jsword.append(getStringIfExists(data, "before", ""))
							.append(getStringIfExists(data,"value", ""))
							.append(getStringIfExists(data, "after", ""));
					
					lineBefore = getStringIfExists(data,"lineBefore", lineBefore);	 
				    lineAfter = getStringIfExists(data,"lineAfter", lineAfter);   
				    
				    
				    parsedWords.add(jsword.toString());
				    		
				}
				
				
				boolean lineCallable = true;
				for(String key : wordMap.keySet()){
					String array[] = new String[]{":","="};
					if(Arrays.asList(array).contains(key)){
						continue;
					}
					
					if(Arrays.asList(words).contains(key)){
						lineCallable = false;
						break;
					}
				}
				
				
				int openAt = 0;
				int closeAt = 0;						
				if(lineCallable){
					for(int i=0;i<parsedWords.size();i++){
						String word = parsedWords.get(i);
						if(StringUtils.indexOfAny(word, new String[]{" ", "(", ")","'", ",", ".", ":","="}) != -1){
							continue;
						}
						
						if(openAt == 0){
							openAt = i;
						}
						closeAt = i;						
					}
					
					if(openAt > 0){
						parsedWords.add(openAt+1,"(");
						parsedWords.add(closeAt+2,")");
					}
				}
				
				
				outputBuilder.append(lineBefore);
			    for(String pw : parsedWords){
			    	outputBuilder.append(pw);
			    }
			    outputBuilder.append(lineAfter);
			    outputBuilder.append("\n");
			    
			    if(keyWordFoundFlag){
			    	hitCounter++;
			    }else{
			    	missedCounter++;
			    }
			}
			inputReader.close();
			
			
			smcCalculatorVo.setHitCount(hitCounter);
			smcCalculatorVo.setMissedCount(missedCounter);
			smcCalculatorVo.setCustomFunctionsCount(customFucntionsCount);
			smcCalculatorVo.setLineCount(lineCounter);
			smcCalculatorVo.setName(outputJsScriptPath);
			
			
			//Create parse content to file
			writeAppScriptToFile(outputBuilder.toString(),outputJsScriptPath);
			
		}catch(Exception ex){
			logger.error("Error in conversion ", ex);
		}
	*/	
		
		return smcCalculatorVo;
	}
	
	
	
	private static void convertToAppscript(String targetFilePath) {
		StringBuilder outputBuilder = new StringBuilder();
		try{
			String line = null;
			BufferedReader inputReader = new BufferedReader(new FileReader(targetFilePath));
			List<String> functionsList = new ArrayList<>();
			while((line =inputReader.readLine())!= null){
				 String forRegex ="(For\\s)"; 
				 String forEachregex = "(for\\(\\s)(var)";
				 String functionRegex ="(function\\s)";	
				 String whileRegex = "(while\\()";
				 String withRegex  = "(with\\()";
				 String switchRegex = "(switch\\()";
				 String caseRegex = "(case\\s)";
				 
			     line = computeRegexFor(forRegex, line,templateStringFor);
			     line = computeRegexFunction(functionRegex,line);
			     
			     line = computeRegexWhileAndWith(whileRegex,line);
			     line = computeRegexWhileAndWith(withRegex,line);
			     line = computeRegexWhileAndWith(switchRegex,line);
			     line = computeRegexWhileAndWith(forEachregex,line);
			     line = computeRegexCase(caseRegex,line);
			     
				  outputBuilder.append(line);
				  outputBuilder.append("\n");
				  
			}
			inputReader.close();
			//Create parse content to file
			writeAppScriptToFile(outputBuilder.toString(),targetFilePath);
				
			
			
			
		}catch(Exception ex){
			System.out.println("Error in conversion "+ex);
		}
		 
	     
		
	}



	private static String computeRegexCase(String caseRegex, String line) {
		Pattern r = Pattern.compile(caseRegex);
		 Matcher m = null;
		 m = r.matcher(line);
		if (m.find()) {
			 line = line+" :"; 
		 }else{
		 }
		return line;
		
		
		
	}



	/**
	 * @param line
	 * @return
	 */
	private static String computeRegexFor(String regexPattern, String line, String templateString) {
		Pattern r = Pattern.compile(regexPattern);
		 Matcher m = null;
		 String variable = ""; 
		 String intialize = "";
		 String conditionOperand ="";
		 
		 
		
		m = r.matcher(line);
		if (m.find()) {
			 
			 String[] splitStr = split(line);
			 
			 for(String s : splitStr){
				 
				 s = s.trim();
				 if(s.startsWith("For")){
	    			 variable = s.substring(s.indexOf("For")+3,s.length());
	    		 }else if(s.indexOf("To")>=1){
	    			 intialize = s.substring(0,s.indexOf("To")-1);
	    			 conditionOperand = s.substring(s.indexOf("To")+2,s.length());
	    		 }

			 }
			 
			
			 templateString = templateString.replace("<variable>", variable);
			 templateString = templateString.replace("<intialize>", intialize); 
		     templateString = templateString.replace(" <conditionoperand>", conditionOperand);
			 
		     line = templateString;
		     
		 }else{
		 }
		return line;
	}


	private static String computeRegexFunction(String regexPattern, String line) {
		Pattern r = Pattern.compile(regexPattern);
		 Matcher m = null;
		 m = r.matcher(line);
		if (m.find()) {
			 System.out.println("function regex matched "+line+" function Names "+line.substring(9, line.indexOf("(")));
//			 functionNamesList.add(line.substring(9, line.indexOf("(")).trim());
			
			 line = line+" { ";
			
		 }else{

		 }
		return line;
	}

	private static String computeRegexWhileAndWith(String regexPattern, String line) {
		Pattern r = Pattern.compile(regexPattern);
		 Matcher m = null;
		 m = r.matcher(line);
		if (m.find()) {
			 line = line+"){ "; 
		 }else{
		   //System.out.println("NO MATCH");
		 }
		return line;
	}


	
	/**
	 * 
	 * @param content
	 * @param outputFile
	 * @throws Exception
	 */
	private static void writeAppScriptToFile(String content, String outputFile) throws Exception {
		BufferedWriter writer = new BufferedWriter(new FileWriter(outputFile));
		
		writer.write(content);
		writer.flush();
		
		writer.close();
	}
	
	
	private static String[] split(String line){
		if(line != null){
			return line.split("\\=");
		}
		return new String[0];
	}


	
	
	

}
