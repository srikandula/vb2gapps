package com.infy.gcoe.transform.util;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.FileInputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;

import org.apache.commons.io.IOUtils;
import org.apache.commons.lang3.StringUtils;
import org.json.JSONObject;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;


import com.infy.gcoe.vo.SMCCalculatorVO;

public class VB2JavaScriptUtil {
	
	private JSONObject wordMap = null;
	private static Logger logger = LoggerFactory.getLogger(VB2JavaScriptUtil.class);
	
	
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
		
		
		return smcCalculatorVo;
	}
	
	/**
	 * 
	 * @param content
	 * @param outputFile
	 * @throws Exception
	 */
	private void writeAppScriptToFile(String content, String outputFile) throws Exception {
		BufferedWriter writer = new BufferedWriter(new FileWriter(outputFile));
		
		writer.write(content);
		writer.flush();
		
		writer.close();
	}
	
	/**
	 * Splits give line into string [] based on characters blank space or ()\',.:=
	 * 
	 * @param line
	 * @return
	 */
	private String[] split(String line){
		if(line != null){
			return line.split("(?=[ ()\'.,:=])|(?<=[ ()\'.,:=])");
		}
		return new String[0];
	}
	
	/**
	 * Search the JSONMap
	 * 1. If both name and rule are passed, search 'name.rule' that as key
	 * 2. If only name is passed, search 'name' as key
	 * 3. If value not found return the defaultVal
	 * 
	 * @param name
	 * @param rule
	 * @param defaultVal
	 * @return
	 */
	private JSONObject get(String name, String rule, String defaultVal){
		
		JSONObject val = null;		
		StringBuilder keyBuilder = new StringBuilder(name);
	
		if(!wordMap.isNull(keyBuilder.toString())){
			val = wordMap.getJSONObject(keyBuilder.toString());
			if(rule != null && !val.isNull(rule)){
				val = val.getJSONObject(rule);
			}
		}
		
		if(val == null && defaultVal != null){
			val = new JSONObject("{\"value\" : "+ defaultVal +"}");
		}
		
		return val;
	}
	
	/**
	 * Returns value if exists otherwise default value if not exists
	 * 
	 * @param json
	 * @param name
	 * @param defaultVal
	 * @return
	 */
	private String getStringIfExists(JSONObject json, String name, String defaultVal){
		String val = defaultVal;
		
		
		if(json != null && !json.isNull(name)){
			val = json.getString(name);			
		}
		
		return val;		
	}

}
