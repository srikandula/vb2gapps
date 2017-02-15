package com.infy.gcoe.transform.util;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.springframework.context.annotation.Profile;
import org.springframework.stereotype.Component;

public class Test {

	/*
	private static final String inputfile = "D:/Infosys/PWC/excel-to-appscript/doc/python/VB2JS/test.vbs.txt";
	private static final String rulefile  = "D:/Infosys/PWC/excel-to-appscript/doc/python/VB2JS/map.json";*/
	
/*public static void main(String[] args) throws Exception {
	String	line = " ValidateMutuallyinclusive STPRDurColumn, STPRTypeColumn)";
	String[] functionNamesList = {"DataValidation_Click","ValidClean_Click","ValidateData","ValidationCleanup","ValidateNumber","ValidateUnnamedBranches","ValidateWithCodeLists","Check_Value","ValidateDate","ValidateMutuallyExclusive","ValidateMutuallyinclusive","ValidateRequiredField","ValidateRequiredEmptyField","ValidateWithString","MarkWrong","CheckUnnamed","FindinStrForward","isUnnamed"};
	

	for(String s: functionNamesList){
		System.out.println("line.trim().indexOf(s)>=0 "+line.indexOf(s));
		if(line.trim().indexOf(s)>=0){
			
			System.out.println("index matched "+s);		
		}
	}
	
	String[] str = split(line);
		
	
	ClassLoader loader          		 = Thread.currentThread().getContextClassLoader();
	InputStream loadBaseFunctionStream   = loader.getResourceAsStream("./vb2js/Test.gs");
	
	
	
     String[] sourceString = {"function DataValidation_Click()","function ValidateUnnamedBranches(Col )","for( A = 1 To 10) {","For FindCha = 1 To (Len(FindIn) - Len(ToFind) + 1))","for ( var cprop in SpreadsheetApp.getActiveSpreadsheet().getSheetByName('FVE Validation').CustomProperties) {"};     
     
     String pattern = "(For\\s)";
     String functionRegex ="(function\\s)";
     
     Pattern r = Pattern.compile(functionRegex);
     
     Matcher m = r.matcher(sourceString[1]);
     String variable = ""; 
	 String intialize = "";
	 String conditionOperand ="";
	 
     if (m.find()) {
    	 
    	 String[] splitStr = split(sourceString[1]);
    	 for(String s : splitStr){
    		 System.out.println("s "+s);
    		 if(s.startsWith("For")){
    			 variable = s.substring(s.indexOf("For")+3,s.length());
    		 }else if(s.indexOf("To")>=1){
    			 intialize = s.substring(0,s.indexOf("To")-1);
    			 conditionOperand = s.substring(s.indexOf("To")+2,s.length());
    		 }
    	 }
    	 
    	 System.out.println(" variable "+variable);
    	 System.out.println(" intialize "+intialize);
    	 System.out.println(" conditionOperand "+conditionOperand);
    	 
    	 String newForString = "for(<variable> =<intialize>; <variable> <= <conditionoperand>;<variable>++) {";
    	 newForString = newForString.replace("<variable>", variable);
    	 newForString = newForString.replace("<intialize>", intialize); 
         newForString = newForString.replace(" <conditionoperand>", conditionOperand);
    	 
    	System.out.println("new For String "+newForString);
    	
//        System.out.println("value 0: " + m.group(0) );
//        System.out.println("value 1: " + m.group(1) );
//        System.out.println("value 1: " + m.group(2) );
//        System.out.println("group count "+m.groupCount());
//       StringBuffer sb = new StringBuffer();
//      
//      
//      System.out.println("sb "+sb.toString());
     }else {
        System.out.println("NO MATCH");
     }
     
   	    
	}
	 */


private static void listBaseFunctionNamesTest(InputStream loadBaseFunctionStream) {
	BufferedReader reader = null;
	 List<String> baseFunctionNamesList = new ArrayList<>();
	   
          
          
          

	}



private static String[] split(String line){
	if(line != null){
		return line.split("function");
	}
	return new String[0];
}


/**
 * @param loadBaseFunctionStream
 * @return
 */
private static List<String> listBaseFunctionNames(InputStream loadBaseFunctionStream) {
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
        	System.out.println("File Not Found exception "+ex);
        } catch (IOException ex) {
        	System.out.println("IOException occured while reading the file "+ex);
          
        } finally {
            try {
                reader.close();
                loadBaseFunctionStream.close();
            } catch (IOException ex) {
            	System.out.println("IOException occured while closing the file "+ex);
            }
        }
	return baseFunctionNamesList;
}






}