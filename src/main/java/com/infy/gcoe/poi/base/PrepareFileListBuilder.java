package com.infy.gcoe.poi.base;

import java.io.File;
import java.util.List;

import org.springframework.stereotype.Service;

import com.infy.gcoe.poi.vo.ExcelReportVO;

/**
 * Step 1: Identify the list of files which needs to be used for migration
 * 
 * @author srinivas.kandula
 *
 */
@Service
public class PrepareFileListBuilder implements IReportBuilder{
	
	private String source = null;
	private File sourceFile = null;
	
	/**
	 * Update the file or folder location from where to read the file(s)
	 * 
	 * @param source
	 */
	public void setSource(String source){
		this.source = source;
		sourceFile = new File(this.source);
	}

	/**
	 * Builds the list of files to the reportVO
	 * 
	 */
	@Override
	public List<ExcelReportVO> update(List<ExcelReportVO> reportList) throws Exception {
		if(this.source == null){
			throw new Exception("Source folder need to be set before invoking this");			
		}
		
		updateFileDetails(sourceFile,reportList);
		
		return reportList;
	}
	
	/**
	 * Recursively iterate the give file and extract all the files and add them to reportList
	 * 
	 * @param file
	 * @param reportList
	 * @return
	 * @throws Exception
	 */
	private List<ExcelReportVO> updateFileDetails(File file, List<ExcelReportVO> reportList) throws Exception {
		
		if (file != null && file.exists()){
			if (file.isDirectory()) {
				File[] fileList = file.listFiles();
				if (fileList != null) {
					for (int i = 0; i < fileList.length; i++) {
						updateFileDetails(fileList[i],reportList);
					}
				}
			} else {
				reportList.add(new ExcelReportVO(file, file.getName(), file.getAbsolutePath(), file.length()));
			}
		}
		
		return reportList;
	}

}
