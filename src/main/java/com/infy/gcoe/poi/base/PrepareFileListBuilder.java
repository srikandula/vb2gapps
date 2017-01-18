package com.infy.gcoe.poi.base;

import java.io.File;
import java.util.List;

import org.springframework.stereotype.Service;

import com.infy.gcoe.poi.vo.ExcelReportVO;

@Service
public class PrepareFileListBuilder {

	public List<ExcelReportVO> updateFileDetails(File file, List<ExcelReportVO> reportList) throws Exception {
		
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
