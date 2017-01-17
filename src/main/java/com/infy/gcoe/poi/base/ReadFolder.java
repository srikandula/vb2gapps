package com.infy.gcoe.poi.base;

import java.io.File;
import java.util.List;

import org.springframework.stereotype.Service;

import com.infy.gcoe.poi.vo.ExcelReportVO;

@Service
public class ReadFolder {

	public void read(File file, List<ExcelReportVO> report) throws Exception {
		
		if (file != null && file.exists()){
			if (file.isDirectory()) {
				File[] fileList = file.listFiles();
				if (fileList != null) {
					for (int i = 0; i < fileList.length; i++) {
						read(fileList[i],report);
					}
				}
			} else {
				report.add(new ExcelReportVO(file, file.getName(), file.getAbsolutePath(), file.length()));
			}
		}
	}
	
}
