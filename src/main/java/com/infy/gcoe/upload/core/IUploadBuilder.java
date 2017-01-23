package com.infy.gcoe.upload.core;

import org.springframework.stereotype.Service;

import com.infy.gcoe.vo.ExcelReportVO;

@Service
public interface IUploadBuilder {
	
	public ExcelReportVO run(ExcelReportVO report) throws Exception;
	
}
