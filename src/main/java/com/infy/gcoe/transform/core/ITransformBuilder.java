package com.infy.gcoe.transform.core;

import java.util.List;

import com.infy.gcoe.poi.vo.ExcelReportVO;

public interface ITransformBuilder {
	
	public List<ExcelReportVO> run(List<ExcelReportVO> reportList) throws Exception;
	
}
