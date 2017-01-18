package com.infy.gcoe.poi.base;

import java.util.List;

import com.infy.gcoe.poi.vo.ExcelReportVO;

public interface IReportBuilder {
	
	public List<ExcelReportVO> update(List<ExcelReportVO> reportList) throws Exception;
	
}
