package com.infy.gcoe.transform.core;

import com.infy.gcoe.vo.ExcelReportVO;

public interface ITransformBuilder {
	
	public ExcelReportVO run(ExcelReportVO report) throws Exception;
	
}
