package com.infy.gcoe.upload.core;

import java.io.File;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import com.google.api.client.http.FileContent;
import com.google.api.services.drive.Drive;
import com.infy.gcoe.util.MimeTypes;
import com.infy.gcoe.vo.ExcelReportVO;

@Service
public class UploadFileToDrive implements IUploadBuilder {

	private static Logger logger = LoggerFactory.getLogger(CreateDriveManager.class);
	
	private Drive service = null; 
	
	public void setDrive(Drive service){
		this.service = service;
	}
	
	
	@Override
	public ExcelReportVO run(ExcelReportVO report) throws Exception {
		
		File fileContent = new File(report.getAbsolutePath());
		String mimeType = MimeTypes.getMimeType(report.getFileExtension());

		com.google.api.services.drive.model.File file = new com.google.api.services.drive.model.File();
		file.setTitle(report.getFileName());
		file.setMimeType(mimeType);

		FileContent mediaContent = new FileContent(mimeType, fileContent);
		file = service.files().insert(file, mediaContent).setConvert(true).execute();

		logger.debug("Successfully uploaded file Name " + file.getTitle() + " file Id " + file.getId());

		return report;
	}

}
