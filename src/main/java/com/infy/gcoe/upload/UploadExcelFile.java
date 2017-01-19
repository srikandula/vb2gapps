package com.infy.gcoe.upload;

import java.util.Arrays;
import java.util.List;
import java.util.Set;

import org.springframework.boot.ApplicationArguments;
import org.springframework.boot.CommandLineRunner;
import org.springframework.context.annotation.Profile;
import org.springframework.stereotype.Component;

import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;

import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.FileContent;
import com.google.api.client.http.HttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.drive.Drive;
import com.google.api.services.drive.DriveScopes;

@Component
@Profile(value="UploadExcelFile")
public class UploadExcelFile implements CommandLineRunner {
	
	/** Application name. */
    private static final String APPLICATION_NAME = "Upload Excel Sheet to Google Drive";

    /** Directory to store user credentials for this application. */
    private static final java.io.File DATA_STORE_DIR = new java.io.File(".");

    /** Global instance of the {@link FileDataStoreFactory}. */
    private static FileDataStoreFactory DATA_STORE_FACTORY_FILE;

    /** Global instance of the JSON factory. */
    private static final JsonFactory JSON_FACTORY = JacksonFactory.getDefaultInstance();

    /** Global instance of the HTTP transport. */
    private static HttpTransport HTTP_TRANSPORT;

    /** Global instance of the scopes required by this quickstart.
     *
     * If modifying these scopes, delete your previously saved credentials
     * at ~/.credentials/drive-java-quickstart
     */
    private static final List<String> SCOPES = Arrays.asList(DriveScopes.all().toArray(new String[0]));

    private List<String> folderPath = null;
    
    static {
        try {
        	System.out.println("In static block.........");
            HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
            DATA_STORE_FACTORY_FILE = new FileDataStoreFactory(DATA_STORE_DIR);
        } catch (Throwable t) {
        	System.out.println("In static block exception.........");
            t.printStackTrace();
        }
    }

    /**
     * Creates an authorized Credential object.
     * @return an authorized Credential object.
     * @throws IOException
     */
    public static Credential authorize() throws Exception {
    	System.out.println("About to start authorize()");
        // Load client secrets.
    	ClassLoader classLoader = Thread.currentThread().getContextClassLoader();
        
    	InputStream in = classLoader.getResourceAsStream("client_secrets.json");        
        System.out.println("Reading the file with access informaiton (available) = " + in.available());
        
        GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(in));
        System.out.println("Secrets loaded into memory");
        
        // Build flow and trigger user authorization request.
        GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(HTTP_TRANSPORT, JSON_FACTORY, clientSecrets, SCOPES)
                .setDataStoreFactory(DATA_STORE_FACTORY_FILE)
                .setAccessType("offline")
                .build();

        
        System.out.println("GoogleAuthorizationCodeFlow........flow creation complete");
        Credential credential = new AuthorizationCodeInstalledApp(flow, new LocalServerReceiver()).authorize("user");
        
        System.out.println("Credentials saved to " + DATA_STORE_DIR.getAbsolutePath());
        return credential;
    }	

    /**
     * Build and return an authorized Drive client service.
     * @return an authorized Drive client service
     * @throws IOException
     */
    public static Drive getDriveService() throws Exception {
    	System.out.println("About to start getDriveService()");
        Credential credential = authorize();
        return new Drive.Builder(
                HTTP_TRANSPORT, JSON_FACTORY, credential)
                .setApplicationName(APPLICATION_NAME) 
                .build();
    }


    
	public UploadExcelFile(ApplicationArguments args){
		System.out.println("Inside constructor UploadExcelFile...........");
		
		boolean path = args.containsOption("path");
		System.out.println("Has path Variable " + path);
		
        List<String> files = args.getNonOptionArgs();
        for(String name:files){
        	System.out.println(name);
        }

        Set<String> optionNames = args.getOptionNames();
        for(String name:optionNames){
        	System.out.println(name);
        }
        
        folderPath = args.getOptionValues("path");
	}

	@Override
	public void run(String[] args) throws Exception {
		System.out.println("Inside run method of UploadExcelFile...........");
		if(args != null){
			System.out.println("Inside UploadExcelFile run() : ");
			for(int i=0;i<args.length;i++){
				System.out.println(args[i]);
			}
		}
		
		String fileName = "";
        String extn = "";
        String mimeType = "";
        
    	// Build a new authorized API client service.
        Drive service = getDriveService();
        
        //Fetch files from the given folder
        java.io.File[] listOfFiles = fetchFilesFromFolder(folderPath);
        
        for (java.io.File fileRef : listOfFiles) {
            if (fileRef.isFile()) {
            	fileName =  fileRef.getName();
            	extn     =  fileName.substring(fileName.lastIndexOf(".") + 1);
            	mimeType = MimeTypes.getMimeType(extn);
            	com.google.api.services.drive.model.File file = new com.google.api.services.drive.model.File();
            	file.setTitle(fileName);
        		file.setMimeType(mimeType);
        		// File's content.
        		java.io.File fileContent = new java.io.File(folderPath.get(0).substring(0,folderPath.get(0).indexOf("\""))+"\\"+fileName);
        		FileContent mediaContent = new FileContent(mimeType,fileContent);
        		file = service.files().insert(file, mediaContent).setConvert(true).execute();
        		System.out.println("file Name "+file.getTitle()+" file Id "+file.getId());
            }
        }
    
		
	}
	
	
	private java.io.File[] fetchFilesFromFolder(List<String> folderPath) {
		System.out.println("folderPath.get(0) "+folderPath.get(0).substring(0,folderPath.get(0).indexOf("\"")));
		java.io.File folder = new java.io.File(folderPath.get(0).substring(0,folderPath.get(0).indexOf("\"")));
		if(folder.listFiles()!=null && folder.listFiles().length>0){
			System.out.println("folder.listFiles() "+folder.listFiles().length);	
		}
		
        return folder.listFiles();
       	
	}

}
