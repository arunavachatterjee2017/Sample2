package sample;

import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.BatchUpdateSpreadsheetRequest;
import com.google.api.services.sheets.v4.model.BatchUpdateSpreadsheetResponse;
import com.google.api.services.sheets.v4.model.FindReplaceRequest;
import com.google.api.services.sheets.v4.model.FindReplaceResponse;
import com.google.api.services.sheets.v4.model.Request;
import com.google.api.services.sheets.v4.model.SpreadsheetProperties;
import com.google.api.services.sheets.v4.model.UpdateSpreadsheetPropertiesRequest;
import com.google.api.services.sheets.v4.model.ValueRange;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;


public class ReadWriteGoogleSheet {

	// https://developers.google.com/sheets/api/quickstart/java#prerequisites
		private final String APPLICATION_NAME = "Google Sheets API Java Quickstart";
		private final JsonFactory JSON_FACTORY = JacksonFactory.getDefaultInstance();
		private final String TOKENS_DIRECTORY_PATH = System.getProperty("user.dir")
				+ "/src/test/java/sample";
		private final List<String> SCOPES = Collections.singletonList(SheetsScopes.SPREADSHEETS);
		//private static final List<String> SCOPES = Arrays.asList("https://spreadsheets.google.com/feeds","https://www.googleapis.com/auth/drive");
		private final String CREDENTIALS_FILE_PATH = System.getProperty("user.dir")
				+ "/src/test/java/sample/credentials.json";
		final String range = "A1:B";
		private Sheets service;
		private final String title = "DemoQA";
		
		public String getData(String spreadsheetId, String Columnname) {
			String name="";
			try {
				
				final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
				service = new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, getCredentials(HTTP_TRANSPORT))
						.setApplicationName(APPLICATION_NAME).build();
				ValueRange response = service.spreadsheets().values().get(spreadsheetId, range).execute();
				List<List<Object>> values = response.getValues();
				if (values == null || values.isEmpty()) {
					System.out.println("No data found.");
				} else {
					// Read Data
					for (List row : values) {
						// Print columns A and E, which correspond to indices 0 and 4.
						System.out.println(values.size());
						System.out.printf("%s, %s\n", row.get(0), row.get(1));
					}
					
					for(int i=0; i< values.size(); i++) {
						if(values.get(i).get(0).toString() != Columnname) {
							name = values.get(i).get(0).toString();
						}
					}
				}
			}
			catch(Exception e) {
				System.out.println(e.getMessage());
			}
			
			return name;
		}
		
		public BatchUpdateSpreadsheetResponse updateData(String spreadsheetId, String find,String replacement) throws IOException {
			BatchUpdateSpreadsheetResponse response = null;
			try {
				final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
				service = new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, getCredentials(HTTP_TRANSPORT))
						.setApplicationName(APPLICATION_NAME).build();
				// [START sheets_batch_update]
				List<Request> requests = new ArrayList<>();
				// Change the spreadsheet's title.
				requests.add(new Request().setUpdateSpreadsheetProperties(new UpdateSpreadsheetPropertiesRequest()
						.setProperties(new SpreadsheetProperties().setTitle(title)).setFields("title")));
				// Find and replace text.
				requests.add(new Request()
						.setFindReplace(new FindReplaceRequest().setFind(find).setReplacement(replacement).setAllSheets(true)));
				// Add additional requests (operations) ...

				BatchUpdateSpreadsheetRequest body = new BatchUpdateSpreadsheetRequest().setRequests(requests);
				response = service.spreadsheets().batchUpdate(spreadsheetId, body).execute();
				FindReplaceResponse findReplaceResponse = response.getReplies().get(1).getFindReplace();
				System.out.printf("%d replacements made.", findReplaceResponse.getOccurrencesChanged());
				// [END sheets_batch_update]
			}
			catch(Exception e) {
				System.out.println(e.getMessage());
			}
			return response;
		}

		private Credential getCredentials(final NetHttpTransport HTTP_TRANSPORT) throws IOException {
			// Load client secrets.
			File initialFile = new File(CREDENTIALS_FILE_PATH);
			InputStream in = new FileInputStream(initialFile);
			if (in == null) {
				throw new FileNotFoundException("Resource not found: " + CREDENTIALS_FILE_PATH);
			}
			GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(in));

			// Build flow and trigger user authorization request.
			GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(HTTP_TRANSPORT, JSON_FACTORY,
					clientSecrets, SCOPES)
							.setDataStoreFactory(new FileDataStoreFactory(new java.io.File(TOKENS_DIRECTORY_PATH)))
							.setAccessType("offline").build();
			LocalServerReceiver receiver = new LocalServerReceiver.Builder().setPort(8888).build();
			return new AuthorizationCodeInstalledApp(flow, receiver).authorize("user");
		}
}
