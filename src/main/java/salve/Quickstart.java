package salve;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.temporal.TemporalAdjusters;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Date;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Properties;
import java.util.concurrent.TimeUnit;
import java.util.function.Predicate;
import java.util.stream.Collectors;

import org.apache.commons.lang3.StringUtils;

import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.HttpRequest;
import com.google.api.client.http.HttpRequestInitializer;
import com.google.api.client.http.HttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.util.DateTime;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.calendar.CalendarScopes;
import com.google.api.services.calendar.model.Event;
import com.google.api.services.calendar.model.Events;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.ValueRange;

public class Quickstart {
	// https://developers.google.com/google-apps/calendar/quickstart/java
	// https://developers.google.com/sheets/api/quickstart/java

	/** Application name. */
	private static final String APPLICATION_NAME_CALENDAR = "Google Calendar Salve Salve";

	private static final String APPLICATION_NAME_SHEET = "Google Sheet Salve Salve";

	/** Directory to store user credentials for this application. */
	private static final java.io.File DATA_STORE_DIR_CALENDAR = new java.io.File(System.getProperty("user.home"),
			".credentials/calendar");

	private static final java.io.File DATA_STORE_DIR_SHEET = new java.io.File(System.getProperty("user.home"),
			".credentials/sheets");

	/** Global instance of the {@link FileDataStoreFactory}. */
	private static FileDataStoreFactory DATA_STORE_FACTORY_CALENDAR;
	private static FileDataStoreFactory DATA_STORE_FACTORY_SHEET;

	/** Global instance of the JSON factory. */
	private static final JsonFactory JSON_FACTORY = JacksonFactory.getDefaultInstance();

	/** Global instance of the HTTP transport. */
	private static HttpTransport HTTP_TRANSPORT;

	/**
	 * Global instance of the scopes required by this quickstart.
	 *
	 * If modifying these scopes, delete your previously saved credentials at
	 * ~/.credentials/calendar-java-quickstart
	 */
	private static final List<String> SCOPES_CALENDAR = Arrays.asList(CalendarScopes.CALENDAR_READONLY);

	private static final List<String> SCOPES_SHEET = Arrays.asList(SheetsScopes.SPREADSHEETS);

	static {
		try {
			HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
			DATA_STORE_FACTORY_CALENDAR = new FileDataStoreFactory(DATA_STORE_DIR_CALENDAR);
			DATA_STORE_FACTORY_SHEET = new FileDataStoreFactory(DATA_STORE_DIR_SHEET);
		} catch (Throwable t) {
			t.printStackTrace();
			System.exit(1);
		}
	}

	/**
	 * Creates an authorized Credential object.
	 * 
	 * @return an authorized Credential object.
	 * @throws IOException
	 */
	public static Credential authorizeCalendar(String secretCalendarPath) throws IOException {
		InputStream in = null;
		Credential credential;
		try {
			in = new FileInputStream(new File(secretCalendarPath));
			GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(in));

			// Build flow and trigger user authorization request.
			GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(HTTP_TRANSPORT, JSON_FACTORY,
					clientSecrets, SCOPES_CALENDAR).setDataStoreFactory(DATA_STORE_FACTORY_CALENDAR)
							.setAccessType("offline").build();
			credential = new AuthorizationCodeInstalledApp(flow, new LocalServerReceiver())
					.authorize("user");
			System.out.println("Credentials saved to " + DATA_STORE_DIR_CALENDAR.getAbsolutePath());
		} finally {
			if (in != null)
				in.close();
		}

		return credential;
	}

	public static Credential authorizeSheet(String secretSheetPath) throws IOException {
		
		InputStream in = null;
		Credential credential;
		
		try {
			in = new FileInputStream(new File(secretSheetPath));

			GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(in));

			// Build flow and trigger user authorization request.
			GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(HTTP_TRANSPORT, JSON_FACTORY,
					clientSecrets, SCOPES_SHEET).setDataStoreFactory(DATA_STORE_FACTORY_SHEET).setAccessType("offline")
							.build();
			credential = new AuthorizationCodeInstalledApp(flow, new LocalServerReceiver()).authorize("user");
			System.out.println("Credentials saved to " + DATA_STORE_DIR_SHEET.getAbsolutePath());
			
		} finally {
			if(in != null)
				in.close();
		}

		return credential;
	}

	/**
	 * Build and return an authorized Calendar client service.
	 * 
	 * @return an authorized Calendar client service
	 * @throws IOException
	 */
	public static com.google.api.services.calendar.Calendar getCalendarService(String secretCalendarPath)
			throws IOException {
		Credential credential = authorizeCalendar(secretCalendarPath);
		return new com.google.api.services.calendar.Calendar.Builder(HTTP_TRANSPORT, JSON_FACTORY, credential)
				.setApplicationName(APPLICATION_NAME_CALENDAR).build();
	}

	/**
	 * Build and return an authorized Sheets API client service.
	 * 
	 * @return an authorized Sheets API client service
	 * @throws IOException
	 */
	public static Sheets getSheetsService(String secretSheetPath) throws IOException {
		final Credential credential = authorizeSheet(secretSheetPath);
		return new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, credential)
				.setHttpRequestInitializer(new HttpRequestInitializer() {

					public void initialize(HttpRequest req) throws IOException {
						credential.initialize(req);
						req.setConnectTimeout(5 * 60000);
						req.setReadTimeout(5 * 60000);

					}
				}).setApplicationName(APPLICATION_NAME_SHEET).build();
	}

	public static void main(String[] args) throws Exception {

		Properties prop = loadProperties("./config.properties");

		String secretSheetPath = prop.getProperty("secret_calendar", "./secret_calendar.json");
		String secretCalendarPath = prop.getProperty("secret_sheet", "./secret_sheet.json");

		// Build a new authorized API client service.
		Sheets serviceSheet = getSheetsService(secretSheetPath);

		// ID da planilha do Boris
		// String spreadsheetId = "1B0G1GgEFN6Uw7ZwYx5mQiLxq58BKDgPqjxEsZcCfPLc";
		String spreadsheetId = prop.getProperty("spreadsheetId");

		// Busca na planilha do Boris as semanas do quarter e os seus
		// respectivos dias
		// Nome da planilha e lista de colunas a serem buscadas.
		// String rangeQuarterWeek = "Listas!F1:G";
		String rangeQuarterWeek = prop.getProperty("rangeQuarterWeek", "Listas!F1:G");
		Map<String, Date> quarterWeek = getQuarterWeek(serviceSheet, spreadsheetId, rangeQuarterWeek);

		// Busca a lista de atividades padrões na planilha do Boris
		// Nome da planilha e lista de colunas a serem buscadas.
		// String rangeListaAtividades = "Listas!A:A";
		String rangeListaAtividades = prop.getProperty("rangeListaAtividades", "Listas!A:A");
		Map<String, String> atividades = getAtividades(serviceSheet, spreadsheetId, rangeListaAtividades);
		
		String defaultAccountType = prop.getProperty("defaultAccountType", "GV");
		final String defaultAccountTypeFinal;
		
		if (defaultAccountType == null || defaultAccountType.isEmpty()) {
			System.out.println("Campo 'defaultAccountType' no config.properties deve ser preenchido!");
			System.exit(1);
		} else {
			defaultAccountType = defaultAccountType.replaceAll("\"", "");
		}
		
		defaultAccountTypeFinal = defaultAccountType;
		
		// Busca na planilha do Boris os tipos de clientes
		// Nome da planilha e lista de colunas a serem buscadas.
		// String rangeTipoCliente = "Listas!J:J";
		// List<String> tipoCliente = getTipoCliente(serviceSheet, spreadsheetId,
		// rangeTipoCliente);

		String user = prop.getProperty("user");

		if (user == null || user.isEmpty()) {
			System.out.println("Campo 'user' no config.properties deve ser preenchido!");
			System.exit(1);
		} else {
			user = user.replaceAll("\"", "");
		}

		System.out.println("User: " + user);

		// Busca a ultima semana preenchida e retorna a proxima da lista
		String userRangeWeeks = user + prop.getProperty("userRangeWeeks", "!A2:A");
		String nextWeek = getNextWeek(serviceSheet, spreadsheetId, userRangeWeeks, quarterWeek);

		System.out.println("Next week to be fulfilled is: " + nextWeek);

		// Data de inicio do calendario
		DateTime inicio = new DateTime(quarterWeek.get(nextWeek));

		// Build a new authorized API client service.
		// Note: Do not confuse this class with the
		// com.google.api.services.calendar.model.Calendar class.
		com.google.api.services.calendar.Calendar service = getCalendarService(secretCalendarPath);

		//
		DateTime lastSunday = new DateTime(Date.from(LocalDate.now().with(TemporalAdjusters.previous(DayOfWeek.SUNDAY))
				.atStartOfDay(ZoneId.systemDefault()).toInstant()));

		Events events = service.events().list("primary").setMaxResults(500).setTimeMin(inicio).setTimeMax(lastSunday)
				.setOrderBy("startTime").setSingleEvents(true).execute();

		List<Event> items = events.getItems();

		List<Evento> eventos = null;

		if (items.size() == 0) {
			System.out.println("No upcoming events found.");
			System.exit(0);
		} else {
			eventos = items.stream().filter(summaryNotNull().and(summaryContainsHash()))
					.map(event -> createEvento(event, atividades, quarterWeek, defaultAccountTypeFinal)).collect(Collectors.toList());
		}

		writeSheet(eventos, spreadsheetId, serviceSheet,
				getNextAvaiableRow(serviceSheet, spreadsheetId, userRangeWeeks), user);
	}

	private static Properties loadProperties(String filePath) {
		Properties prop = new Properties();
		InputStream input = null;

		try {

			input = new FileInputStream(filePath);

			// load a properties file
			prop.load(input);

			return prop;

		} catch (IOException ex) {
			ex.printStackTrace();
			System.exit(1);
		} finally {
			if (input != null) {
				try {
					input.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}

		return null;
	}

	private static Predicate<Event> summaryNotNull() {
		return it -> it.getSummary() != null;
	}

	private static Predicate<Event> summaryContainsHash() {
		return it -> it.getSummary().contains("#");
	}

	private static Evento createEvento(Event event, Map<String, String> atividades, Map<String, Date> quarterDia, String defaultAccountType) {
		Evento evento = new Evento();
		String summary = event.getSummary();
		DateTime start = event.getStart().getDateTime();
		DateTime end = event.getEnd().getDateTime();
		if (start == null) {
			start = event.getStart().getDate();
		}
		if (end == null) {
			end = event.getEnd().getDate();
		}

		// Preenche effort
		Date inicioTime = new Date(start.getValue());
		Date fimTime = new Date(end.getValue());
		evento.setEffort(getDateDiff(inicioTime, fimTime, TimeUnit.MINUTES) / 60.0);

		// Preenche description, account type e customer

		String description = event.getDescription();

		if (summary.contains("[") || summary.contains("]")) {
			evento.setDescription(StringUtils.substringAfter(summary, "]").trim());

			String subSummary = StringUtils.substringBetween(summary, "[", "]");

			if (subSummary.contains("-")) {
				evento.setAccountType(StringUtils.substringBetween(summary, "-", "]"));
				evento.setCustomer(StringUtils.substringBetween(summary, "[", "-"));
			} else {
				evento.setAccountType(defaultAccountType);
				evento.setCustomer(subSummary);
			}
		} else {
			evento.setDescription(StringUtils.substringAfter(summary, "#").trim());
		}

		// Preenche type
		evento.setType(atividades.get(StringUtils.substringBefore(summary, "#").toUpperCase()));

		evento.setWeek(getWeek(inicioTime, quarterDia));

		if (description != null) {
			if (description.toLowerCase().contains("#city:")) {
				evento.setCity(StringUtils.substringBetween(description, "#city:", "#").trim());
			}

			if (description.toLowerCase().contains("#partner:")) {
				evento.setPartner(StringUtils.substringBetween(description, "#partner:", "#").trim());
				if (description.toLowerCase().contains("#presale"))
					evento.setPreSalePartner("Y");
				else
					evento.setPreSalePartner("N");
			}
		}
		return evento;
	}

	// private static void updateSheets(List<Evento> eventos, String sheetId,
	// Sheets serviceSheet) throws IOException
	// {
	//
	// String range = "Gustavo Luszczynski!A15:E18";
	// ValueRange response = serviceSheet.spreadsheets().values()
	// .get(sheetId, range)
	// .execute();
	// List<List<Object>> values = response.getValues();
	//
	// for (List row : values)
	// {
	// row.set(0, "oi"); //a string
	// row.set(1, "teste"); //a string
	// row.set(2, "olha"); //a string
	// }
	// }

	private static void writeSheet(List<Evento> eventos, String sheetId, Sheets serviceSheet, Integer row,
			String user) {
		try {

			if (eventos.size() != 0) {
				String writeRange = user + "!A" + row + ":I";

				List<List<Object>> writeData = new ArrayList<List<Object>>();

				for (Evento evento : eventos) {
					List<Object> dataRow = new ArrayList<Object>();
					dataRow.add(evento.getWeek());
					dataRow.add(evento.getType());
					dataRow.add(evento.getCustomer());
					dataRow.add(evento.getAccountType());
					dataRow.add(evento.getDescription());
					dataRow.add(evento.getEffort());
					dataRow.add(evento.getPartner());
					dataRow.add(evento.getPreSalePartner());
					dataRow.add(evento.getCity());
					writeData.add(dataRow);
				}

				ValueRange vr = new ValueRange().setValues(writeData).setMajorDimension("ROWS");
				serviceSheet.spreadsheets().values().update(sheetId, writeRange, vr).setValueInputOption("RAW")
						.execute();
			}
		} catch (Exception e) {
			// handle exception
			e.printStackTrace();
		}
	}

	public static long getDateDiff(Date date1, Date date2, TimeUnit timeUnit) {
		long diffInMillies = date2.getTime() - date1.getTime();
		return timeUnit.convert(diffInMillies, TimeUnit.MILLISECONDS);
	}

	private static String getNextWeek(Sheets serviceSheet, String spreadsheetId, String rangeWeek,
			Map<String, Date> listaWeek) throws IOException {
		ValueRange response;
		List<List<Object>> values;
		// Executa a busca
		response = serviceSheet.spreadsheets().values().get(spreadsheetId, rangeWeek).execute();

		values = response.getValues();

		String lastWeek = null;

		if (values == null || values.size() == 0) {
			return "Q1W01";
		} else {
			for (int i = 0; i < values.size(); i++) {
				if (values.get(i).size() != 0) {
					lastWeek = values.get(i).get(0).toString();
				}
			}
		}

		Iterator<Entry<String, Date>> it = listaWeek.entrySet().iterator();

		while (it.hasNext()) {
			if (it.next().getKey().equals(lastWeek))
				return it.next().getKey();
		}

		return null;
	}

	private static Integer getNextAvaiableRow(Sheets serviceSheet, String spreadsheetId, String rangeWeek)
			throws IOException {
		ValueRange response;
		Integer intReturn;
	
		// Executa a busca
		response = serviceSheet.spreadsheets().values().get(spreadsheetId, rangeWeek).execute();
		
		List<List<Object>> values = response.getValues();
		
		if(values != null) {
			intReturn = values.size() + 2;
		}
		else {
			intReturn = 2;
		}
		
		return intReturn;
	}

	private static String getWeek(Date current, Map<String, Date> quarterWeek) {

		for (Map.Entry<String, Date> entry : quarterWeek.entrySet()) {
			String key = entry.getKey();
			Date value = entry.getValue();

			Calendar cal = Calendar.getInstance();
			cal.setTime(value);
			cal.add(Calendar.DATE, 7);

			// DateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");
			// System.out.println("Valor original: " +
			// dateFormat.format(value));
			// System.out.println("Valor adiantado: " +
			// dateFormat.format(cal.getTime()));

			if (isDateBetween(value, cal.getTime(), current)) {
				return key;
			}
		}

		return null;
	}

	private static List<String> getTipoCliente(Sheets serviceSheet, String spreadsheetId, String rangeTipoCliente)
			throws IOException {
		ValueRange response;
		List<List<Object>> values;
		// Executa a busca
		response = serviceSheet.spreadsheets().values().get(spreadsheetId, rangeTipoCliente).execute();

		values = response.getValues();

		List<String> tipoCliente = null;
		if (values == null || values.size() == 0) {
			System.out.println("No data found.");
		} else {
			tipoCliente = new ArrayList<String>();

			for (List row : values) {
				tipoCliente.add(row.get(0).toString());
			}
		}

		return tipoCliente;
	}

	private static Map<String, String> getAtividades(Sheets serviceSheet, String spreadsheetId,
			String rangeListaAtividades) throws IOException {
		ValueRange response;
		List<List<Object>> values;
		// Executa a busca
		response = serviceSheet.spreadsheets().values().get(spreadsheetId, rangeListaAtividades).execute();

		values = response.getValues();

		Map<String, String> atividades = null;
		if (values == null || values.size() == 0) {
			System.out.println("No data found.");
		} else {
			atividades = new LinkedHashMap<String, String>();

			for (List row : values) {
				if (row.get(0).toString().contains("Translado para outra cidade"))
					atividades.put("TRA", row.get(0).toString());
				else if (row.get(0).toString().contains("Horas em contato direto com clientes"))
					atividades.put("CLI", row.get(0).toString());
				else if (row.get(0).toString().contains("Horas utilizadas em POC"))
					atividades.put("POC", row.get(0).toString());
				else if (row.get(0).toString().contains("Horas em contato com Parceiros, Canais, SIs"))
					atividades.put("PAR", row.get(0).toString());
				else if (row.get(0).toString().contains("Horas usadas em estudos"))
					atividades.put("EST", row.get(0).toString());
				else if (row.get(0).toString().contains("Horas usadas em atividades de pré-vendas"))
					atividades.put("PRE", row.get(0).toString());
				else if (row.get(0).toString().contains("Horas usadas em atividades de pós-vendas"))
					atividades.put("POS", row.get(0).toString());
				else if (row.get(0).toString().contains("Trabalho de Backoffice"))
					atividades.put("BAC", row.get(0).toString());
				else if (row.get(0).toString().contains("Férias/PTO"))
					atividades.put("PTO", row.get(0).toString());

			}
		}

		return atividades;
	}

	private static Map<String, Date> getQuarterWeek(Sheets serviceSheet, String spreadsheetId, String rangeQuarterWeek)
			throws IOException, ParseException {
		// Executa a busca
		ValueRange response = serviceSheet.spreadsheets().values().get(spreadsheetId, rangeQuarterWeek).execute();

		List<List<Object>> values = response.getValues();

		Map<String, Date> quarterDia = null;
		if (values == null || values.size() == 0) {
			System.out.println("No data found.");
		} else {
			quarterDia = new LinkedHashMap<String, Date>();

			SimpleDateFormat simpleDateFormat = new SimpleDateFormat("dd-MMM-yyyy");

			for (List row : values) {
				quarterDia.put(row.get(0).toString(), simpleDateFormat.parse(row.get(1).toString()));
			}
		}

		return quarterDia;
	}

	private static Boolean isDateBetween(Date min, Date max, Date date) {
		return min.compareTo(date) * date.compareTo(max) > 0;
	}

}
