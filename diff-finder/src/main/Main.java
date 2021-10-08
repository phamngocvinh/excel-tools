package main;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.net.HttpURLConnection;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.LinkedList;
import java.util.List;
import java.util.Properties;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.commons.lang3.StringUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {

	/**
	 * Application Version
	 */
	private static String appVersion = "";

	/**
	 * Properties
	 */
	private static Properties prop = new Properties();

	/**
	 * Logger
	 */
	private static final Logger logger = LogManager.getLogger(Main.class);

	/**
	 * Logger
	 */
	private static final int LOG_NUM = 50;

	/**
	 * Current path
	 */
	private static final String dir = System.getProperty("user.dir");

	/**
	 * Separator
	 */
	private static final String SEP = "&%&";

	/**
	 * List Header
	 */
	private static final List<String> listHeader = Arrays.asList("Text    ", "Location    ", "Sheet    ",
			"Filename    ", "Path    ");

	/**
	 * Name: Config file name
	 */
	private static String name_config = "config.properties";

	/**
	 * Config: Search path
	 */
	private static String config_search_path = "";

	/**
	 * Config: Is Search in Folder
	 */
	private static boolean config_IsFolder = false;

	/**
	 * Config: Search Conditions
	 */
	private static List<String> config_SrchCond = new LinkedList<>();

	/**
	 * Result workbook
	 */
	private static XSSFWorkbook wb_Result;

	/**
	 * DataFormatter
	 */
	private static DataFormatter formatter = new DataFormatter();

	/**
	 * List Result
	 */
	private static List<String> listResult = new LinkedList<>();

	/**
	 * Main class
	 * 
	 * @param args
	 */
	public static void main(String[] args) {

		try {
			logger.info(StringUtils.rightPad("", 90, "="));

			// Get config file
			FileInputStream fis = new FileInputStream(new File(dir + "\\" + name_config));
			prop.load(fis);

			// Get app version from properties
			appVersion = getProp("app.version");

			// Check for new version
			checkVersion();

			logger.info(StringUtils.rightPad("=== Diff Finder v" + appVersion + " ", LOG_NUM, "="));
			logger.info(StringUtils.rightPad("=== START =========", LOG_NUM, "="));

			// Check if config file valid
			if (!isValidConfig()) {
				return;
			}

			// Read config file
			try {
				if (!readConfig()) {
					return;
				}
			} catch (InvalidFormatException e1) {
				logger.error("InvalidFormatException: Config file");
				return;
			} catch (IOException e1) {
				logger.error("IOException: Config file");
				return;
			}

			// Initialize Result workbook
			wb_Result = new XSSFWorkbook();
			wb_Result.createSheet("Result");

			// Write Search Result
			writeResult();

			// Push result to OutputStream
			FileOutputStream outputStream;
			try {
				outputStream = new FileOutputStream("Result.xlsx");
				wb_Result.write(outputStream);
			} catch (FileNotFoundException e) {
				logger.error("FileNotFoundException: Write Result");
			} catch (IOException e) {
				logger.error("IOException: Write Result");
			}

		} catch (FileNotFoundException e) {
			logger.error("FileNotFoundException: config.properties");
			return;
		} catch (IOException e) {
			logger.error("IOException: config.properties");
			return;
		} catch (Exception ex) {
			logger.error("Internal Exception: " + ex.getLocalizedMessage());
		} finally {
			logger.info(StringUtils.rightPad("=== END =========", LOG_NUM, "="));
		}
	}

	// Check for newer version
	private static void checkVersion() {
		
		try {
			// Get latest version
			URL url = new URL("https://api.github.com/repos/phamngocvinh/excel-tools/releases");

			BufferedReader in = new BufferedReader(new InputStreamReader(url.openStream()));

			String line = in.readLine();
			in.close();

			// Get release version
			Pattern p = Pattern.compile(".+?name.+?diff-finder-v(.+).zip.+");
			Matcher m = p.matcher(line);
			boolean isMatch = m.matches();

			if (isMatch) {
				String netVersion = m.group(1).substring(0, m.group(1).indexOf(".zip"));

				// If local version is older than newest version
				if (netVersion.compareTo(appVersion) > 0) {
					logger.warn("You're using older version. Please update to the latest version in the link below");
					logger.info("Current: v" + appVersion);
					logger.info("Latest: v" + netVersion);
					logger.info(
							"Official Link: https://github.com/phamngocvinh/excel-tools/releases/tag/v" + netVersion);
					logger.info(StringUtils.rightPad("", 90, "="));
				}
			}
		} catch (Exception e) {
			logger.error("Internal Exception: " + e.getLocalizedMessage());
		} 
	}

	/**
	 * Write Result to Workbook
	 */
	private static void writeResult() {

		// Get Result Sheet
		XSSFSheet sheet = wb_Result.getSheet("Result");

		// Write Headers
		XSSFCellStyle cellStyle = wb_Result.createCellStyle();
		Font headerFont = wb_Result.createFont();
		headerFont.setBold(true);
		cellStyle.setFont(headerFont);

		sheet.createRow(0);
		for (int idx = 0; idx < listHeader.size(); idx++) {
			sheet.getRow(0).createCell(idx).setCellValue(listHeader.get(idx));
			sheet.getRow(0).getCell(idx).setCellStyle(cellStyle);
		}

		// Loop through all results to write
		for (int idx = 0; idx < listResult.size(); idx++) {
		}

		// Set filter
		sheet.setAutoFilter(new CellRangeAddress(0, sheet.getLastRowNum(), 0, listHeader.size() - 1));

		// Set Column fit contents
		for (int idx = 0; idx < listHeader.size(); idx++) {
			sheet.autoSizeColumn(idx);
		}
	}

	/**
	 * Read Config File to get settings
	 * 
	 * @throws IOException
	 * @throws InvalidFormatException
	 * 
	 * @throws Exception
	 */
	private static boolean readConfig() throws InvalidFormatException, IOException {

		// Check if config excel exists
		File fileWorkBook = new File(dir + prop.getProperty("config.file"));
		if (fileWorkBook.exists()) {
			logger.info("Read config.xlsx");
		} else {
			logger.error("Config file not exists");
			return false;
		}

		// Open config file
		XSSFWorkbook workbook = new XSSFWorkbook(fileWorkBook);

		// Open config sheet
		XSSFSheet sheet = workbook.getSheet(prop.getProperty("config.sheet"));

		// Check if config sheet exists
		if (sheet == null) {
			logger.error("Config sheet not found");
			workbook.close();
			return false;
		}

		// Get search path
		int col_Path = Integer.parseInt(getProp("config.search.path.col")) - 1;
		int row_Path = Integer.parseInt(getProp("config.search.path.row")) - 1;
		config_search_path = formatter.formatCellValue(sheet.getRow(row_Path).getCell(col_Path));

		// Check if search path is file or folder
		if (Files.isDirectory(Path.of(config_search_path))) {
			config_IsFolder = true;
			if (Path.of(config_search_path).toFile().listFiles().length == 0) {
				logger.info("Search path is empty");
				workbook.close();
				return false;
			}
			logger.info("Search in " + config_search_path);
		} else if (Files.isRegularFile(Path.of(config_search_path))) {
			config_IsFolder = false;
			logger.info("Search in " + config_search_path);
		} else {
			logger.error("Search path not exist. Check config.xlsx");
			workbook.close();
			return false;
		}

		// Get Search Condition
		int col_Cond = Integer.parseInt(getProp("config.search.cond.col")) - 1;
		int row_Cond = Integer.parseInt(getProp("config.search.cond.row")) - 1;
		int row_Last = sheet.getLastRowNum();

		for (int idx = row_Cond; idx <= row_Last; idx++) {
			config_SrchCond.add(formatter.formatCellValue(sheet.getRow(idx).getCell(col_Cond)));
		}
		logger.info("Search condition: " + config_SrchCond);

		workbook.close();

		return true;
	}

	/**
	 * Check if Config file is valid
	 * 
	 * @return
	 */
	private static boolean isValidConfig() {

		// Check if correctly config
		if (StringUtils.isBlank(getProp("app.version"))) {
			logger.error("Config: Missing app.version property");
			return false;
		} else if (StringUtils.isBlank(getProp("config.file"))) {
			logger.error("Config: Missing app.version property");
			return false;
		} else if (StringUtils.isBlank(getProp("config.sheet"))) {
			logger.error("Config: Missing app.version property");
			return false;
		} else if (StringUtils.isBlank(getProp("config.search.path.col"))) {
			logger.error("Config: Missing config.search.path.col property");
			return false;
		} else if (StringUtils.isBlank(getProp("config.search.path.row"))) {
			logger.error("Config: Missing config.search.path.row property");
			return false;
		}

		// If config version different from app version, throw error
		if (!appVersion.equals(getProp("app.version"))) {
			logger.error("App Version do not match Config Version");
			return false;
		}

		return true;
	}

	/**
	 * Get Property Value
	 * 
	 * @param key Key
	 * @return Property Value
	 */
	private static String getProp(String key) {
		return prop.getProperty(key);
	}
}
