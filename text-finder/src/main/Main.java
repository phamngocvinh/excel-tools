package main;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {

	/**
	 * Application Version
	 */
	private static final String appVersion = "0.1";

	/**
	 * Proerties
	 */
	private static Properties prop = new Properties();

	/**
	 * Logger
	 */
	private static final Logger logger = LogManager.getLogger(Main.class);

	/**
	 * Current path
	 */
	private static final String dir = System.getProperty("user.dir");

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
	private static List<String> config_SrchCond = new ArrayList<>();

	/**
	 * Main class
	 * 
	 * @param args
	 */
	public static void main(String[] args) {

		// Get config file
		try {
			FileInputStream fis = new FileInputStream(new File(dir + "\\" + name_config));
			prop.load(fis);
		} catch (FileNotFoundException e) {
			logger.error("config.properties not found");
			return;
		} catch (IOException e) {
			logger.error("config.properties not found");
			return;
		}

		// Check if config file valid
		if (!isValidConfig()) {
			return;
		}

		// Read config file
		try {
			readConfig();
		} catch (Exception e) {
			logger.error(e.getMessage());
			return;
		}
	}

	/**
	 * Read Config File to get settings
	 * 
	 * @throws Exception
	 */
	private static void readConfig() throws Exception {

		DataFormatter formatter = new DataFormatter();

		// Check if config excel exists
		File fileWorkBook = new File(dir + prop.getProperty("config.file"));
		if (fileWorkBook.exists()) {
			logger.info("Read config.xlsx");
		} else {
			logger.error("Config file not exists");
			return;
		}

		// Open config file
		XSSFWorkbook workbook = new XSSFWorkbook(fileWorkBook);

		// Open config sheet
		XSSFSheet sheet = workbook.getSheet(prop.getProperty("config.sheet"));

		// Check if config sheet exists
		if (sheet == null) {
			workbook.close();
			throw new Exception("Config sheet not found");
		}

		// Get search path
		int col_Path = Integer.parseInt(getProp("config.search.path.col")) - 1;
		int row_Path = Integer.parseInt(getProp("config.search.path.row")) - 1;
		config_search_path = formatter.formatCellValue(sheet.getRow(row_Path).getCell(col_Path));

		// Check if search path is file or folder
		if (Files.isDirectory(Path.of(config_search_path))) {
			config_IsFolder = true;
			logger.info("Search in " + config_search_path);
		} else if (Files.isRegularFile(Path.of(config_search_path))) {
			config_IsFolder = false;
			logger.info("Search in " + config_search_path);
		} else {
			logger.error("Search path not exist. Check config.xlsx");
			workbook.close();
			return;
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
	}

	private static boolean isValidConfig() {

		// Check if correctly config
		if (Common.isNullOrEmpty(getProp("app.version"))) {
			logger.error("Config: Missing app.version property");
			return false;
		} else if (Common.isNullOrEmpty(getProp("config.file"))) {
			logger.error("Config: Missing app.version property");
			return false;
		} else if (Common.isNullOrEmpty(getProp("config.sheet"))) {
			logger.error("Config: Missing app.version property");
			return false;
		} else if (Common.isNullOrEmpty(getProp("config.search.path.col"))) {
			logger.error("Config: Missing config.search.path.col property");
			return false;
		} else if (Common.isNullOrEmpty(getProp("config.search.path.row"))) {
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
