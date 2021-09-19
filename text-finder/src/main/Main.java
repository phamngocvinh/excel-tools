package main;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Properties;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {

	/**
	 * Application Version
	 */
	private static final String appVersion = "0.1";

	/**
	 * Current path
	 */
	private static final String dir = System.getProperty("user.dir");

	/**
	 * Name: Config file name
	 */
	private static String name_config = "config.properties";

	/**
	 * Proerties
	 */
	private static Properties prop = new Properties();

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
			log("ERROR config.properties not found");
			return;
		} catch (IOException e) {
			log("ERROR config.properties not found");
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
			log(e.getMessage());
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
			log("INFO Found config file");
		} else {
			log("ERRO Config file not exists");
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

		int col = Integer.parseInt(getProp("config.search.path.col")) - 1;
		int row = Integer.parseInt(getProp("config.search.path.row")) - 1;
		System.out.println(formatter.formatCellValue(sheet.getRow(row).getCell(col)));

		workbook.close();
	}

	private static boolean isValidConfig() {

		// Check if correctly config
		if (Common.isNullOrEmpty(getProp("app.version"))) {
			log("ERROR Config: Missing app.version property");
			return false;
		} else if (Common.isNullOrEmpty(getProp("config.file"))) {
			log("ERROR Config: Missing config.file property");
			return false;
		} else if (Common.isNullOrEmpty(getProp("config.sheet"))) {
			log("ERROR Config: Missing config.sheet property");
			return false;
		} else if (Common.isNullOrEmpty(getProp("config.search.path.col"))) {
			log("ERROR Config: Missing config.search.path.col property");
			return false;
		} else if (Common.isNullOrEmpty(getProp("config.search.path.row"))) {
			log("ERROR Config: Missing config.search.path.row property");
			return false;
		}

		// If config version different from app version, throw error
		if (!appVersion.equals(getProp("app.version"))) {
			log("ERROR App Version do not match Config Version");
			return false;
		}

		return true;
	}

	/**
	 * Write console log
	 * @param msg Message
	 */
	private static void log(String msg) {
		Date date = new Date();
		SimpleDateFormat sdf = new SimpleDateFormat("yyyy/MM/dd hh:mm:ss");

		System.out.println(sdf.format(date) + " " + msg);
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
