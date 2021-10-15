package main;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.LinkedList;
import java.util.List;
import java.util.Properties;

import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
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
	 * Latest Application Version
	 */
	private static String netVersion = "";

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
	 * Config: File 1 path
	 */
	private static String config_path_1 = "";

	/**
	 * Config: File 2 path
	 */
	private static String config_path_2 = "";

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
	 * Result workbook
	 */
	private static boolean isNewVersionExists = false;

	/**
	 * Main class
	 * 
	 * @param args
	 */
	public static void main(String[] args) {

		try {
			logger.info(StringUtils.rightPad("", LOG_NUM, "="));

			// Get config file
			FileInputStream fis = new FileInputStream(new File(dir + "\\" + name_config));
			prop.load(fis);

			// Get app version from properties
			appVersion = getProp("app.version");

			// Check for new version
			checkVersion();

			logger.info(StringUtils.rightPad("=== Diff Finder v" + appVersion + " ", LOG_NUM, "="));
			logger.info(StringUtils.rightPad("=== START ", LOG_NUM, "="));

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

			// Compare process
			doCompare();

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
			logger.info(StringUtils.rightPad("=== END ", LOG_NUM, "="));
			if (isNewVersionExists) {
				logger.info(StringUtils.rightPad("=== New Version Availiable ", LOG_NUM, "="));
				logger.warn("You're using older version. Please update to the latest version in the link below");
				logger.info("Current: v" + appVersion);
				logger.info("Latest: v" + netVersion);
				logger.info("Official Link: https://github.com/phamngocvinh/excel-tools/releases/tag/v" + netVersion);
				logger.info(StringUtils.rightPad("", LOG_NUM, "="));
			}
		}
	}

	// Check for newer version
	private static void checkVersion() {

		Thread thread = new Thread() {
			public void run() {

				// Get latest version
				netVersion = ETCommons.getLatestVersion(ETCommons.Project.DIFF_FINDER);

				// If local version is older than newest version
				if (!StringUtils.isEmpty(netVersion) && netVersion.compareTo(appVersion) > 0) {
					isNewVersionExists = true;
				}
			}
		};
		thread.start();
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
			String[] arr = listResult.get(idx).split(SEP);

			// Create new row
			int rIdx = idx + 1;
			sheet.createRow(rIdx);

			// Write result to new row
			for (int aIdx = 0; aIdx < arr.length; aIdx++) {
				sheet.getRow(rIdx).createCell(aIdx).setCellValue(arr[aIdx]);
			}
		}

		// Set filter
		sheet.setAutoFilter(new CellRangeAddress(0, sheet.getLastRowNum(), 0, listHeader.size() - 1));

		// Set Column fit contents
		for (int idx = 0; idx < listHeader.size(); idx++) {
			sheet.autoSizeColumn(idx);
		}
	}

	/**
	 * Search Process
	 * 
	 * @param file File
	 */
	private static void doCompare() {

		try {
			// File 1
			File file_1 = new File(config_path_1);

			// Get file extension
			String ext_1 = FilenameUtils.getExtension(file_1.getName());

			if (file_1.getName().startsWith("~")) {
				logger.info("Ignored: " + file_1.getName());
				return;
			}

			Workbook workbook_1;
			// If Excel 2007 file format
			if (ext_1.equals("xlsx")) {
				workbook_1 = new XSSFWorkbook(new FileInputStream(file_1));
			}
			// If Excel 97-2003 file format
			else if (ext_1.equals("xls")) {
				workbook_1 = new HSSFWorkbook(new FileInputStream(file_1));
			}
			// If none above
			else {
				logger.warn("File not Supported: " + file_1.getName());
				return;
			}

			// File 2
			File file_2 = new File(config_path_2);

			// Get file extension
			String ext_2 = FilenameUtils.getExtension(file_2.getName());

			if (file_2.getName().startsWith("~")) {
				logger.info("Ignored: " + file_2.getName());
				return;
			}

			Workbook workbook_2;
			// If Excel 2007 file format
			if (ext_2.equals("xlsx")) {
				workbook_2 = new XSSFWorkbook(new FileInputStream(file_2));
			}
			// If Excel 97-2003 file format
			else if (ext_2.equals("xls")) {
				workbook_2 = new HSSFWorkbook(new FileInputStream(file_2));
			}
			// If none above
			else {
				logger.warn("File not Supported: " + file_2.getName());
				return;
			}
			
			// Check match number of sheets
			if (workbook_1.getNumberOfSheets() != workbook_2.getNumberOfSheets()) {
				logger.warn("Number of sheets between two file is not same");
				logger.warn(String.format("File 1: %d sheets", workbook_1.getNumberOfSheets()));
				logger.warn(String.format("File 2: %d sheets", workbook_2.getNumberOfSheets()));
				return;
			}

			// Loop though all sheets in workbook
			for (int idx = 0; idx < workbook_1.getNumberOfSheets(); idx++) {

				// Get sheet by index
				Sheet sheet_1 = workbook_1.getSheetAt(idx);

				// Loop though all rows
				for (int rIdx = 0; rIdx < sheet_1.getLastRowNum(); rIdx++) {
					if (sheet_1.getRow(rIdx + 1) != null) {
						// Loop though all columns
						for (int cIdx = 0; cIdx < sheet_1.getRow(rIdx + 1).getLastCellNum(); cIdx++) {
							// Get cell value
							String cellVal = formatter.formatCellValue(sheet_1.getRow(rIdx + 1).getCell(cIdx));
							if (!StringUtils.isBlank(cellVal)) {
							}
						}
					}
				}
			}
		} catch (IOException e) {
			logger.error("IO Exception doCompare()");
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

		// Get file 1 path
		int col_Path_1 = Integer.parseInt(getProp("config.search.path.col")) - 1;
		int row_Path_1 = Integer.parseInt(getProp("config.search.path.1.row")) - 1;
		config_path_1 = formatter.formatCellValue(sheet.getRow(row_Path_1).getCell(col_Path_1));

		// Check if search path is file
		if (!Files.isRegularFile(Path.of(config_path_1))) {
			logger.error("File 1 is not exist");
			workbook.close();
			return false;
		}

		// Get file 2 path
		int col_Path_2 = Integer.parseInt(getProp("config.search.path.col")) - 1;
		int row_Path_2 = Integer.parseInt(getProp("config.search.path.2.row")) - 1;
		config_path_2 = formatter.formatCellValue(sheet.getRow(row_Path_2).getCell(col_Path_2));

		// Check if search path is file
		if (!Files.isRegularFile(Path.of(config_path_2))) {
			logger.error("File 2 is not exist");
			workbook.close();
			return false;
		}

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
		} else if (StringUtils.isBlank(getProp("config.search.path.1.row"))) {
			logger.error("Config: Missing config.search.path.1.row property");
			return false;
		} else if (StringUtils.isBlank(getProp("config.search.path.2.row"))) {
			logger.error("Config: Missing config.search.path.2.row property");
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
