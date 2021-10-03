package main;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Collection;
import java.util.LinkedList;
import java.util.List;
import java.util.Map.Entry;
import java.util.Properties;

import org.apache.commons.io.FileUtils;
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {

	/**
	 * Application Version
	 */
	private static final String appVersion = "0.1.2";

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
	 * Seperator
	 */
	private static final String SEP = "&%&";

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

		logger.info("=== Text Finder ver." + appVersion + " ===");
		logger.info("=== START ===");

		try {
			// Get config file
			try {
				FileInputStream fis = new FileInputStream(new File(dir + "\\" + name_config));
				prop.load(fis);
			} catch (FileNotFoundException e) {
				logger.error("FileNotFoundException: config.properties");
				return;
			} catch (IOException e) {
				logger.error("IOException: config.properties");
				return;
			}

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

			// If search path is folder
			if (config_IsFolder) {
				Collection<File> fileList = FileUtils.listFiles(new File(config_search_path), null, false);
				for (File file : fileList) {
					doSearch(file);
				}
			} else {
				// If search path is file
				doSearch(new File(config_search_path));
			}

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

		} catch (Exception ex) {
			logger.error("Internal Exception: " + ex.getLocalizedMessage());
		} finally {
			logger.info("=== END ===");
		}
	}

	/**
	 * Write Result to Workbook
	 */
	private static void writeResult() {

		// Get Result Sheet
		XSSFSheet sheet = wb_Result.getSheet("Result");

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
	}

	/**
	 * Search Process
	 * 
	 * @param file File
	 */
	private static void doSearch(File file) {
		try {
			Workbook workbook;

			// Get file extension
			String ext = FilenameUtils.getExtension(file.getName());

			// If Excel 2007 file format
			if (ext.equals("xlsx")) {
				workbook = new XSSFWorkbook(new FileInputStream(file));
			}
			// If Excel 97-2003 file format
			else if (ext.equals("xls")) {
				workbook = new HSSFWorkbook(new FileInputStream(file));
			}
			// If none above
			else {
				logger.warn("File not Supported: " + file.getName() + " -> Ignored");
				return;
			}

			logger.info("Searching " + file.getName());

			// Loop though all sheets in workbook
			for (int idx = 0; idx < workbook.getNumberOfSheets(); idx++) {

				// Get sheet by index
				Sheet sheet = workbook.getSheetAt(idx);

				// Loop though all rows
				for (int rIdx = 0; rIdx < sheet.getLastRowNum(); rIdx++) {
					if (sheet.getRow(rIdx + 1) != null) {
						// Loop though all columns
						for (int cIdx = 0; cIdx < sheet.getRow(rIdx + 1).getLastCellNum(); cIdx++) {
							// Get cell value
							String cellVal = formatter.formatCellValue(sheet.getRow(rIdx + 1).getCell(cIdx));
							if (!StringUtils.isBlank(cellVal)) {
								// Loop through search condition
								for (String srchCond : config_SrchCond) {
									// if cell value match search condition
									if (srchCond.equals(cellVal)) {
										// Result = Search Condition + Column + Row + Sheet Name + File Name + File Path
										logger.info("Found " + srchCond + " at "
												+ CellReference.convertNumToColString(cIdx) + (rIdx + 2));
										String result = StringUtils.joinWith(SEP, srchCond,
												CellReference.convertNumToColString(cIdx), (rIdx + 2),
												sheet.getSheetName(), file.getName(), file.getAbsolutePath());
										listResult.add(result);
									}
								}
							}
						}
					}
				}

				// Loop through all comments
				for (Entry<CellAddress, ? extends Comment> entry : sheet.getCellComments().entrySet()) {

					// Cell location
					String location = entry.getKey().toString();
					// Comment string
					String comment = entry.getValue().getString().toString();

					// Loop through search condition
					for (String srchCond : config_SrchCond) {
						// if comment match search condition
						if (srchCond.equals(comment)) {
							logger.info("Found " + srchCond + " at "
									+ location);
							// Result = Search Condition + Column + Row + Sheet Name + File Name + File Path
							String result = StringUtils.joinWith(SEP, srchCond,
									location,
									sheet.getSheetName(), file.getName(), file.getAbsolutePath());
							listResult.add(result);
						}
					}
				}

			}
		} catch (IOException e) {
			logger.error("IOException: " + file.getName());
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
