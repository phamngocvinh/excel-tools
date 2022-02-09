package main;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.Collection;
import java.util.LinkedList;
import java.util.List;
import java.util.Map.Entry;
import java.util.Properties;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.commons.io.FileUtils;
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.io.filefilter.TrueFileFilter;
import org.apache.commons.lang3.StringUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.ShapeContainer;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFShapeGroup;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
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
	 * Config: Search path
	 */
	private static String config_search_path = "";

	/**
	 * Config: Is Search in Folder
	 */
	private static boolean config_IsFolder = false;

	/**
	 * Config: Is Search Recursively
	 */
	private static boolean config_IsSearchRecursively = false;

	/**
	 * Config: Search Conditions
	 */
	private static List<String> config_SrchCond = new LinkedList<>();

	/**
	 * List of Shapes
	 */
	private static List<XSSFSimpleShape> listShapes = new LinkedList<>();

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

			// Get configuration file
			FileInputStream fis = new FileInputStream(new File(dir + "\\" + name_config));
			prop.load(fis);

			// Get application version from properties
			appVersion = getProp("app.version");

			// Check for new version
			checkVersion();

			logger.info(StringUtils.rightPad("=== Text Finder v" + appVersion + " ", LOG_NUM, "="));
			logger.info(StringUtils.rightPad("=== START ", LOG_NUM, "="));

			// Check if configuration file valid
			if (!isValidConfig()) {
				return;
			}

			// Read configuration file
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
				Collection<File> fileList;
				if (config_IsSearchRecursively) {
					fileList = FileUtils.listFiles(new File(config_search_path), TrueFileFilter.INSTANCE,
							TrueFileFilter.INSTANCE);
				} else {
					fileList = FileUtils.listFiles(new File(config_search_path), null, false);
				}

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
				netVersion = getLatestVersion();

				// If local version is older than newest version
				if (netVersion != null && netVersion.compareTo(appVersion) > 0) {
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

		CreationHelper createHelper = wb_Result.getCreationHelper();

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
				// Filename hyperlink
				if (aIdx == 4) {
					sheet.getRow(rIdx).createCell(aIdx).setCellValue(arr[aIdx]);

					try {
						// Change path string format
						XSSFHyperlink link = (XSSFHyperlink) createHelper.createHyperlink(HyperlinkType.FILE);
						link.setAddress(arr[4].replace("\\", "/"));
						sheet.getRow(rIdx).getCell(aIdx).setHyperlink((XSSFHyperlink) link);

						// Set link font
						Font blueFont = wb_Result.createFont();
						blueFont.setColor(IndexedColors.BLUE.getIndex());
						blueFont.setUnderline(Font.U_SINGLE);
						CellUtil.setFont(sheet.getRow(rIdx).getCell(aIdx), blueFont);
					} catch (Exception ex) {

					}
				} else {
					sheet.getRow(rIdx).createCell(aIdx).setCellValue(arr[aIdx]);
				}
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
	@SuppressWarnings("unchecked")
	private static void doSearch(File file) {

		try {
			Workbook workbook;

			// Get file extension
			String ext = FilenameUtils.getExtension(file.getName());

			if (file.getName().startsWith("~")) {
				logger.info("Ignored: " + file.getName());
				return;
			}

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

				// Get all shapes
				getAllShapes((ShapeContainer<XSSFShape>) sheet.getDrawingPatriarch());

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
									if (cellVal.contains(srchCond)) {
										// Result = Search Condition + Column + Row + Sheet Name + File Name + File Path
										logger.info("Found " + srchCond + " at "
												+ CellReference.convertNumToColString(cIdx) + (rIdx + 2));
										String result = StringUtils.joinWith(SEP, srchCond,
												CellReference.convertNumToColString(cIdx) + (rIdx + 2),
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
						if (comment.contains(srchCond)) {
							logger.info("Found " + srchCond + " at " + location);
							// Result = Search Condition + Column + Row + Sheet Name + File Name + File Path
							String result = StringUtils.joinWith(SEP, srchCond, location, sheet.getSheetName(),
									file.getName(), file.getAbsolutePath());
							listResult.add(result);
						}
					}
				}

				// Loop through all shapes
				for (XSSFSimpleShape shape : listShapes) {
					// Loop through search condition
					for (String srchCond : config_SrchCond) {
						// if shape text match search condition
						if (shape.getText().contains(srchCond)) {
							logger.info("Found " + srchCond + " in shape ");
							// Result = Search Condition + Column + Row + Sheet Name + File Name + File Path
							String result = StringUtils.joinWith(SEP, srchCond, null, sheet.getSheetName(),
									file.getName(), file.getAbsolutePath());
							listResult.add(result);
						}
					}
				}
			}
		} catch (IOException e) {
			logger.warn("Cannot read " + file.getName() + " -> Ignored");
		}
	}

	/**
	 * Read Configuration File to get settings
	 * 
	 * @throws IOException
	 * @throws InvalidFormatException
	 * 
	 * @throws Exception
	 */
	private static boolean readConfig() throws InvalidFormatException, IOException {

		// Check if configuration excel exists
		File fileWorkBook = new File(dir + prop.getProperty("config.file"));
		if (fileWorkBook.exists()) {
			logger.info("Read config.xlsx");
		} else {
			logger.error("Config file not exists");
			return false;
		}

		// Open configuration file
		XSSFWorkbook workbook = new XSSFWorkbook(fileWorkBook);

		// Open configuration sheet
		XSSFSheet sheet = workbook.getSheet(prop.getProperty("textfinder.config.sheet"));

		// Check if configuration sheet exists
		if (sheet == null) {
			logger.error("Config sheet not found");
			workbook.close();
			return false;
		}

		// Get recursively option
		int col_Recursive = Integer.parseInt(getProp("textfinder.config.recursive.col")) - 1;
		int row_Recursive = Integer.parseInt(getProp("textfinder.config.recursive.row")) - 1;
		String recursiveOption = formatter.formatCellValue(sheet.getRow(row_Recursive).getCell(col_Recursive))
				.toLowerCase();
		if ("y".equals(recursiveOption) || "yes".equals(recursiveOption)) {
			config_IsSearchRecursively = true;
		}

		// Get search path
		int col_Path = Integer.parseInt(getProp("textfinder.config.path.col")) - 1;
		int row_Path = Integer.parseInt(getProp("textfinder.config.path.row")) - 1;
		config_search_path = formatter.formatCellValue(sheet.getRow(row_Path).getCell(col_Path));

		// Check if search path is file or folder
		Path path = new File(config_search_path).toPath();
		if (Files.isDirectory(path)) {
			config_IsFolder = true;
			if (path.toFile().listFiles().length == 0) {
				logger.info("Search path is empty");
				workbook.close();
				return false;
			}
			logger.info("Search in " + config_search_path);
		} else if (Files.isRegularFile(path)) {
			config_IsFolder = false;
			logger.info("Search in " + config_search_path);
		} else {
			logger.error("Search path not exist. Please check config.xlsx");
			workbook.close();
			return false;
		}

		// Get Search Condition
		int col_Cond = Integer.parseInt(getProp("textfinder.config.cond.col")) - 1;
		int row_Cond = Integer.parseInt(getProp("textfinder.config.cond.row")) - 1;
		int row_Last = sheet.getLastRowNum();

		for (int idx = row_Cond; idx <= row_Last; idx++) {
			config_SrchCond.add(formatter.formatCellValue(sheet.getRow(idx).getCell(col_Cond)));
		}
		logger.info("Search condition: " + config_SrchCond);

		try {
			workbook.close();
		} catch (IOException ex) {
			return true;
		}

		return true;
	}

	/**
	 * Check if Configuration file is valid
	 * 
	 * @return
	 */
	private static boolean isValidConfig() {

		// Check if correctly configuration
		if (StringUtils.isBlank(getProp("app.version"))) {
			logger.error("Config: Missing app.version property");
			return false;
		} else if (StringUtils.isBlank(getProp("config.file"))) {
			logger.error("Config: Missing config.file property");
			return false;
		} else if (StringUtils.isBlank(getProp("textfinder.config.sheet"))) {
			logger.error("Config: Missing textfinder.config.sheet property");
			return false;
		} else if (StringUtils.isBlank(getProp("textfinder.config.path.col"))) {
			logger.error("Config: Missing textfinder.config.path.col property");
			return false;
		} else if (StringUtils.isBlank(getProp("textfinder.config.path.row"))) {
			logger.error("Config: Missing textfinder.config.path.row property");
			return false;
		}

		// If configuration version different from application version, throw error
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

	/**
	 * SetAllShapes
	 * 
	 * @param container
	 */
	private static void getAllShapes(ShapeContainer<XSSFShape> container) {

		if (container != null) {
			for (XSSFShape shape : container) {
				if (shape instanceof XSSFShapeGroup) {
					XSSFShapeGroup shapeGroup = (XSSFShapeGroup) shape;
					getAllShapes(shapeGroup);

				} else if (shape instanceof XSSFSimpleShape) {
					XSSFSimpleShape simpleShape = (XSSFSimpleShape) shape;
					listShapes.add(simpleShape);
				}
			}
		}
	}

	/**
	 * Get latest version
	 * 
	 * @param projectName
	 * @return
	 */
	private static String getLatestVersion() {
		try {
			// Get latest version
			URL url = new URL("https://api.github.com/repos/phamngocvinh/excel-tools/releases");

			BufferedReader in = new BufferedReader(new InputStreamReader(url.openStream()));

			String line = in.readLine();
			in.close();

			// Get release version
			Pattern p = Pattern.compile(".+?name.+?excel-tools-v(.+).zip.+");
			Matcher m = p.matcher(line);
			boolean isMatch = m.matches();

			if (isMatch) {
				return m.group(1).substring(0, m.group(1).indexOf(".zip"));
			}

			return null;
		} catch (Exception e) {
			return null;
		}
	}
}
