import java.io.File;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

public class Main {

	// Logger
	private static final Logger logger = LogManager.getLogger();

	public static void main(String[] args) {

		// Current path
		final String dir = System.getProperty("user.dir");
		System.out.println("current dir = " + dir);

		// Config file path
		final String configPath = dir + "\\config.xlsx";

		// Check if config file exists
		if (new File(configPath).exists()) {
			logger.info("Found config");
		} else {
			logger.error("Config not exists");
		}

		readConfig();
	}

	private static void readConfig() {

	}

}
