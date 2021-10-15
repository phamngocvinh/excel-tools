package main;

import java.io.BufferedReader;
import java.io.InputStreamReader;
import java.net.URL;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ETCommons {

	/**
	 * Project Enums
	 *
	 */
	public static enum Project {
		TEXT_FINDER("text-finder"), DIFF_FINDER("diff-finder");

		private String name;

		private Project(String name) {
			this.name = name;
		}

		public String getName() {
			return this.name;
		}
	}

	/**
	 * Get latest version
	 * 
	 * @param projectName
	 * @return
	 */
	public static String getLatestVersion(Project project) {
		try {
			// Get latest version
			URL url = new URL("https://api.github.com/repos/phamngocvinh/excel-tools/releases");

			BufferedReader in = new BufferedReader(new InputStreamReader(url.openStream()));

			String line = in.readLine();
			in.close();

			// Get release version
			Pattern p = Pattern.compile(".+?name.+?" + project.getName() + "-v(.+).zip.+");
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
