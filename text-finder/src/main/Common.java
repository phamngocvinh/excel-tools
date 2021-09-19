package main;
/**
 *
 */
public class Common {
	/**
	 * Check if String is null or blank empty
	 * 
	 * @param str
	 * @return
	 */
	public static boolean isNullOrEmpty(String str) {

		if (str == null || str.length() == 0) {
			return true;
		}

		return false;
	}
}
