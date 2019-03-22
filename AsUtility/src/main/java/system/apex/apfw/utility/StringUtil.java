package system.apex.apfw.utility;

/**
 * クラス名： StringUtil
 * 概　　要： ストリングに関するUtility
 *
 * @author  ods
 * @version 1.00
 *
 */

public class StringUtil {

	public static String NEW_LINE_CODE = "\r\n";

	/**
	 * 処理名： isNullOrEmpty
	 * 概　要： ストリングはNULL或いは空白の判定
	 * @param  s ストリング
	 * @return true:Nullや空白の場合
	 */
	public static boolean isNullOrEmpty(String s) {
		boolean result;

		if(null == s) {
			return false;
		}

		result = s.isEmpty();

		return result;
	}

}
