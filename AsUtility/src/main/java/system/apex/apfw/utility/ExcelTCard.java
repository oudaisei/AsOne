package system.apex.apfw.utility;

import java.util.Date;

/**
 * Interface 名： ExcelTCard
 * 概　　　　要： Excel勤務表の情報を提供する
 *
 * @author ods
 *
 */
public interface ExcelTCard {

	/**
	 * 処理名： getFilename
	 * 概　要： 勤務表のファイル名を取得する
	 *
	 * @return 取得できた場合、勤務表のファイル名; 取得できなかった場合、空白（""）
	 *
	 */
	String getFilename();

	/**
	 * 処理名： getAbsoluteFilename
	 * 概　要： 勤務表のフルパスのファイル名を取得する
	 *
	 * @return 取得できた場合、勤務表のフルパスのファイル名; 取得できなかった場合、空白（""）
	 *
	 */
	String getAbsoluteFilename();

	/**
	 * 処理名： getSheetName
	 * 概　要： 勤務表のシート名を取得する	 *
	 *
	 * @return 取得できた場合、シート名; 取得できなかった場合、空白（""）
	 */
	String getSheetName();

	/**
	 * 処理名： getRowNumOfWorkDay
	 * 概　要： 勤務日の行番号を取得する
	 *
	 * @param workday 勤務日
	 *
	 * @return 取得できた場合、行番号; 取得できなかった場合、0; 異常の場合、-1
	 */
	int getRowNumOfWorkDay(Date workday);

	/**
	 * 処理名： getCheckInColNumStr
	 * 概　要： 勤務開始時間のセル列を取得する（例："A"、"D"など）
	 *
	 *
	 * @return 取得できた場合、列の文字列; 取得できなかった場合、空白（""）
	 */
	String getCheckInColNumStr();

	/**
	 * 処理名： getCheckOutColNumStr
	 * 概　要： 勤務終了時間のセル列を取得する（例："A"、"D"など）
	 *
	 *
	 * @return 取得できた場合、列の文字列; 取得できなかった場合、空白（""）
	 */
	String getCheckOutColNumStr();

	/**
	 * 処理名： getRestTimeColNumStr
	 * 概　要： 休憩時間のセル列を取得する（例："A"、"D"など）
	 *
	 *
	 * @return 取得できた場合、列の文字列; 取得できなかった場合、空白（""）
	 */
	String getRestTimeColNumStr();

	/**
	 * 処理名： getRemarksColNumStr
	 * 概　要： 備考のセル列を取得する（例："A"、"D"など）
	 *
	 *
	 * @return 取得できた場合、列の文字列; 取得できなかった場合、空白（""）
	 */
	String getRemarksColNumStr();


}
