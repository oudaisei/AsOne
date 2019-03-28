package system.apex.apfw.utility;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.DateUtil;

/**
 * クラス名：TimeCard
 * 概　　要：勤務表に勤務に関する情報（出勤時間や退勤時間など）を記録する
 *
 * @author  ods
 * @version 1.00
 *
 */
public class TimeCard {

	/** 記録種類：勤務開始時間 */
	public static final String CHECK_IN  = "CHECK_IN";

	/** 記録種類：勤務終了時間 */
	public static final String CHECK_OUT = "CHECK_OUT";


	/** エラーメッセージ情報 */
	private List<String> errorMessages = new ArrayList<>();

	/** excel勤務表 */
	private ExcelTCard etCard;


	public List<String> getErrorMessages() {
		return errorMessages;
	}

	public ExcelTCard getExcelTCard() {
		return etCard;
	}

	public void setExcelTCard(ExcelTCard etCard) {
		this.etCard = etCard;
	}


	/**
	 * 処理名：checkIn
	 * 概　要：勤務開始時間の記入
	 *
	 *
	 * @return true:記録成功; false:記録失敗
	 *
	 */
	public boolean checkIn() {
		boolean result;

		// エラーメッセージ情報をクリアする
		errorMessages.clear();

		// 勤務開始時間の記入
		result = printWorkTime(new Date(), CHECK_IN);

		return result;
	}

	/**
	 * 処理名：checkOut
	 * 概　要：勤務終了時間の記入
	 *
	 *
	 * @return true:記録成功; false:記録失敗
	 *
	 */
	public boolean checkOut() {
		boolean result;

		// エラーメッセージ情報をクリアする
		errorMessages.clear();

		// 勤務終了時間の記入
		result = printWorkTime(new Date(), CHECK_OUT);

		return result;
	}

	/**
	 * 処理名：printRemarks
	 * 概　要：備考の記入
	 *
	 * @param date    備考の勤務日
	 * @param remarks 備考内容
	 *
	 * @return true:記録成功; false:記録失敗
	 *
	 */
	public boolean printRemarks(Date date, String remarks) {

		boolean result;

		// エラーメッセージ情報をクリアする
		errorMessages.clear();

		// 勤務表のフルパスのファイル名（入力ファイル）を取得する
		String inputFilename = etCard.getAbsoluteFilename();
		if(Files.notExists(Paths.get(inputFilename))) {
			String s = String.format("勤務表ファイルが存在しないです。\r\n%s", inputFilename);
			errorMessages.add(s);
			return false;
		}

		// 勤務表のシート名
		String sheetName = etCard.getSheetName();
		if(AsStringUtil.isNullOrEmpty(sheetName)) {
			String s = String.format("勤務表のシート名が空白です。\r\n%s", inputFilename);
			errorMessages.add(s);
			return false;
		}

		// 備考欄の列
		String colNum = etCard.getRemarksColNumStr();
		if(AsStringUtil.isNullOrEmpty(colNum)) {
			String s = "勤務表から備考列の情報を取得できません。\r\n" + inputFilename;
			errorMessages.add(s);
			return false;
		}

		// 勤務日の行番号
		int rowIndex;
		rowIndex = etCard.getRowNumOfWorkDay(date);
		if(rowIndex <= 0) {
			SimpleDateFormat sdf = new SimpleDateFormat("yyyy/MM/dd");
			String pDate = sdf.format(date);
			String s = "勤務日（" + pDate +"）は勤務表に見つかりませんでした。\r\n" + inputFilename;
			errorMessages.add(s);
			return false;
		}

		// 勤務表に書き込む
		result = printValue(inputFilename, sheetName, colNum, rowIndex, remarks, true, ";");

		return result;
	}

	/**
	 * 処理名：printRestTime
	 * 概　要：休憩時間の記入
	 *
	 * @param date     休憩時間の勤務日
	 * @param restTime 休憩時間（string of format "HH:MM" or "HH:MM:SS"）
	 *
	 * @return true:記録成功; false:記録失敗
	 *
	 */
	public boolean printRestTime(Date date, String restTime) {
		boolean result;

		// エラーメッセージ情報をクリアする
		errorMessages.clear();

		// Converts a string of format "HH:MM" or "HH:MM:SS" to its (Excel) numeric equivalent
		result = printRestTime(date, DateUtil.convertTime(restTime));

		return result;
	}

	/**
	 * 処理名：printRestTime
	 * 概　要：休憩時間の記入
	 *
	 * @param date     休憩時間の勤務日
	 * @param restTime 休憩時間（double）
	 *
	 * @return true:記録成功; false:記録失敗
	 *
	 */
	public boolean printRestTime(Date date, double restTime) {

		boolean result;

		// エラーメッセージ情報をクリアする
		errorMessages.clear();

		// 勤務表のフルパスのファイル名（入力ファイル）を取得する
		String inputFilename = etCard.getAbsoluteFilename();
		if(Files.notExists(Paths.get(inputFilename))) {
			String s = String.format("勤務表ファイルが存在しないです。\r\n%s", inputFilename);
			errorMessages.add(s);
			return false;
		}

		// 勤務表のシート名
		String sheetName = etCard.getSheetName();
		if(AsStringUtil.isNullOrEmpty(sheetName)) {
			String s = String.format("勤務表のシート名が空白です。\r\n%s", inputFilename);
			errorMessages.add(s);
			return false;
		}

		// 休憩時間の列
		String colNum = etCard.getRestTimeColNumStr();
		if(AsStringUtil.isNullOrEmpty(colNum)) {
			String s = "勤務表から休憩列の情報を取得できません。\r\n" + inputFilename;
			errorMessages.add(s);
			return false;
		}

		// 勤務日の行番号
		int rowIndex;
		rowIndex = etCard.getRowNumOfWorkDay(date);
		if(rowIndex <= 0) {
			SimpleDateFormat sdf = new SimpleDateFormat("yyyy/MM/dd");
			String pDate = sdf.format(date);
			String s = "勤務日（" + pDate +"）は勤務表に見つかりませんでした。\r\n" + inputFilename;
			errorMessages.add(s);
			return false;
		}

		// 勤務表に書き込む
		result = printValue(inputFilename, sheetName, colNum, rowIndex, restTime, false, null);

		return result;
	}


	/**
	 * 処理名：printWorkTime
	 * 概　要：勤務開始・終了（出勤・退勤）時間記入
	 *
	 * @param date  勤務開始・終了の日時
	 * @param type  記録種類：勤務開始時間（CHECK_IN）, 勤務終了時間（CHECK_OUT）
	 *
	 * @return true:記録成功; false:記録失敗
	 *
	 */
	public boolean printWorkTime(Date date, String type) {

		boolean result;

		String colNum;
		// 出勤・退勤の判定
		if(CHECK_IN.equalsIgnoreCase(type)) {
			// 出勤
			colNum = etCard.getCheckInColNumStr();
		} else if(CHECK_OUT.equalsIgnoreCase(type)) {
			// 退勤
			colNum = etCard.getCheckOutColNumStr();
		} else {
			String s = String.format("記録種類は出勤（%s）または退勤（%s）を指定してください。", CHECK_IN, CHECK_OUT);
			errorMessages.add(s);
			return false;
		}

		// 勤務表のファイル名（入力ファイル）を取得する
		String inputFilename = etCard.getAbsoluteFilename();
		if(Files.notExists(Paths.get(inputFilename))) {
			String s = String.format("勤務表ファイルが存在しないです。\r\n%s", inputFilename);
			errorMessages.add(s);
			return false;
		}

		// 勤務表のシート名
		String sheetName = etCard.getSheetName();
		if(AsStringUtil.isNullOrEmpty(sheetName)) {
			String s = String.format("勤務表のシート名が空白です。\r\n%s", inputFilename);
			errorMessages.add(s);
			return false;
		}

		// 勤務日の行番号を取得する
		int rowIndex = etCard.getRowNumOfWorkDay(date);
		if(rowIndex <= 0) {
			SimpleDateFormat sdf = new SimpleDateFormat("yyyy/MM/dd");
			String s = "勤務日（" + sdf.format(date) +"）は勤務表に見つかりませんでした。\r\n" + inputFilename;
			errorMessages.add(s);
			return false;
		}

		// 勤務時間の時刻を記入
		result = printValue(inputFilename, sheetName, colNum, rowIndex, ExcelUtil.timeOnly(date), false, null);

		return result;

	}

	/**
	 * 処理名：printValue
	 * 概　要：出勤表に値を記入する
	 *
	 * @param inputFilename 記入対象の勤務表（Excelファイル）
	 * @param sheetName     勤務表のシート名
	 * @param colNum        列（"A"や"B"など）
	 * @param rowIndex      行番号
	 * @param cellValue     記入値
	 * @param appendMode    追加するかどうか　※追加の場合、cellValueのデータ型はStringのみ対応
	 * @param separator     追加モードの場合、区切り文字の指定（null或いは""の場合、区切り文字）
	 *
	 * @return true:記録成功; false:記録失敗
	 *
	 */
	private boolean printValue(String inputFilename, String sheetName, String colNum, int rowIndex, Object cellValue, boolean appendMode, String separator) {
		boolean result;
		ExcelUtil util;
		String outputFilename;

		// 追加モードの場合、cellValueのデータ型はStringのみ対応
		if(appendMode) {
			if(!(cellValue instanceof String)) {
				errorMessages.add("勤務表の既存内容に追加できる記入値のデータ型はStringのみです。");
				return false;
			}
		}

		result = true;
		outputFilename = null;
		util = null;
		try {
			// 記録されたデータは一時ファイル（出力ファイル）に出力する
			outputFilename = inputFilename + ".output" + String.valueOf(System.currentTimeMillis());

			// ExcelUtil作成
			util = new ExcelUtil.Builder()
								.fromFile(inputFilename) //入力ファイル
								.output(outputFilename)  //出力ファイル
								.build();

			// 追加モードはStringのみ対応
			if(appendMode) {
				String s = (String) util.sheet(sheetName).getResultValue(colNum, rowIndex);
				if(!AsStringUtil.isNullOrEmpty(s)) {
					if(!AsStringUtil.isNullOrEmpty(separator)) {
						s = s + separator + (String) cellValue;
					} else {
						s = s + (String) cellValue;
					}
				} else {
					s = (String) cellValue;
				}

				cellValue = s;
			}

			// cellValueを書き込む
			util.sheet(sheetName).put(colNum, rowIndex, cellValue);

		} catch(Exception e) {
			result = false;
			errorMessages.add(e.getMessage());
			return false;
		}
		finally {
			try {
				// excelに書き込んでからクローズする
				if(null != util) {
					util.close();
				}

				// 出力された勤務表ファイル（一時ファイル）を削除する
				if(!AsStringUtil.isNullOrEmpty(outputFilename)) {
					// 出力された勤務表ファイル（一時ファイル）を勤務表ファイルに上書き
					if(result) {
						Files.copy(Paths.get(outputFilename), Paths.get(inputFilename), StandardCopyOption.REPLACE_EXISTING);
					}
					Files.deleteIfExists(Paths.get(outputFilename));
				}

			} catch(IOException e) {
				result = false;
				errorMessages.add(e.getMessage());
				return false;
			}
		}

		return result;
	}

}
