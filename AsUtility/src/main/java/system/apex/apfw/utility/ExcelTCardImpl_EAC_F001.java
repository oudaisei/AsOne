package system.apex.apfw.utility;

import java.io.File;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Sheet;

/**
 * クラス名：ExcelTCardImpl_EAC_F001
 * 概　　要：（株）EACの勤務表（フォーマット：F001）
 *
 * @author  ods
 * @version 1.00
 *
 */
public class ExcelTCardImpl_EAC_F001 implements ExcelTCard {

	/** 勤務表ファイルパス */
	private String path = "";

	/** 勤務表のシート名 */
	private String sheetName = "";

	/** 勤務表基準年月 */
	private Date baseDate = null;

	/** 勤務者の氏名 */
	private String name = "";


	public String getPath() {
		return path;
	}

	public void setPath(String path) {
		this.path = path;
	}

	public Date getBaseDate() {
		return baseDate;
	}

	public void setBaseDate(Date baseDate) {
		this.baseDate = baseDate;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public String getBaseDateYYYYMM() {
		// 勤務表の基準年月
		SimpleDateFormat sdf = new SimpleDateFormat("yyyyMM");
		return sdf.format(this.getBaseDate());
	}


	@Override
	public String getFilename() {
		StringBuilder strBuilder = new StringBuilder();
		String filename;
		String yyyymm;

		// 勤務表の基準年月
		yyyymm = getBaseDateYYYYMM();

		strBuilder.append("EAC作業実績報告書_")
				  .append(yyyymm)
				  .append("_")
				  .append(this.getName())
				  .append(".xls");
		filename = strBuilder.toString();
		return filename;
	}

	/**
	 * 処理名： getAbsoluteFilename
	 * 概　要： 出勤表の絶対パスのファイル名を取得する
	 *
	 *
	 */
	@Override
	public String getAbsoluteFilename() {
		String absoluteFilename;

		absoluteFilename = getPath();
		if((absoluteFilename == null) || (absoluteFilename.isEmpty()) ) {
			return getFilename();
		}
		if(!absoluteFilename.endsWith(File.separator)) {
			absoluteFilename += File.separator;
		}
		absoluteFilename += getFilename();

		return absoluteFilename;
	}

	@Override
	public String getSheetName() {
		try {
			if( (null == this.sheetName) || this.sheetName.isEmpty() ) {
				//this.sheetName = getBaseDateYYYYMM();
				AsExcelUtil util;
				util = new AsExcelUtil.Builder()
			            .fromFile(getAbsoluteFilename()) //入力のエクセル
			            .build();
				Sheet sheet = util.getSheetAt(0);
				this.sheetName = sheet.getSheetName();
			}
		} catch (Exception e) {
			return "";
		}

		return this.sheetName;
	}

	@Override
	public int getRowNumOfWorkDay(Date workday) {
		AsExcelUtil util;
		int cellRowIndex;
		Object cellResultValue;

		SimpleDateFormat sdf = new SimpleDateFormat("d");
		String sDate = sdf.format(workday);

		try {
			util = new AsExcelUtil.Builder()
					            .fromFile(getAbsoluteFilename())//入力のエクセル
					            .build();

			for(cellRowIndex = 11; cellRowIndex <=11+30;  cellRowIndex++) {
				cellResultValue = util.sheet(0).getResultValue("A", cellRowIndex);
				if(null ==  cellResultValue) {
					util.close();
					return 0;
				}
				if(cellResultValue instanceof Double) {
					int cellDay = ((Double) cellResultValue).intValue();
					int day = Integer.parseInt(sDate, 10);
					if(cellDay == day) {
						break;
					}
				}
			}
			util.close();
			if(cellRowIndex > 6+30) {
				// 取得できない
				return 0;
			}
		} catch(Exception e) {
			// 異常の場合
			return -1;
		}

		return cellRowIndex;
	}

	@Override
	public String getCheckInColNumStr() {
		// 勤務開始時間の列
		return "C";
	}

	@Override
	public String getCheckOutColNumStr() {
		// 勤務終了時間の列
		return "F";
	}

	@Override
	public String getRestTimeColNumStr() {
		// 休憩時間の列
		return "I";
	}

	@Override
	public String getRemarksColNumStr() {
		// 備考の列
		return "X";
	}

}
