package system.apex.apfw.utility;

import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.Date;

/**
 * TimeCard App
 *
 */
public class TimeCardApp
{
    public static void main( String[] args ) throws Exception {

    	boolean result;
    	Date date;
    	String s;

    	// タイムカードの作成
    	TimeCard tc = new TimeCard();

    	// アペックス勤務表情報の作成
    	ExcelTCardImpl_AS_F001 etCard = new ExcelTCardImpl_AS_F001();

    	// EAC勤務表情報の作成
    	//ExcelTCardImpl_EAC_F001 etCard = new ExcelTCardImpl_EAC_F001();


    	// 勤務表のデータ
    	etCard.setBaseDate(new Date());
    	etCard.setName("田中健一");
    	etCard.setPath("C:\\test");

    	// 作成されたアペックス勤務表情報をタイムカードにセットする
    	tc.setExcelTCard(etCard);


    	// タイムカードで勤務時間を記録する
    	s = "出勤時間";
    	//result = tc.checkIn();
    	LocalDateTime ldt = LocalDateTime.now();
    	DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss");
    	String dateTimeStr = ldt.format(DateTimeFormatter.ofPattern("yyyy/MM/dd")) + " 10:30:00";
    	ldt = LocalDateTime.parse(dateTimeStr, formatter);
    	date = Date.from(ldt.atZone(ZoneId.systemDefault()).toInstant());
    	result = tc.printWorkTime(date, TimeCard.CHECK_IN);
    	printResult(tc, s, result);

    	s = "退勤時間";
    	result = tc.checkOut();
    	printResult(tc, s, result);

    	// 休憩時間
    	s = "休憩時間";
    	result = tc.printRestTime(new Date(), "1:30");
    	printResult(tc, s, result);

    	s = "備考";
    	result = tc.printRemarks(new Date(), "備考内容１");
    	printResult(tc, s, result);
    	//result = tc.printRemarks(new Date(), "備考内容２");
    	//printResult(tc, s, result);

    }

    static void printResult(TimeCard tc, String action, boolean result) {
    	if(result) {
    		System.out.println("☆勤務表に" + action + "の記入は成功しました。");
    	} else {
    		System.out.println("★勤務表に" + action + "の記入は失敗しました。");
    		System.out.println("エラー情報は下記です。");
    		tc.getErrorMessages().forEach(System.out::println);
    	}
    }
}
