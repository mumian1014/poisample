package com.mumian.excel;

import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelManager {

	private Workbook workBook;
	private Sheet currentSheet;

	/**
	 * コンストラクタです。
	 * 
	 * @param filePath テンプレートファイルパス
	 */
	public ExcelManager(String filePath) {
		try {
			XSSFWorkbook loadbook = new XSSFWorkbook(getClass().getResourceAsStream(filePath));
			workBook = new SXSSFWorkbook(loadbook);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	/**
	 * シート番号を設定します。
	 * 
	 * @param sheetNo シート番号
	 */
	public void setSheetNo(int sheetNo) {
		currentSheet = workBook.getSheetAt(sheetNo);
	}

	/**
	 * 指定された名前のシートに移動します。
	 * 
	 * @param sheetName シート名
	 */
	public void moveSheet(String sheetName) {
		int index = workBook.getSheetIndex(sheetName);
		currentSheet = workBook.getSheetAt(index);
	}

	/**
	 * Excelのワークブックに書き込みを行います。
	 * 
	 * @param os 書き込みストリーム
	 */
	public void writeFile(OutputStream os) {
		try {
			workBook.write(os);
		} catch (IOException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	/**
	 * シートを取得します。
	 * 
	 * @param sheetName シート名
	 * @return 取得されたシート
	 */
	public Sheet setCurrentSheet(String sheetName) {
		currentSheet = workBook.getSheet(sheetName);
		return currentSheet;
	}

	/**
	 * シートの複製を行います。
	 * 
	 * @param sheetName         作成するシート名
	 * @param templateSheetName テンプレートとなるシート名
	 */
	public void createSheet(String sheetName, String templateSheetName) {
		int index = workBook.getSheetIndex(templateSheetName);
		workBook.cloneSheet(index);
		int cnt = workBook.getNumberOfSheets();
		cnt--;
		workBook.setSheetName(cnt, sheetName);
	}

	/**
	 * 指定されたシートを削除します。
	 * 
	 * @param sheetName 削除したいシートの名称
	 */
	public void deleteSheet(String sheetName) {
		int index = workBook.getSheetIndex(sheetName);
		workBook.removeSheetAt(index);
	}

	/**
	 * ワークシートのセルに値を設定します。
	 * 
	 * @param index シートのインデックス
	 * @param data  設定内容
	 */
	public void setCellValue(String index, String data) {
		SXSSFCell cell = getCell(index);
		if (cell != null) {
			cell.setCellValue(data);
		}
	}

	/**
	 * ワークシートのセルに値を設定します。
	 * 
	 * @param index シートのインデックス
	 * @param data  設定内容
	 */
	public void setCellValue(String index, int data) {
		SXSSFCell cell = getCell(index);
		if (cell != null) {
			cell.setCellValue((double) data);
		}
	}

	/**
	 * ワークシートのセルに値を設定します。
	 * 
	 * @param index シートのインデックス
	 * @param data  設定内容
	 */
	public void setCellValue(String index, double data) {
		SXSSFCell cell = getCell(index);
		if (cell != null) {
			cell.setCellValue(data);
		}
	}

	/**
	 * ワークシートからセルを取得します。
	 * 
	 * @param index シートのインデックス
	 * @return 取得結果
	 */
	public String getCellValue(String index) {
		SXSSFCell cell = getCell(index);
		return cell.getStringCellValue();
	}

	/**
	 * ワークシートからセルを取得します。
	 * 
	 * @param index シートのインデックス
	 * @return セル
	 */
	protected SXSSFCell getCell(String index) {

		ExcelPoint point = toPoint(index);

		// TODO 確認
		SXSSFRow row = (SXSSFRow) currentSheet.getRow(point.getRow() - 1);
		if (row == null) {
			row = (SXSSFRow) currentSheet.createRow(point.getRow() - 1);
		}
		SXSSFCell cell = row.getCell(point.getColumn());
		if (cell == null) {
			cell = row.createCell(point.getColumn());
		}

		return cell;

	}

	/**
	 * Excelのセルの位置を特定します。
	 * 
	 * @param index シートのインデックス
	 * @return セルの位置情報
	 */
	private ExcelPoint toPoint(String index) {
		String row = "";
		String column = "";

		for (int i = 0; i < index.length(); i++) {
			char c = index.charAt(i);

			if (Character.isDigit(c)) {
				row = index.substring(i);
				break;
			} else {
				column += c;
			}
		}

		return new ExcelPoint(row, column);
	}

	/**
	 * シート名として成立するかをチェックします。 <br>
	 * 成立しない場合は、成立する形に変換します。
	 * 
	 * <pre>
	 * 以下に示すものは全角・半角ともにシート名に使用することができない文字です。
	 * 1.コロン        : <br>
	 * 2.円記号       \ <br>
	 * 3.疑問符       ? <br>
	 * 4.角かっこ     [ ]<br>
	 * 5.スラッシュ     / <br>
	 * 6.アスタリスク   *  <br>
	 * </pre>
	 * 
	 * @param name シート名
	 * @return 禁止されている文字が含まれる場合はその文字が除かれたシート名、含まれない場合はそのままのシート名。
	 */
	public String checkSheetName(String name) {

		// コロン(半角)
		name = name.replaceAll(":", "");
		// コロン(全角)
		name = name.replaceAll("：", "");

		// 円記号(半角)
		name = name.replaceAll("\\\\", "");
		// 円記号(全角)
		name = name.replaceAll("￥", "");

		// 疑問符(半角)
		name = name.replaceAll("\\?", "");
		// 疑問符(全角)
		name = name.replaceAll("？", "");

		// 角かっこ (半角)
		name = name.replaceAll("\\[", "");
		name = name.replaceAll("]", "");

		// 角かっこ(全角)
		name = name.replaceAll("\\［", "");
		;
		name = name.replaceAll("］", "");

		// スラッシュ(半角)
		name = name.replaceAll("/", "");
		// スラッシュ(全角)
		name = name.replaceAll("／", "");

		// アスタリスク(半角)
		name = name.replaceAll("\\*", "");
		// アスタリスク(全角)
		name = name.replaceAll("＊", "");

		return name;
	}

}
