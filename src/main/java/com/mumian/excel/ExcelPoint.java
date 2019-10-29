package com.mumian.excel;

/**
 * Excelの座標を表すクラスです。
 *
 */
public class ExcelPoint {

	/** 行 */
	private int row;
	/** 列英字 */
	private String column;

	/**
	 * コンストラクタです。
	 * 
	 * @param row 行
	 * @param columns 列
	 */
	public ExcelPoint(int row, String column) {
		this.row = row;
		this.column = column;
	}
	
	/**
	 * コンストラクタです。
	 * 
	 * @param row 行
	 * @param columns 列
	 */
	public ExcelPoint(String row, String column) {
		this.row = Integer.parseInt(row);
		this.column = column;
	}
	
	/**
	 * 行を取得します。
	 * @return 行インデックス
	 */
	public int getRow() {
		return row;
	}
	
	/**
	 * 列を取得します。
	 * @return 列インデックス
	 */
	public int getColumn() {
		return toInt(column);
	}
	
	/**
	 * Excelのカラム英字からRowのindexに変換します。
	 * @param str カラム英字
	 * @return index
	 */
	private int toInt(String str) {

		char c = str.charAt(0);
		int value = Character.getNumericValue(c) - 10;

		if (str.length() == 2) {
			char c2 = str.charAt(1);
			int value2 = Character.getNumericValue(c2) - 10;

			// 1桁目がAの場合
			if (c == 'A') {
				value = 26 + value2;

			// 1桁目がBの場合
			} else if (c == 'B') {
				value = 52 + value2;
			
			// 1桁目がCの場合
			} else if (c == 'C') {
				value = 78 + value2;
				
			// 1桁目がdの場合
			} else if (c == 'D') {
				value = 104 + value2;
			}
		}

		return value;

	}

}
