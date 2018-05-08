#include <stdlib.h>

char* checkNumeralString(char* str, unsigned char radix);

//
//n進整数として正しいか検査する
//正しかった場合は、最終文字位置アドレスを返す
//誤りがあった場合は、その文字アドレスを返す
//引数に指定可能な数値は16迄で、
//16を超えた数値を設定した場合もしくは文字列がNULLの場合は、
//検査対象文字列の先頭アドレスを返す
//判定可能文字列長は255文字まで
//
char* checkNumeralString(char* str, unsigned char radix) {

	//変数宣言
	char minOfRange1;
	char maxOfRange1;
	char minOfRange2;
	char maxOfRange2;

	char radixIsBiggerThan10;

	unsigned char idxOfStr;

	//引数チェック
	if (str == NULL || 16 < radix) {
		return str;
	}

	//可能文字範囲設定
	minOfRange1 = '0';

	if (radix <= 10) { //基数は10以下
		radixIsBiggerThan10 = 0;
		maxOfRange1 = '9' - (10 - radix);
		

	}else{ //基数は10より大きい
		radixIsBiggerThan10 = 1;
		maxOfRange1 = '9';
		minOfRange2 = 'A';
		maxOfRange2 = 'F' - (16 - radix);

	}

	for (idxOfStr = 0; (str[idxOfStr] != '\0') && (idxOfStr < 255u); idxOfStr++) {
		
		if ((str[idxOfStr] < minOfRange1) || (maxOfRange1 < str[idxOfStr])) { //文字範囲1の範囲外

			if (radixIsBiggerThan10) { //基数は10より大きい
				if ((str[idxOfStr] < minOfRange2) || (maxOfRange2 < str[idxOfStr])) { //文字範囲2の範囲外
					break;
				}
			
			} else { //基数は10以下
				break;
			}
		}
	}

	return &str[idxOfStr];
}