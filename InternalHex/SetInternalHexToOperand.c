#include "SetInternalHexToOperand.h"
#include "CheckNumeralString.h"
#include "IsBigEndian.h"

#include <string.h>

int setInternalHexToOperand(char* hexStr, void* toWrite, int operandType);
static unsigned char get4BitsFromHexChar(char ch);
int getSizeOfOperand(int type);

//
// 0:正常終了
// 1以上:hex文字列中に誤りがあった文字位置
// -1:hex文字列の字数不足
// -2:引き数がNULL
// -3:operandTypeが不明
//
//operandType
// 0:double
// 1:float
// 2:long
//
int setInternalHexToOperand(char* hexStr, void* toWrite, int operandType) {

	int expectedValLength;
	char* retOfCheckNumeralString;
	int cnt;
	unsigned char startRelativeAddress;
	char addressDir;
	unsigned char mxWrite;
	unsigned char indexOfToWrite;

	//引き数チェック
	if (hexStr == NULL || toWrite == NULL) {
		return -2;
	}

	switch (operandType) { //データ型判定
	case TYPE_DOUBLE:
		mxWrite = sizeof(double);
		break;

	case TYPE_FLOAT:
		mxWrite = sizeof(float);
		break;

	case TYPE_LONG:
		mxWrite = sizeof(long);
		break;

	default: //不明型
		return -3;
		break;
	}

	//文字列長チェック
	expectedValLength = mxWrite * 2;
	if (strlen(hexStr) != expectedValLength) {
		return -1;
	}

	//文字列チェック
	retOfCheckNumeralString = checkNumeralString(hexStr, 16);

	if (*retOfCheckNumeralString != '\0') { //hex文字列中に誤りがあった時
		return (int)(retOfCheckNumeralString - hexStr) + 1; //誤りがあった文字位置を返す
	}

	//書込方向
	if (isBigEndian()) {
		startRelativeAddress = 0;
		addressDir = 1;

	}else {
		startRelativeAddress = mxWrite - 1;
		addressDir = -1;
	}

	//書込ループ
	indexOfToWrite = startRelativeAddress;
	for (cnt = 0; cnt < (expectedValLength); cnt += 2) {
		
		((unsigned char*)toWrite)[indexOfToWrite] = (get4BitsFromHexChar(hexStr[cnt]) << 4) + (get4BitsFromHexChar(hexStr[cnt + 1]));
		indexOfToWrite += addressDir;
	}

	return 0;
}

static unsigned char get4BitsFromHexChar(char ch) {
	
	unsigned char ret;

	if ('0' <= ch && ch <= '9') {
		ret = 0 + (ch - '0');
	}
	else {
		ret = 10 + (ch - 'A');
	}

	return ret;
}

//
//変数型のサイズを返す
//引数指定が不明数値の場合は、-1を返す
//
//type
// 0:double
// 1:float
// 2:long
//
int getSizeOfOperand(int type) {

	int ret;

	switch (type) {

	case TYPE_DOUBLE:
		ret = sizeof(double);
		break;

	case TYPE_FLOAT:
		ret = sizeof(float);
		break;

	case TYPE_LONG:
		ret = sizeof(long);
		break;

	default:
		ret = -1;
		break;
	}

	return ret;

}