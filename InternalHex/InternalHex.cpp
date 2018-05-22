#include "InternalHex.h"
#include "SetInternalHexToOperand.h"
#include "GetInternalHexFromOperand.h"

#include <Windows.h>

#include <stdlib.h>
// free

#include <string.h>
// strlen


__declspec (dllexport) int WINAPI convDecStrToOperandAndGetInternalHex(char* toWriteStr, int lenOfToWriteStr, char* toConvStr, int lenOfToConvStr, int type);
__declspec (dllexport) int WINAPI operateArithmeticByInternalHex(char* val1, char* val2, char* sum, int lenOfSum, int operandType, int operateType);
__declspec (dllexport) int WINAPI getSizeOfOperandExp(int operandType);

static void copyStr(char* destStr, char* sourceStr, unsigned int bytesOfStr);

//
//10進数値文字列を変数に格納して、
//その変数の内部表現をhexで返す
//
//type
// 0:double
// 1:float
// 2:long
//
//戻り値
// 1以上:書き込み文字列byte長('\0'抜き)
// -1:指定文字列がNULL
// -2:変換タイプが不明
// -3:変換文字列が数値変更不可
// -4:書き込み先バッファ長不足
// -5:メモリ不足
//
__declspec (dllexport) int WINAPI convDecStrToOperandAndGetInternalHex(char* toWriteStr, int lenOfToWriteStr, char* toConvStr, int lenOfToConvStr, int type) {

	//変数宣言
	float   resultOfstrtof;    //文字列→変数変換後用一時格納先(float)
	double  resultOfstrtod;    //文字列→変数変換後用一時格納先(double)
	long    resultOfstrtol;    //文字列→変数変換後用一時格納先(long)
	char*   endptr;            //strtoXコール用
	char*   generatedHexStr;   //生成Hex文字列
	unsigned char    lenOfGeneratedStr; //生成Hex文字列のbyte長
	unsigned char    lenOfWroteStr;     //書き込み先byte長

										//引数チェック
	if (toWriteStr == NULL || toConvStr == NULL) {
		return -1;
	}

	switch (type) {
	case TYPE_DOUBLE:
		resultOfstrtod = strtod(toConvStr, &endptr);
		if (*endptr != '\0') { //変換に失敗した時(文字列中に変換不可能な文字があった時)
			return -3;
		}
		generatedHexStr = getInternalHexFromDouble(resultOfstrtod, &lenOfGeneratedStr);

		break;
	case TYPE_FLOAT:
		resultOfstrtof = (float)strtod(toConvStr, &endptr);
		if (*endptr != '\0') { //変換に失敗した時(文字列中に変換不可能な文字があった時)
			return -3;
		}
		generatedHexStr = getInternalHexFromFloat(resultOfstrtof, &lenOfGeneratedStr);
		break;

	case TYPE_LONG:
		resultOfstrtol = strtol(toConvStr, &endptr, 10);
		if (*endptr != '\0') { //変換に失敗した時(文字列中に変換不可能な文字があった時)
			return -3;
		}
		generatedHexStr = getInternalHexFromLong(resultOfstrtol, &lenOfGeneratedStr);
		break;

	default: //引数変換タイプが不明
		return -2;
		break;
	}

	if (generatedHexStr == NULL) { //一時文字列保存用メモリ確保失敗
		return -5;
	}

	lenOfWroteStr = lenOfGeneratedStr - 1;

	if (lenOfToWriteStr < (lenOfWroteStr)) { //書き込み先メモリ不足
		free(generatedHexStr); 
		return -4;
	
	}

	copyStr(toWriteStr, generatedHexStr, lenOfWroteStr);
	free(generatedHexStr);

	return lenOfWroteStr;
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
__declspec (dllexport) int WINAPI getSizeOfOperandExp(int operandType) {

	return getSizeOfOperand(operandType);

}

//
//hex内部表現の2数で四則演算する
//
//operandType
// 0:double
// 1:float
// 2:long
//
//operateType
// 0:addition
// 1:subtraction
// 2:multiplication
// 3:division
//
//返却値
// 1以上:計算後のhex文字列長
//-1:val1のhex文字列が不正
//-2:val2のhex文字列が不正
//-3:演算結果格納先文字列長が不足している
//-4:operandTypeが不正
//-5:operateTypeが不正
//-6:メモリ不足
//-7:整数型オペランドによる0除算
//
__declspec (dllexport) int WINAPI operateArithmeticByInternalHex(char* val1Str, char* val2Str, char* ansStr, int lenOfAns, int operandType, int operateType) {

	//変数宣言
	void* val1Ope;
	void* val2Ope;
	void* ansOpe;

	unsigned char requiredSize;
	unsigned char requiredLenOfToWrite;

	int retOfSetInternalHexToOperand;

	char* generatedHexStr;   //生成Hex文字列
	unsigned char    lenOfGeneratedStr; //生成Hex文字列のbyte長

	char is0Divide;


	//引数チェック
	if (operandType < MIN_OF_TYPE || MAX_OF_TYPE < operandType) {
		return -4;
	}
	if (operateType < OPERATE_TYPE_ADDITION || OPERATE_TYPE_DIVISION < operateType) {
		return -5;
	}

	//変数格納領域確保
	requiredSize = getSizeOfOperand(operandType); //必要バイト数取得
	requiredLenOfToWrite = requiredSize * 2;

	if (lenOfAns < (requiredLenOfToWrite)) { //演算結果格納先文字列長不足
		return -3;
	}

	val1Ope = (void*)malloc(requiredSize);
	if (val1Ope == NULL) { //メモリ確保失敗
		return -6;
	}

	val2Ope = (void*)malloc(requiredSize);
	if (val2Ope == NULL) { //メモリ確保失敗
		free(val1Ope);
		return -6;
	}

	ansOpe = (void*)malloc(requiredSize);
	if (ansOpe == NULL) { //メモリ確保失敗
		free(val1Ope);
		free(val2Ope);
		return -6;
	}

	//val1のオペランド獲得
	retOfSetInternalHexToOperand = setInternalHexToOperand(val1Str, val1Ope, operandType);

	if (retOfSetInternalHexToOperand != 0) { //格納結果確認
		free(val1Ope);
		free(val2Ope);
		free(ansOpe);
		return -1;
	}

	//val1のオペランド獲得
	retOfSetInternalHexToOperand = setInternalHexToOperand(val2Str, val2Ope, operandType);

	if (retOfSetInternalHexToOperand != 0) { //格納結果確認
		free(val1Ope);
		free(val2Ope);
		free(ansOpe);
		return -2;
	}

	//演算
	is0Divide = 0;
	switch (operandType) {
	case TYPE_DOUBLE:

		switch (operateType) {
		case OPERATE_TYPE_ADDITION:
			(*(double*)ansOpe) = (*(double*)val1Ope) + (*(double*)val2Ope);
			break;
		case OPERATE_TYPE_SUBSTRACTION:
			(*(double*)ansOpe) = (*(double*)val1Ope) - (*(double*)val2Ope);
			break;
		case OPERATE_TYPE_MULTIPLICATION:
			(*(double*)ansOpe) = (*(double*)val1Ope) * (*(double*)val2Ope);
			break;
		case OPERATE_TYPE_DIVISION:
			(*(double*)ansOpe) = (*(double*)val1Ope) / (*(double*)val2Ope);
			break;

		default: //引き数チェックで弾いているので、このルートは発生し得ない
			break;
		}

		break;

	case TYPE_FLOAT:

		switch (operateType) {
		case OPERATE_TYPE_ADDITION:
			(*(float*)ansOpe) = (*(float*)val1Ope) + (*(float*)val2Ope);
			break;
		case OPERATE_TYPE_SUBSTRACTION:
			(*(float*)ansOpe) = (*(float*)val1Ope) - (*(float*)val2Ope);
			break;
		case OPERATE_TYPE_MULTIPLICATION:
			(*(float*)ansOpe) = (*(float*)val1Ope) * (*(float*)val2Ope);
			break;
		case OPERATE_TYPE_DIVISION:
			(*(float*)ansOpe) = (*(float*)val1Ope) / (*(float*)val2Ope);
			break;

		default: //引き数チェックで弾いているので、このルートは発生し得ない
			break;
		}

		break;

	case TYPE_LONG:

		switch (operateType) {
		case OPERATE_TYPE_ADDITION:
			(*(long*)ansOpe) = (*(long*)val1Ope) + (*(long*)val2Ope);
			break;
		case OPERATE_TYPE_SUBSTRACTION:
			(*(long*)ansOpe) = (*(long*)val1Ope) - (*(long*)val2Ope);
			break;
		case OPERATE_TYPE_MULTIPLICATION:
			(*(long*)ansOpe) = (*(long*)val1Ope) * (*(long*)val2Ope);
			break;
		case OPERATE_TYPE_DIVISION:
			if ((*(long*)val2Ope) == 0) { //0除算チェック
				is0Divide = 1;
			}
			else {
				(*(long*)ansOpe) = (*(long*)val1Ope) / (*(long*)val2Ope);
			}
			break;

		default: //引き数チェックで弾いているので、このルートは発生し得ない
			break;
		}

		break;

	default: //変数型が不明。
			 //引き数チェックで弾いているので、このルートは発生し得ない
		break;

	}

	generatedHexStr = NULL;
	if (is0Divide == 0) { //0除算ではなかった
		generatedHexStr = getInternalHexFromOperand((unsigned char*)ansOpe, requiredSize, &lenOfGeneratedStr); //内部表現取得
	}

	//開放
	free(val1Ope);
	free(val2Ope);
	free(ansOpe);

	if(is0Divide){ //0除算
		return -7;

	} else if(generatedHexStr == NULL) { //メモリ不足
		return -6;
	
	}

	//文字列コピー
	copyStr(ansStr, generatedHexStr, requiredLenOfToWrite);
	free(generatedHexStr);

	return requiredLenOfToWrite; //計算後のhex文字列長を返す

}

//
//指定byte数分の文字をコピーする
//
static void copyStr(char* destStr, char* sourceStr, unsigned int bytesOfStr) {

	unsigned int cnt;

	//引数チェック
	if (destStr == NULL || sourceStr == NULL) {
		return;
	}

	//コピー
	for (cnt = 0; cnt <= bytesOfStr; cnt++) {
		destStr[cnt] = sourceStr[cnt];
	}

	return;

}

