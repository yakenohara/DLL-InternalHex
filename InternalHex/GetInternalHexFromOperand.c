#include <stdio.h>
// sizeof
// snprintf
// NULL

#include <stdlib.h>
// malloc

#include <string.h>
// memset

#include "IsBigEndian.h"

//Function Prototypes
char* getInternalHexFromDouble(double val, unsigned char* len);
char* getInternalHexFromFloat(float val, unsigned char* len);
char* getInternalHexFromLong(long val, unsigned char* len);
char* getInternalHexFromOperand(unsigned char* val, unsigned char bytesOfVal, unsigned char* len);
static unsigned char writeHexfromChar(char* toWrite, unsigned char ch);
static char getCharFrom4Bits(unsigned char bits);


//
//1st引数の内部表現(Hex列)を文字列で返す
//2nd引数には文字列長を入れる
//
//以下の場合はNULLを返す
//・2nd引数がNULLの場合
//・文字列を格納するメモリ確保失敗の場合
//
char* getInternalHexFromDouble(double val, unsigned char* len) {

	return getInternalHexFromOperand((unsigned char*)&val, sizeof(val), len);

}

//
//1st引数の内部表現(Hex列)を文字列で返す
//2nd引数には文字列長を入れる
//
//以下の場合はNULLを返す
//・2nd引数がNULLの場合
//・文字列を格納するメモリ確保失敗の場合
//
char* getInternalHexFromFloat(float val, unsigned char* len) {

	return getInternalHexFromOperand((unsigned char*)&val, sizeof(val), len);

}

//
//1st引数の内部表現(Hex列)を文字列で返す
//2nd引数には文字列長を入れる
//
//以下の場合はNULLを返す
//・2nd引数がNULLの場合
//・文字列を格納するメモリ確保失敗の場合
//
char* getInternalHexFromLong(long val, unsigned char* len) {

	return getInternalHexFromOperand((unsigned char*)&val, sizeof(val), len);

}

//共通関数
char* getInternalHexFromOperand(unsigned char* val, unsigned char bytesOfVal, unsigned char* len) {

	//変数宣言
	unsigned char   bytesToAlloc;     //返却文字列のByte長
	         char   relativeAddress;  //アドレス毎読み込みループ用カウンタ
	         char*  hexStr;           //返却文字列
	unsigned char   indexOfHexStr;    //返却文字列参照用
	         char   addressDirection; //アドレス読み込み方向
	unsigned char   loopLimit;        //ループ制御用

	//引数チェック
	if (len == NULL) {
		return NULL;
	}

	bytesToAlloc = bytesOfVal * 2 + 1; //1バイトは2文字で表現される(+1 は文字列最後の'\0'用)

	//Alloc
	hexStr = (char*)malloc(bytesToAlloc);
	if (hexStr == NULL) { //メモリ確保失敗時
		return NULL;
	}

	memset(hexStr, '\0', bytesToAlloc); //'\0'で初期化

	//Address読み込み方向
	if (isBigEndian()) { //Big Endianの場合
		addressDirection = 1;
		relativeAddress = 0u; //読み込み開始位置は最小Addressから

	}else { //Little Endianの場合
		addressDirection = -1;
		relativeAddress = (bytesOfVal - 1u); //読み込み開始位置は最大Addressから

	}
	
	//文字列生成
	loopLimit = bytesToAlloc - 1;
	for (indexOfHexStr = 0; indexOfHexStr < loopLimit ; indexOfHexStr += 2) {
		writeHexfromChar(&hexStr[indexOfHexStr], val[relativeAddress]); //2文字毎に記録
		relativeAddress += addressDirection;

	}

	*len = bytesToAlloc;
	return hexStr;
}

//
//2nd引数のBit表現をhex文字で1st引数に返す
//終了時に0を返す
//2nd引数がNULLの場合は、1を返す。
//
static unsigned char writeHexfromChar(char* toWrite, unsigned char ch) {

	//引数チェック
	if (toWrite == NULL) {
		return 1;
	}

	//上位4ビット取得
	toWrite[0] = getCharFrom4Bits(ch >> 4);

	//下位4ビット取得
	toWrite[1] = getCharFrom4Bits(ch & 0xF);

	return 0;
}


//
//0~16の数値からhex表現文字を返す
//0~17以外が指定された場合は、'G'を返す
//
static char getCharFrom4Bits(unsigned char bits) {

	char chToRet;

	chToRet = 'G'; //判別不能を設定(仮

	if (0u <= bits && bits < 10u) { //'0'~'9'の時
		chToRet = '0' + bits;

	}
	else if (10u <= bits && bits <= 16u) { //'A'~'F'の時
		chToRet = 'A' + bits - 10u;

	}

	return chToRet;
}
