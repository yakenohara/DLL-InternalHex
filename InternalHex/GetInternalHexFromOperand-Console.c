#include <stdio.h>
#include <stdlib.h>
#include <string.h>

#include "GetInternalHexFromOperand.h"
#include "PrintCursorInLine.h"

#define OPTION_SPECIFIER ('/')
#define OPTION_DOUBLE ('d')
#define OPTION_FLOAT ('f')
#define OPTION_LONG ('l')

//
//Main
//
int main( int argc, char *argv[] ) {
	
	//変数宣言
	char option; //変換オプション
	
	unsigned char indexOfValueInArg;  //変換対象文字列のargv[]内index
	unsigned char indexOfOptionInArg; //変換オプション指定のargv[]内index
	
	unsigned char len; //Bin変換後文字列長
	char* txt;         //Bin変換後文字列
	
	unsigned char argcnt; //argv[]検索ループ用カウンタ
	
	float resultOfstrtof;  //文字列→変数変換後用一時格納先(float)
	double resultOfstrtod; //文字列→変数変換後用一時格納先(double)
	long resultOfstrtol;   //文字列→変数変換後用一時格納先(long)
	
	char *endptr;          //strtoXコール用
	
	//デフォルト設定
	option = OPTION_DOUBLE;
	
	//引数チェック
	if( argc == 1 ){ //引数指定がない場合
		printf("Invalid Argments");
		return 1; //終了
	}
	
	//オプションチェック
	indexOfOptionInArg = 0;
	if( argc >= 2){ //引数が2つ以上の場合
		
		//オプション指定子を探す
		for(argcnt = 1 ; argcnt < argc ; argcnt++){
			if (argv[argcnt][0] == OPTION_SPECIFIER){ //オプション指定子の場合
				
				indexOfOptionInArg = argcnt;
				
				//オプション内容チェック
				if(strlen(argv[argcnt]) != 2){ //オプション指定子含めて2文字でないといけない
					printf("Invalid Option");
					return 1;
					
				}else{
					switch(argv[argcnt][1]){
						case OPTION_DOUBLE:
							option = OPTION_DOUBLE;
							break;
						
						case OPTION_FLOAT:
							option = OPTION_FLOAT;
							break;
						
						case OPTION_LONG:
							option = OPTION_LONG;
							break;
						
						default:
							printf("Unknown Option");
							return 1;
							break;
					}
				}
				break;
			}
		}
		
	}
	
	//変換対象文字列存在チェック
	if(indexOfOptionInArg >= 1){ //オプション指定有の時
		if (argc != 3){ //変換対象文字列指定無しの場合
			printf("Value not found in argments");
			return 1;
			
		}
		
		if(indexOfOptionInArg == 2){
			indexOfValueInArg = 1;
			
		}else{
			indexOfValueInArg = 2;
		}
		
	}else{ //オプション指定無しの場合
		if (argc != 2){ //変換対象文字列指定無しの場合
			printf("Value not found in argments");
			return 1;
			
		}
		
		indexOfValueInArg = 1;
	}
	
	//変換処理
	switch(option){
		case OPTION_DOUBLE:
			resultOfstrtod = strtod(argv[indexOfValueInArg], &endptr); //変換
			if (*endptr != '\0'){ //変換に失敗した時(文字列中に変換不可能な文字があった時)
				printf("Cannot convert to double\n");
				printCursorInLine(argv[indexOfValueInArg], endptr);
				return 1;
			}
			txt = getInternalHexFromDouble(resultOfstrtod, &len);
			
			break;
		
		case OPTION_FLOAT:
			resultOfstrtof = (float)strtod(argv[indexOfValueInArg], &endptr); //変換
			if (*endptr != '\0'){ //変換に失敗した時(文字列中に変換不可能な文字があった時)
				printf("Cannot convert to float\n");
				printCursorInLine(argv[indexOfValueInArg], endptr);
				return 1;
			}
			txt = getInternalHexFromFloat(resultOfstrtof, &len);
			
			break;
		
		case OPTION_LONG:
			resultOfstrtol = strtol(argv[indexOfValueInArg], &endptr, 10); //変換
			if (*endptr != '\0'){ //変換に失敗した時(文字列中に変換不可能な文字があった時)
				printf("Cannot convert to long\n");
				printCursorInLine(argv[indexOfValueInArg], endptr);
				return 1;
			}
			txt = getInternalHexFromLong(resultOfstrtol, &len);
			
			break;
		
		default:
			printf("Unknown Error");
			return 1;
			break;
	}
	
	if(txt == NULL){ //メモリ確保失敗
		printf("Cannot allocate memory");
		return 1;
	}
	printf("%s", txt);
	free(txt);
	
	return 0;
}

