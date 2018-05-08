#include <stdio.h>

void printCursorInLine(char* line, char* here);

void printCursorInLine(char* line, char* here){
	
	//変数宣言
	char* address;
	
	//引数チェック
	if( (line == NULL) || (here == NULL)){
		return;
	}
	
	printf("%s\n", line);
	
	for(address = line ; address < here ; address++){
		printf("%c",' ');
	}
	
	printf("^");

}
