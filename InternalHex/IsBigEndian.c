//Function Prototype
char isBigEndian(void);

//
//実行環境がビッグエンディアンかどうかを返す
//ビッグエンディアンの場合は1
//リトルエンディアンの場合は0を返す
//
char isBigEndian(void) {

	int x = 1;
	char ret;

	if ( *(char *)&x ){ //Little Endian
		ret = 0;

	}else{ //Big Endian

		ret = 1;
	}

	return ret;
	
}
