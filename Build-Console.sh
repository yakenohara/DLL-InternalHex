i686-w64-mingw32-gcc \
    -o ./Console/GetInternalHexFromOperand-Console.exe \
    -I./InternalHex \
    ./InternalHex/GetInternalHexFromOperand-Console.c \
    ./InternalHex/GetInternalHexFromOperand.c \
    ./InternalHex/IsBigEndian.c \
    ./InternalHex/PrintCursorInLine.c