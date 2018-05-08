#ifndef SET_INTERNAL_HEX_TO_OPERAND_H
#define SET_INTERNAL_HEX_TO_OPERAND_H

#ifdef __cplusplus
extern "C" {
#endif /* __cplusplus */

#define TYPE_DOUBLE (0U)
#define TYPE_FLOAT (1U)
#define TYPE_LONG (2U)

#define MIN_OF_TYPE (0U)
#define MAX_OF_TYPE (2U)

extern int setInternalHexToOperand(char* hexStr, void* toWrite, int operandType);
extern int getSizeOfOperand(int type);

#ifdef __cplusplus
}
#endif /* __cplusplus */

#endif // !SET_INTERNAL_HEX_TO_OPERAND_H

