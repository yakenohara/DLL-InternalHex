#ifndef GET_INTERNAL_HEX_FROM_OPERAND_H
#define GET_INTERNAL_HEX_FROM_OPERAND_H

#ifdef __cplusplus
extern "C" {
#endif /* __cplusplus */

extern char* getInternalHexFromDouble(double val, unsigned char* len);
extern char* getInternalHexFromFloat(float val, unsigned char* len);
extern char* getInternalHexFromLong(long val, unsigned char* len);
extern char* getInternalHexFromOperand(unsigned char* val, unsigned char bytesOfVal, unsigned char* len);

#ifdef __cplusplus
}
#endif /* __cplusplus */

#endif // !GET_INTERNAL_HEX_FROM_OPERAND_H
