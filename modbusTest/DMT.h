#include <windows.h>
/////////////////////////////////////////////////////////////////////////
//
//	Name:
//		DMT.h
//
//	Description:
//		DMT Library Header File
//
//	History:
//		Date			Author			Plant			Version			Comment
//      13/02/2012		AllenCX			Taoyuan III		Version 2.2
//		05/01/2009		Anderson		Taoyuan I		Version 2.0
//		10/01/2007		Anderson		Taoyuan I		Version 1.1
//		08/01/2007		Anderson		Taoyuan I		Version 1.0
//	
/////////////////////////////////////////////////////////////////////////

// Export functions
#ifdef	__cplusplus
extern "C"
{
#endif	/*  __cplusplus */

// Data Access
int WINAPI RequestData(int comm_type, int conn_num, int slave_addr, int func_code, const unsigned char *sendbuf, int sendlen);
int WINAPI ResponseData(int comm_type, int conn_num, int *slave_addr, int *func_code, unsigned char *recvbuf);

// Serial Communication
int WINAPI OpenModbusSerial(int conn_num, int baud_rate, int data_len, char parity, int stop_bits, int modbus_mode);
void WINAPI CloseSerial(int conn_num);
int WINAPI GetLastSerialErr();
void WINAPI ResetSerialErr();

// Socket Communication
int WINAPI OpenModbusTCPSocket(int conn_num, int ipaddr, int port = 502);
void WINAPI CloseSocket(int conn_num);
int WINAPI GetLastSocketErr();
void WINAPI ResetSocketErr();
int WINAPI ReadSelect(int conn_num, int millisecs);

// MODBUS Address Calculation
int WINAPI DevToAddrA(const char *series, const char *device, int qty);
int WINAPI DevToAddrW(const wchar_t *series, const wchar_t *device, int qty);

// Wrapped MODBUS Funcion : 0x01
int WINAPI ReadCoilsA(int comm_type, int conn_num, int slave_addr, int dev_addr, int qty,
					  unsigned int *data_r, char *req, char *res);

int WINAPI ReadCoilsW(int comm_type, int conn_num, int slave_addr, int dev_addr, int qty,
					  unsigned int *data_r, wchar_t *req, wchar_t *res);

// Wrapped MODBUS Funcion : 0x02
int WINAPI ReadInputsA(int comm_type, int conn_num, int slave_addr, int dev_addr, int qty,
					   unsigned int *data_r, char *req, char *res);

int WINAPI ReadInputsW(int comm_type, int conn_num, int slave_addr, int dev_addr, int qty,
					   unsigned int *data_r, wchar_t *req, wchar_t *res);

// Wrapped MODBUS Funcion : 0x03
int WINAPI ReadHoldRegsA(int comm_type, int conn_num, int slave_addr, int dev_addr, int qty,
						 unsigned int *data_r, char *req, char *res);

int WINAPI ReadHoldRegsW(int comm_type, int conn_num, int slave_addr, int dev_addr, int qty,
						 unsigned int *data_r, wchar_t *req, wchar_t *res);

int WINAPI ReadHoldRegs32A(int comm_type, int conn_num, int slave_addr, int dev_addr, int qty,
						   unsigned int *data_r, char *req, char *res);

int WINAPI ReadHoldRegs32W(int comm_type, int conn_num, int slave_addr, int dev_addr, int qty, 
						   unsigned int *data_r, wchar_t *req, wchar_t *res);

// Wrapped MODBUS Funcion : 0x04
int WINAPI ReadInputRegsA(int comm_type, int conn_num, int slave_addr, int dev_addr, int qty,
						  unsigned int *data_r, char *req, char *res);

int WINAPI ReadInputRegsW(int comm_type, int conn_num, int slave_addr, int dev_addr, int qty,
						  unsigned int *data_r, wchar_t *req, wchar_t *res);

// Wrapped MODBUS Funcion : 0x05		   
int WINAPI WriteSingleCoilA(int comm_type, int conn_num, int slave_addr, int dev_addr, 
							unsigned int data_w, char *req , char *res);

int WINAPI WriteSingleCoilW(int comm_type, int conn_num, int slave_addr, int dev_addr, 
							unsigned int data_w, wchar_t *req , wchar_t *res);

// Wrapped MODBUS Funcion : 0x06
int WINAPI WriteSingleRegA(int comm_type, int conn_num, int slave_addr, int dev_addr, 
						   unsigned int data_w, char *req , char *res);

int WINAPI WriteSingleRegW(int comm_type, int conn_num, int slave_addr, int dev_addr, 
						   unsigned int data_w, wchar_t *req , wchar_t *res);

int WINAPI WriteSingleReg32A(int comm_type, int conn_num, int slave_addr, int dev_addr, 
							 unsigned int data_w, char *req , char *res);

int WINAPI WriteSingleReg32W(int comm_type, int conn_num, int slave_addr, int dev_addr, 
							 unsigned int data_w, wchar_t *req , wchar_t *res);

// Wrapped MODBUS Funcion : 0x0F
int WINAPI WriteMultiCoilsA(int comm_type, int conn_num, int slave_addr, int dev_addr, int qty,
							const unsigned int *data_w, char *req, char *res);

int WINAPI WriteMultiCoilsW(int comm_type, int conn_num, int slave_addr, int dev_addr, int qty,
							const unsigned int *data_w, wchar_t *req, wchar_t *res);

// Wrapped MODBUS Funcion : 0x10
int WINAPI WriteMultiRegsA(int comm_type, int conn_num, int slave_addr, int dev_addr, int qty,
						   const unsigned int *data_w, char *req, char *res);

int WINAPI WriteMultiRegsW(int comm_type, int conn_num, int slave_addr, int dev_addr, int qty,
						   const unsigned int *data_w, wchar_t *req, wchar_t *res);

int WINAPI WriteMultiRegs32A(int comm_type, int conn_num, int slave_addr, int dev_addr, int qty, 
							 const unsigned int *data_w, char *req, char *res);

int WINAPI WriteMultiRegs32W(int comm_type, int conn_num, int slave_addr, int dev_addr, int qty,
						   const unsigned int *data_w, wchar_t *req, wchar_t *res);



#ifdef	__cplusplus
}
#endif	/*  __cplusplus */
