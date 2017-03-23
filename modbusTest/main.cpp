/*
	数据采集脚本，12.27，可运行，未优化
	12.28 优化,修改 输入电流信号格式：int 2 double
	原始脚本，重要！
*/
#include "illusionExcelFile.h"
#include "DMT.h"
#include <iostream>
#include <stdio.h>
#include <opencv2/opencv.hpp>
#include <stdlib.h>
#include <time.h>
#include <string>
#include <sstream>
#include "GigECamera_Types.h"
#include "MVGigE.h"
#include "MVCamProptySheet.h"

using namespace std;
using namespace cv;

//图像类全局变量
HANDLE m_hCam = NULL;//相机句柄
MVImage m_image;//mv图像格式
MV_PixelFormatEnums m_PixelFormat;//图像像素格式
BOOL	m_bLive = FALSE;//是否采集标志
Mat	m_cvSrc;
Mat m_cvDst;
double fps = 15;//帧率（最大15）
unsigned long nExp_us = 35;//曝光时间（35~2260000us）
long paraForSet = 0;

int iNumFrame = 1;//帧计数
//modbus模块全局变量
const int comm_type = 0x00;// modbus通讯种类, 0:串行端口, 1:以太网络【不可更改】
const int conn_num = 0x03;// modbus串行端口编号, COM3【不可更改】
const int station_addr = 0x01;//从站地址0x01【不可更改】
char dataForRead[9] = "03000001";//读取输入电流的大小：0300寄存器地址，0001寄存器数量【不可更改】

stringstream tmp;//int转string中间变量
//stringstream tmp2;//int转string中间变量
string sDataInput;//产量信号string格式
CString csDataInput;//产量信号cstring格式（0-4095）
string sOutputValue;//输出信号
CString csOutputValue;//输出信号
//文档处理模块全局变量
int iSysTime[7] = {0};//系统时间
char cTime[50] = {0};//系统时间(字符串格式)
CString csTime;//系统时间(CString格式)
IllusionExcelFile excelProcessor;//excel处理类

//函数声明
LPCWSTR string2LPCWSTR(string orig);
void getTheTime(int *iSysTime);
int initModbus();
int setModbusOutput(unsigned char *data);
int readModbusInput(unsigned char *data);
void closeModbus();
int __stdcall StreamCB(MV_IMAGE_INFO *pInfo,long nUserVal);
int char2CString(char *in, CString &out);

int main(/*int argc,char *args[]*/)
{   
	//改
	char argv[4][100] = {0};
	strcpy(argv[0],"address");//地址

	/*------------------------------------------------------------------------------------------------------------*/
	/*------------------------------------------------------操作台-------------------------------------------------*/
	strcpy(argv[1],"8000");//曝光时间（35~2260000us）
	/*--------记录-------------------------------------*/
	strcpy(argv[2],"10.5");//流股直径（流速）
	strcpy(argv[3],"030404FF");//下落点坐标报文，与模块输出电流有关(0-4095)0000-0FFF

	int mod = 1;//1.2.3
	/*------------------------------------------------------------------------------------------------------------*/
	string mfile;
	if( mod == 1)
	{
		mfile = "mod1\\";
	}
	else if (mod ==2)
	{
		mfile = "mod2\\";
	}
	else if (mod ==3)
	{
		mfile = "mod3\\";
	}

	
	//①读取输入参数（曝光值-流速-下落点）
	cout<<"--参数初始化"<<endl;
	nExp_us = (unsigned long)atol(argv[1]);//第一个参数：曝光值 字符数组转换为unsigned long

	double dDiameterOfStream=atof(argv[2]);//第二个参数：流股直径（流速）控制值转换为CString格式
	
	CString csDiameterOfStream;
	char2CString(argv[2], csDiameterOfStream);//把流股直径转换为CString格式

	//第三个参数：下落点位置(0-4095)控制值字符数组转换为int
	char cOutputValue[100] = {0};//提取输出电流值(0-4095)
	for(int k = 4;k < 9;k++)
	{
		cOutputValue[k-4] = argv[3][k];
	}
	char *tmpPtr = NULL;
	int iFallingPosition;
	iFallingPosition = strtol(cOutputValue, &tmpPtr, 16);//转为10进制
	cout<<cOutputValue<<":"<<iFallingPosition<<endl;
	double dOutputValue = ((double)iFallingPosition * 16.0 / 4096.0) + 4.0;//输出电流值mA
	tmp << dOutputValue;
	tmp >> sOutputValue;
	csOutputValue = sOutputValue.c_str();//string转CString
	
	//②生成视频名称
	const string file = "C:\\Users\\Administrator\\Desktop\\workspace\\data\\";//视频存储路径
	string a1 = argv[1] ;string a2 = argv[2] ;string a3 = argv[3] ;//中间变量
	string video_name = a1 +"_" + a2 +"_" + sOutputValue + ".avi" ;//视频名称：曝光时间_流速（流股直径）_输出电流值（下落点坐标）
	string video_name_time = a1 +"_" + a2 +"_" + sOutputValue + "(time).avi" ;//时间戳视频名称：曝光时间_流速（流股直径）_输出电流值（下落点坐标）
	string video_file = file + mfile + video_name;//视频路径+名称
	string video_file_time = file + mfile + video_name_time;//时间戳视频路径+名称
	//生成excel名称
	string excel_name = a1 +"_" + a2 +"_" + sOutputValue + ".xlsx";//要生成的excel文件名
	string excel_file = file + mfile + excel_name;//要生成的excel文件全路径

	string init_excel_name = "init.xlsx";//种子excel文件名
	string init_excel_file = file + init_excel_name;//种子excel文件全路径

	//③excel处理模块初始化
	cout<<"--excel处理模块初始化"<<endl;
	bool bInit0 = excelProcessor.InitExcel();//初始化excel处理类

	LPCWSTR L_init_excel_file = string2LPCWSTR(init_excel_file);//文件名string转换为LPCWSTR
	bool bRet0 = excelProcessor.OpenExcelFile(L_init_excel_file);//打开种子excel文件
	CString strSheetName = excelProcessor.GetSheetName(1);//获取第一工作区的名称
	bool bLoad = excelProcessor.LoadSheet(strSheetName);//根据名称打开该工作区
	string textHead[5] = {"记录时间","流股直径（mm）","流股厚度（mm）","流股落点控制信号（mA）","产量信号（mA）"};//生成表头
	for(int j=1;j < 6;j++)//在第一行输入表头
	{
		CString text;
		text = textHead[j - 1].c_str();
		excelProcessor.SetCellString(1,j,text);
	}
	


	//④modbus初始化
	cout<<"--modbus协议模块初始化"<<endl;
	int isInit = initModbus();//初始化
	if(isInit == -1)
	{
		cout<<"modbus初始化失败"<<endl;
		return 0;
	}

	//⑤设置输出电流大小（设置报文）
	cout<<"--modbus设置输出电流"<<endl;
	unsigned char setData[1024] = {0};//输出电流数据
	int sendlen = 0;//索引
	char A2H[3] = { 0 };
	for (int i = 0; i < 8; i += 2)//char数组转换为unsigned char数组
	{
		int num = 0;
		A2H[0] = argv[3][i];
		A2H[1] = argv[3][i+1];
		sscanf(A2H, "%02X", &num);
		setData[sendlen++] = num;
	}

	int isSet = setModbusOutput(setData);//设置输出，0304寄存器地址，0FFF电流大小
	if (isSet == -1)
	{
		cout<<"设置输出电流失败"<<endl;
		return 0;
	}

	//设置输出电流的报文
	//char dataForRead[9] = "03000001";
	unsigned char uDataForRead[1024] = { 0 };//输出电流数据
	int sendlen2 = 0;
	char A2H2[3] = { 0 };
	for (int k = 0; k < 8; k += 2)
	{
		int num = 0;
		A2H2[0] = dataForRead[k];
		A2H2[1] = dataForRead[k + 1];
		sscanf(A2H2, "%02X", &num);
		uDataForRead[sendlen2++] = num;
	}


	//⑥MV初始化
	cout<<"--MV操作模块初始化"<<endl;
	MVSTATUS_CODES mvInit;
	mvInit = MVInitLib();//初始化
	if(mvInit != MVST_SUCCESS)
	{
		cout<<"相机库初始化失败，请检测环境"<<endl;
		return 0;
	}
	int nCams;//相机数量
	MVGetNumOfCameras(&nCams);//获取相机总数量nCams
	if( nCams == 0 )
	{
		cout<<"没有找到相机,请确认连接和相机IP设置"<<endl;
		return 0;
	}
	MVSTATUS_CODES r = MVOpenCamByIndex(0,&m_hCam);//通过索引号打开索引号为0的相机
	if( m_hCam == NULL )//相机句柄为空
	{
		if( r == MVST_ACCESS_DENIED )//未能打开相机
		{
			cout<<"无法打开相机，可能正被别的软件控制"<<endl;
			return 0;
		}
	}
	MVSetExposureTime (m_hCam,nExp_us);//设置曝光时间
	MVSetFrameRate (m_hCam,fps);//设置当前帧率

	int w,h;//根据相机句柄获得宽和高信息/图像像素格式			
	MVGetWidth(m_hCam,&w);//宽
	MVGetHeight(m_hCam,&h);//高
	MVGetPixelFormat(m_hCam,&m_PixelFormat);//获取图像像素格式

	if( m_PixelFormat == PixelFormat_Mono8 )//如果是灰度图
	{
		m_image.Create(w,h,8);
		m_cvSrc.create(h,w,CV_8UC1);//创建单通道Mat
	}
	else
	{
		m_image.Create(w,h,24);
		m_cvSrc.create(h,w,CV_8UC3);//创建三通道彩色Mat
	}
	m_cvDst = m_cvSrc.clone();
	
	//视频存储初始化
	CvSize size =  m_cvDst.size();//参数：视频尺寸
	VideoWriter outputVideo(video_file,CV_FOURCC('D', 'I', 'V', 'X'),/*帧率*/3,size,1/*彩色1黑白0*/);
	VideoWriter outputVideo_time(video_file_time,CV_FOURCC('D', 'I', 'V', 'X'),/*帧率*/3,size,1/*彩色1黑白0*/);
	cout<<"视频分辨率："<<w<<"x"<<h<<endl;
	cout<<"曝光："<<nExp_us<<endl;
	cout<<"帧率："<<fps<<endl;

	//开始读取每一帧
	cout<<"--初始化完毕！开始采集;"<<endl;
	int excel_y = 2;//表格纵轴起始点
	while(true)
	{
		iNumFrame++;//帧序号统计,从2开始
		getTheTime(iSysTime);//获取当前时间
		//1写入时间
		sprintf(cTime,"%dYear%dMon%dDay %d:%d:%d.%d",iSysTime[0],iSysTime[1],iSysTime[2],iSysTime[3],iSysTime[4],iSysTime[5],iSysTime[6]);
		char2CString(cTime, csTime);//格式转换
		if((iNumFrame % 1) == 0)//每1帧在表格储存一次数据
		{
			excelProcessor.SetCellString(excel_y,1,csTime);
			//2写入流股直径
			excelProcessor.SetCellString(excel_y,2,csDiameterOfStream);
			//3流股厚度置空

			//4流股落点位置存入电流信号（ma）csOutputValue
			excelProcessor.SetCellString(excel_y,4,csOutputValue);
			//读取此刻的输入电流
			int isRead = readModbusInput(uDataForRead);//读取此刻的输入电流值（产量信号）
			
			double dInputValue = ((double)isRead / 4096.0  * 16.0) + 4.0;
			//5写入当前时刻产量信号
			stringstream tmp2;
			tmp2 << dInputValue;//数字转中间变量
			tmp2 >> sDataInput;//中间变量转string
			 cout<<"int"<<isRead<<"double"<<dInputValue<<"string"<<sDataInput<<endl;
			csDataInput = sDataInput.c_str();//string转CString
			excelProcessor.SetCellString(excel_y,5,csDataInput);
			excel_y++;
		}

		MVStartGrab(m_hCam,StreamCB,paraForSet);//采集
		outputVideo << m_cvSrc;//源视频保存

		string text(cTime);//时间戳
		putText(m_cvSrc,text,Point(100,100),FONT_HERSHEY_PLAIN,4.0,Scalar(255, 255, 255),3.0);//加入时间戳
		outputVideo_time << m_cvSrc;//时间戳视频保存

		m_cvDst = m_cvSrc.clone();	//获得屏幕输出的Mat
		resize(m_cvDst,m_cvDst,Size(),0.5,0.5,1);
		imshow("视频",m_cvDst);
		

		if(waitKey(200) == 27 )
		{
			MVStopGrab(m_hCam);//停止抓取
			closeModbus();
			if( m_hCam != NULL )
			{
				MVCloseCam(m_hCam);//关闭摄像头
			}
			MVTerminateLib();//释放资源
			break;
		}
	}
	//存储对应的表格
	LPCWSTR L_excel_file = string2LPCWSTR(excel_file);//把转存excel名string转为LPCWSTR
	excelProcessor.SaveasXSLFile(L_excel_file);
	excelProcessor.CloseExcelFile(1);
	excelProcessor.ReleaseExcel(); 

	cout<<"完毕!存储帧数："<<iNumFrame + 1<<endl;
	cout<<"----------------------------------------------"<<endl;
	return 0;
}



int initModbus()//初始化modbus，开启串行通讯，若成功则返回1，失败则返回-1
{
	int status  = OpenModbusSerial(conn_num , 115200, 8, 'N', 1, 2);
	if (status == -1)
	{
		return -1;
	}
	else
	{
		return 0;
	}
}
int setModbusOutput(unsigned char *data /*"03040FFF"*/)//设置电流输出(0-4095)，从而控制下落点坐标,成功返回1，失败返回-1
{
	const int func_code_6_AO = 0x06;//写入输出电流 功能码
	int recieve_station_addr = station_addr;//一会待返回的地址
	int recieve_func_code_6_AO = func_code_6_AO;//一会待返回的功能码
	unsigned char recieve_data[1024] = {0};//待接受的数据
	//通信
	int datalen2 = RequestData(comm_type, conn_num, station_addr, func_code_6_AO,data, 4); // 数据请求
	if(datalen2 == -1)
	{
		return -1;//请求失败
	}

	//接受
	int bytecount2 = ResponseData(comm_type, conn_num, &recieve_station_addr, &recieve_func_code_6_AO, recieve_data);    // 数据回复
	//验证
	if(bytecount2 > 0)
	{
		if(recieve_func_code_6_AO !=  func_code_6_AO)//如果回复的功能码不同 则为发生了错误
		{
			
			return -1;//写值失败
		}
		else
		{
			return 1;//写值成功
		}
	}
	else
	{
		return -1;//未收到数据
	}
}

int readModbusInput(unsigned char *data/*"03000001"*/)//获取电流输入值（0-4095），并返回相应int值,失败返回-1
{
	const int func_code_3_AI = 0x03;//读取输入电流 功能码
	int recieve_station_addr = station_addr;//一会待返回的地址
	int recieve_func_code_3_AI = func_code_3_AI;//一会待返回的功能码
	unsigned char recieve_data[1024] = {0};//待接受的数据

	int datalen1 = RequestData(comm_type, conn_num, station_addr, func_code_3_AI,data , 4); // 数据请求
	if(datalen1 == -1)
	{
		return -1;//请求失败
	}
	int bytecount1 = ResponseData(comm_type, conn_num, &recieve_station_addr, &recieve_func_code_3_AI, recieve_data);    // 数据回复
	if(bytecount1 > 0)
	{
		if(recieve_func_code_3_AI !=  func_code_3_AI)//如果回复的功能码不同 则为发生了错误
		{
			return -1;//收到数据，但读取失败
		}
		else//如果回复的功能码没有变化 则说明正常
		{//将收到的数据recieve_data(unsigned char)转换为int型十进制数据
			char A2H[3] = {0};
			char s[1024] = {0};
			int index = 0;
			for (int i = 0; i < bytecount1; ++i) // recover a string from recvbuf
			{
				sprintf(&s[index], "%02X", recieve_data[i]);
				index += 2;;
			}

			char *endptr = NULL;
			int iElectValue;
			iElectValue = strtol(s, &endptr, 16);
			return iElectValue;
		}
	}
	else
	{
		return -1;//未收到数据
	}

}

void closeModbus()
{
	CloseSerial(conn_num); // 关闭通讯 
}

void getTheTime(int *iSysTime)//获取当前的系统时间
{
	SYSTEMTIME sys;
	GetLocalTime( &sys );
	iSysTime[0] = sys.wYear;//年
	iSysTime[1] = sys.wMonth;//月
	iSysTime[2] = sys.wDay;//日
	iSysTime[3] = sys.wHour;//时
	iSysTime[4] = sys.wMinute;//分
	iSysTime[5] = sys.wSecond;//秒
	iSysTime[6] = sys.wMilliseconds;//毫秒
}

LPCWSTR string2LPCWSTR(string orig)//把string转换为LPCWSTR
{
	size_t origsize = orig.length() + 1;
    const size_t newsize = 100;
    size_t convertedChars = 0;
	wchar_t *wcstring = (wchar_t *)malloc(sizeof(wchar_t)*(orig.length()-1));
	mbstowcs_s(&convertedChars, wcstring, origsize, orig.c_str(), _TRUNCATE);
	return wcstring;
}

int __stdcall StreamCB(MV_IMAGE_INFO *pInfo,long nUserVal)//MV借调函数,将MV格式数据转换为MAT
{
	//获取图像的尺寸
	int w = m_image.GetWidth();
	int h = m_image.GetHeight();

	if( m_PixelFormat == PixelFormat_Mono8 )//灰度图处理
	{
		ASSERT( m_image.GetPitch() == m_image.GetWidth() );
		memcpy(m_cvSrc.data,pInfo->pImageBuffer,w*h);
		
	}
	else//彩图
	{
		MVBayerToBGR(m_hCam,pInfo->pImageBuffer,m_cvSrc.data,w*3,w,h,m_PixelFormat);
	}
	return 0;
}

int char2CString(char *in, CString &out)//字符数组转换为CString
{
	if(NULL == in)
	{
		cout<<"int char2CString(char *in, CString &out) 错误：输入为空"<<endl;
		return -1;
	}
	string argv2(in);//先转为string
	out = argv2.c_str();//转为CString
	return 0;
}