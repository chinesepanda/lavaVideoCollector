/*
	���ݲɼ��ű���12.27�������У�δ�Ż�
	12.28 �Ż�,�޸� ��������źŸ�ʽ��int 2 double
	ԭʼ�ű�����Ҫ��
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

//ͼ����ȫ�ֱ���
HANDLE m_hCam = NULL;//������
MVImage m_image;//mvͼ���ʽ
MV_PixelFormatEnums m_PixelFormat;//ͼ�����ظ�ʽ
BOOL	m_bLive = FALSE;//�Ƿ�ɼ���־
Mat	m_cvSrc;
Mat m_cvDst;
double fps = 15;//֡�ʣ����15��
unsigned long nExp_us = 35;//�ع�ʱ�䣨35~2260000us��
long paraForSet = 0;

int iNumFrame = 1;//֡����
//modbusģ��ȫ�ֱ���
const int comm_type = 0x00;// modbusͨѶ����, 0:���ж˿�, 1:��̫���硾���ɸ��ġ�
const int conn_num = 0x03;// modbus���ж˿ڱ��, COM3�����ɸ��ġ�
const int station_addr = 0x01;//��վ��ַ0x01�����ɸ��ġ�
char dataForRead[9] = "03000001";//��ȡ��������Ĵ�С��0300�Ĵ�����ַ��0001�Ĵ������������ɸ��ġ�

stringstream tmp;//intתstring�м����
//stringstream tmp2;//intתstring�м����
string sDataInput;//�����ź�string��ʽ
CString csDataInput;//�����ź�cstring��ʽ��0-4095��
string sOutputValue;//����ź�
CString csOutputValue;//����ź�
//�ĵ�����ģ��ȫ�ֱ���
int iSysTime[7] = {0};//ϵͳʱ��
char cTime[50] = {0};//ϵͳʱ��(�ַ�����ʽ)
CString csTime;//ϵͳʱ��(CString��ʽ)
IllusionExcelFile excelProcessor;//excel������

//��������
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
	//��
	char argv[4][100] = {0};
	strcpy(argv[0],"address");//��ַ

	/*------------------------------------------------------------------------------------------------------------*/
	/*------------------------------------------------------����̨-------------------------------------------------*/
	strcpy(argv[1],"8000");//�ع�ʱ�䣨35~2260000us��
	/*--------��¼-------------------------------------*/
	strcpy(argv[2],"10.5");//����ֱ�������٣�
	strcpy(argv[3],"030404FF");//��������걨�ģ���ģ����������й�(0-4095)0000-0FFF

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

	
	//�ٶ�ȡ����������ع�ֵ-����-����㣩
	cout<<"--������ʼ��"<<endl;
	nExp_us = (unsigned long)atol(argv[1]);//��һ���������ع�ֵ �ַ�����ת��Ϊunsigned long

	double dDiameterOfStream=atof(argv[2]);//�ڶ�������������ֱ�������٣�����ֵת��ΪCString��ʽ
	
	CString csDiameterOfStream;
	char2CString(argv[2], csDiameterOfStream);//������ֱ��ת��ΪCString��ʽ

	//�����������������λ��(0-4095)����ֵ�ַ�����ת��Ϊint
	char cOutputValue[100] = {0};//��ȡ�������ֵ(0-4095)
	for(int k = 4;k < 9;k++)
	{
		cOutputValue[k-4] = argv[3][k];
	}
	char *tmpPtr = NULL;
	int iFallingPosition;
	iFallingPosition = strtol(cOutputValue, &tmpPtr, 16);//תΪ10����
	cout<<cOutputValue<<":"<<iFallingPosition<<endl;
	double dOutputValue = ((double)iFallingPosition * 16.0 / 4096.0) + 4.0;//�������ֵmA
	tmp << dOutputValue;
	tmp >> sOutputValue;
	csOutputValue = sOutputValue.c_str();//stringתCString
	
	//��������Ƶ����
	const string file = "C:\\Users\\Administrator\\Desktop\\workspace\\data\\";//��Ƶ�洢·��
	string a1 = argv[1] ;string a2 = argv[2] ;string a3 = argv[3] ;//�м����
	string video_name = a1 +"_" + a2 +"_" + sOutputValue + ".avi" ;//��Ƶ���ƣ��ع�ʱ��_���٣�����ֱ����_�������ֵ����������꣩
	string video_name_time = a1 +"_" + a2 +"_" + sOutputValue + "(time).avi" ;//ʱ�����Ƶ���ƣ��ع�ʱ��_���٣�����ֱ����_�������ֵ����������꣩
	string video_file = file + mfile + video_name;//��Ƶ·��+����
	string video_file_time = file + mfile + video_name_time;//ʱ�����Ƶ·��+����
	//����excel����
	string excel_name = a1 +"_" + a2 +"_" + sOutputValue + ".xlsx";//Ҫ���ɵ�excel�ļ���
	string excel_file = file + mfile + excel_name;//Ҫ���ɵ�excel�ļ�ȫ·��

	string init_excel_name = "init.xlsx";//����excel�ļ���
	string init_excel_file = file + init_excel_name;//����excel�ļ�ȫ·��

	//��excel����ģ���ʼ��
	cout<<"--excel����ģ���ʼ��"<<endl;
	bool bInit0 = excelProcessor.InitExcel();//��ʼ��excel������

	LPCWSTR L_init_excel_file = string2LPCWSTR(init_excel_file);//�ļ���stringת��ΪLPCWSTR
	bool bRet0 = excelProcessor.OpenExcelFile(L_init_excel_file);//������excel�ļ�
	CString strSheetName = excelProcessor.GetSheetName(1);//��ȡ��һ������������
	bool bLoad = excelProcessor.LoadSheet(strSheetName);//�������ƴ򿪸ù�����
	string textHead[5] = {"��¼ʱ��","����ֱ����mm��","���ɺ�ȣ�mm��","�����������źţ�mA��","�����źţ�mA��"};//���ɱ�ͷ
	for(int j=1;j < 6;j++)//�ڵ�һ�������ͷ
	{
		CString text;
		text = textHead[j - 1].c_str();
		excelProcessor.SetCellString(1,j,text);
	}
	


	//��modbus��ʼ��
	cout<<"--modbusЭ��ģ���ʼ��"<<endl;
	int isInit = initModbus();//��ʼ��
	if(isInit == -1)
	{
		cout<<"modbus��ʼ��ʧ��"<<endl;
		return 0;
	}

	//���������������С�����ñ��ģ�
	cout<<"--modbus�����������"<<endl;
	unsigned char setData[1024] = {0};//�����������
	int sendlen = 0;//����
	char A2H[3] = { 0 };
	for (int i = 0; i < 8; i += 2)//char����ת��Ϊunsigned char����
	{
		int num = 0;
		A2H[0] = argv[3][i];
		A2H[1] = argv[3][i+1];
		sscanf(A2H, "%02X", &num);
		setData[sendlen++] = num;
	}

	int isSet = setModbusOutput(setData);//���������0304�Ĵ�����ַ��0FFF������С
	if (isSet == -1)
	{
		cout<<"�����������ʧ��"<<endl;
		return 0;
	}

	//������������ı���
	//char dataForRead[9] = "03000001";
	unsigned char uDataForRead[1024] = { 0 };//�����������
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


	//��MV��ʼ��
	cout<<"--MV����ģ���ʼ��"<<endl;
	MVSTATUS_CODES mvInit;
	mvInit = MVInitLib();//��ʼ��
	if(mvInit != MVST_SUCCESS)
	{
		cout<<"������ʼ��ʧ�ܣ����⻷��"<<endl;
		return 0;
	}
	int nCams;//�������
	MVGetNumOfCameras(&nCams);//��ȡ���������nCams
	if( nCams == 0 )
	{
		cout<<"û���ҵ����,��ȷ�����Ӻ����IP����"<<endl;
		return 0;
	}
	MVSTATUS_CODES r = MVOpenCamByIndex(0,&m_hCam);//ͨ�������Ŵ�������Ϊ0�����
	if( m_hCam == NULL )//������Ϊ��
	{
		if( r == MVST_ACCESS_DENIED )//δ�ܴ����
		{
			cout<<"�޷��������������������������"<<endl;
			return 0;
		}
	}
	MVSetExposureTime (m_hCam,nExp_us);//�����ع�ʱ��
	MVSetFrameRate (m_hCam,fps);//���õ�ǰ֡��

	int w,h;//������������ÿ�͸���Ϣ/ͼ�����ظ�ʽ			
	MVGetWidth(m_hCam,&w);//��
	MVGetHeight(m_hCam,&h);//��
	MVGetPixelFormat(m_hCam,&m_PixelFormat);//��ȡͼ�����ظ�ʽ

	if( m_PixelFormat == PixelFormat_Mono8 )//����ǻҶ�ͼ
	{
		m_image.Create(w,h,8);
		m_cvSrc.create(h,w,CV_8UC1);//������ͨ��Mat
	}
	else
	{
		m_image.Create(w,h,24);
		m_cvSrc.create(h,w,CV_8UC3);//������ͨ����ɫMat
	}
	m_cvDst = m_cvSrc.clone();
	
	//��Ƶ�洢��ʼ��
	CvSize size =  m_cvDst.size();//��������Ƶ�ߴ�
	VideoWriter outputVideo(video_file,CV_FOURCC('D', 'I', 'V', 'X'),/*֡��*/3,size,1/*��ɫ1�ڰ�0*/);
	VideoWriter outputVideo_time(video_file_time,CV_FOURCC('D', 'I', 'V', 'X'),/*֡��*/3,size,1/*��ɫ1�ڰ�0*/);
	cout<<"��Ƶ�ֱ��ʣ�"<<w<<"x"<<h<<endl;
	cout<<"�ع⣺"<<nExp_us<<endl;
	cout<<"֡�ʣ�"<<fps<<endl;

	//��ʼ��ȡÿһ֡
	cout<<"--��ʼ����ϣ���ʼ�ɼ�;"<<endl;
	int excel_y = 2;//���������ʼ��
	while(true)
	{
		iNumFrame++;//֡���ͳ��,��2��ʼ
		getTheTime(iSysTime);//��ȡ��ǰʱ��
		//1д��ʱ��
		sprintf(cTime,"%dYear%dMon%dDay %d:%d:%d.%d",iSysTime[0],iSysTime[1],iSysTime[2],iSysTime[3],iSysTime[4],iSysTime[5],iSysTime[6]);
		char2CString(cTime, csTime);//��ʽת��
		if((iNumFrame % 1) == 0)//ÿ1֡�ڱ�񴢴�һ������
		{
			excelProcessor.SetCellString(excel_y,1,csTime);
			//2д������ֱ��
			excelProcessor.SetCellString(excel_y,2,csDiameterOfStream);
			//3���ɺ���ÿ�

			//4�������λ�ô�������źţ�ma��csOutputValue
			excelProcessor.SetCellString(excel_y,4,csOutputValue);
			//��ȡ�˿̵��������
			int isRead = readModbusInput(uDataForRead);//��ȡ�˿̵��������ֵ�������źţ�
			
			double dInputValue = ((double)isRead / 4096.0  * 16.0) + 4.0;
			//5д�뵱ǰʱ�̲����ź�
			stringstream tmp2;
			tmp2 << dInputValue;//����ת�м����
			tmp2 >> sDataInput;//�м����תstring
			 cout<<"int"<<isRead<<"double"<<dInputValue<<"string"<<sDataInput<<endl;
			csDataInput = sDataInput.c_str();//stringתCString
			excelProcessor.SetCellString(excel_y,5,csDataInput);
			excel_y++;
		}

		MVStartGrab(m_hCam,StreamCB,paraForSet);//�ɼ�
		outputVideo << m_cvSrc;//Դ��Ƶ����

		string text(cTime);//ʱ���
		putText(m_cvSrc,text,Point(100,100),FONT_HERSHEY_PLAIN,4.0,Scalar(255, 255, 255),3.0);//����ʱ���
		outputVideo_time << m_cvSrc;//ʱ�����Ƶ����

		m_cvDst = m_cvSrc.clone();	//�����Ļ�����Mat
		resize(m_cvDst,m_cvDst,Size(),0.5,0.5,1);
		imshow("��Ƶ",m_cvDst);
		

		if(waitKey(200) == 27 )
		{
			MVStopGrab(m_hCam);//ֹͣץȡ
			closeModbus();
			if( m_hCam != NULL )
			{
				MVCloseCam(m_hCam);//�ر�����ͷ
			}
			MVTerminateLib();//�ͷ���Դ
			break;
		}
	}
	//�洢��Ӧ�ı��
	LPCWSTR L_excel_file = string2LPCWSTR(excel_file);//��ת��excel��stringתΪLPCWSTR
	excelProcessor.SaveasXSLFile(L_excel_file);
	excelProcessor.CloseExcelFile(1);
	excelProcessor.ReleaseExcel(); 

	cout<<"���!�洢֡����"<<iNumFrame + 1<<endl;
	cout<<"----------------------------------------------"<<endl;
	return 0;
}



int initModbus()//��ʼ��modbus����������ͨѶ�����ɹ��򷵻�1��ʧ���򷵻�-1
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
int setModbusOutput(unsigned char *data /*"03040FFF"*/)//���õ������(0-4095)���Ӷ��������������,�ɹ�����1��ʧ�ܷ���-1
{
	const int func_code_6_AO = 0x06;//д��������� ������
	int recieve_station_addr = station_addr;//һ������صĵ�ַ
	int recieve_func_code_6_AO = func_code_6_AO;//һ������صĹ�����
	unsigned char recieve_data[1024] = {0};//�����ܵ�����
	//ͨ��
	int datalen2 = RequestData(comm_type, conn_num, station_addr, func_code_6_AO,data, 4); // ��������
	if(datalen2 == -1)
	{
		return -1;//����ʧ��
	}

	//����
	int bytecount2 = ResponseData(comm_type, conn_num, &recieve_station_addr, &recieve_func_code_6_AO, recieve_data);    // ���ݻظ�
	//��֤
	if(bytecount2 > 0)
	{
		if(recieve_func_code_6_AO !=  func_code_6_AO)//����ظ��Ĺ����벻ͬ ��Ϊ�����˴���
		{
			
			return -1;//дֵʧ��
		}
		else
		{
			return 1;//дֵ�ɹ�
		}
	}
	else
	{
		return -1;//δ�յ�����
	}
}

int readModbusInput(unsigned char *data/*"03000001"*/)//��ȡ��������ֵ��0-4095������������Ӧintֵ,ʧ�ܷ���-1
{
	const int func_code_3_AI = 0x03;//��ȡ������� ������
	int recieve_station_addr = station_addr;//һ������صĵ�ַ
	int recieve_func_code_3_AI = func_code_3_AI;//һ������صĹ�����
	unsigned char recieve_data[1024] = {0};//�����ܵ�����

	int datalen1 = RequestData(comm_type, conn_num, station_addr, func_code_3_AI,data , 4); // ��������
	if(datalen1 == -1)
	{
		return -1;//����ʧ��
	}
	int bytecount1 = ResponseData(comm_type, conn_num, &recieve_station_addr, &recieve_func_code_3_AI, recieve_data);    // ���ݻظ�
	if(bytecount1 > 0)
	{
		if(recieve_func_code_3_AI !=  func_code_3_AI)//����ظ��Ĺ����벻ͬ ��Ϊ�����˴���
		{
			return -1;//�յ����ݣ�����ȡʧ��
		}
		else//����ظ��Ĺ�����û�б仯 ��˵������
		{//���յ�������recieve_data(unsigned char)ת��Ϊint��ʮ��������
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
		return -1;//δ�յ�����
	}

}

void closeModbus()
{
	CloseSerial(conn_num); // �ر�ͨѶ 
}

void getTheTime(int *iSysTime)//��ȡ��ǰ��ϵͳʱ��
{
	SYSTEMTIME sys;
	GetLocalTime( &sys );
	iSysTime[0] = sys.wYear;//��
	iSysTime[1] = sys.wMonth;//��
	iSysTime[2] = sys.wDay;//��
	iSysTime[3] = sys.wHour;//ʱ
	iSysTime[4] = sys.wMinute;//��
	iSysTime[5] = sys.wSecond;//��
	iSysTime[6] = sys.wMilliseconds;//����
}

LPCWSTR string2LPCWSTR(string orig)//��stringת��ΪLPCWSTR
{
	size_t origsize = orig.length() + 1;
    const size_t newsize = 100;
    size_t convertedChars = 0;
	wchar_t *wcstring = (wchar_t *)malloc(sizeof(wchar_t)*(orig.length()-1));
	mbstowcs_s(&convertedChars, wcstring, origsize, orig.c_str(), _TRUNCATE);
	return wcstring;
}

int __stdcall StreamCB(MV_IMAGE_INFO *pInfo,long nUserVal)//MV�������,��MV��ʽ����ת��ΪMAT
{
	//��ȡͼ��ĳߴ�
	int w = m_image.GetWidth();
	int h = m_image.GetHeight();

	if( m_PixelFormat == PixelFormat_Mono8 )//�Ҷ�ͼ����
	{
		ASSERT( m_image.GetPitch() == m_image.GetWidth() );
		memcpy(m_cvSrc.data,pInfo->pImageBuffer,w*h);
		
	}
	else//��ͼ
	{
		MVBayerToBGR(m_hCam,pInfo->pImageBuffer,m_cvSrc.data,w*3,w,h,m_PixelFormat);
	}
	return 0;
}

int char2CString(char *in, CString &out)//�ַ�����ת��ΪCString
{
	if(NULL == in)
	{
		cout<<"int char2CString(char *in, CString &out) ��������Ϊ��"<<endl;
		return -1;
	}
	string argv2(in);//��תΪstring
	out = argv2.c_str();//תΪCString
	return 0;
}