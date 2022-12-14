#include <windows.h>
#include <OLEAUTO.H>
#include <stdio.h>
 
//uses elroys persistant debug print window to receive dbg messages
//  https://github.com/dzzie/libs/tree/master/_elroy/PersistentDebugPrint_mod1


//COM error codes: https://learn.microsoft.com/en-us/windows/win32/com/com-error-codes-1
//                 https://learn.microsoft.com/en-us/windows/win32/com/com-error-codes-2

//iTypeInfo in IDE is in vba.dll
//             ActiveX dll is in oleaut32.dll
//             compiled internal class is a partial implementation in msvbvm60.dll


bool Warned=false;
HWND hServer=0;

void FindVBWindow(){
	char *vbIDEClassName = "ThunderFormDC" ;
	char *vbEXEClassName = "ThunderRT6FormDC" ;
	char *vbEXEClassName2 = "ThunderRT6Form" ;
	char *vbWindowCaption = "Persistent Debug Print Window" ;

	hServer = FindWindowA( vbIDEClassName, vbWindowCaption );
	if(hServer==0) hServer = FindWindowA( vbEXEClassName, vbWindowCaption );
	if(hServer==0) hServer = FindWindowA( vbEXEClassName2, vbWindowCaption );	
} 

int msg(char *Buffer){
  
  if(!IsWindow(hServer)) hServer=0;
  if(hServer==0) FindVBWindow();
  if(!IsWindow(hServer)) return -1;

  COPYDATASTRUCT cpStructData;
  memset(&cpStructData,0, sizeof(struct tagCOPYDATASTRUCT )) ;

  cpStructData.dwData = 3;
  cpStructData.cbData = strlen(Buffer) ;
  cpStructData.lpData = (void*)Buffer;
  int ret = SendMessage(hServer, WM_COPYDATA, 0,(LPARAM)&cpStructData);
  return ret; //log ui can send us a response msg to trigger special reaction in ret

} 

int msg(const char *format, ...)
{
	DWORD dwErr = GetLastError();
	int ret=0;
	if(format){
		char buf[1024]; 
		va_list args; 
		va_start(args,format); 
		try{
 			 _vsnprintf(buf,1024,format,args);
			 ret = msg(buf);
		}
		catch(...){}
	}

	SetLastError(dwErr);
	return ret;
}

int* pPlus(void *p, int increment){
  _asm{
		mov eax, p
		add eax, increment
   }
}

int* pPlus(int p, int increment){
  _asm{
		mov eax, p
		add eax, increment
   }
}

struct funcTyp{
	char stackSize;
	char isFunc;
	short vOff;
	int reserved;
	int nul1;
	int memberID;
};

/* 
	Vtable - 4 is pointer to the target Class ObjInfo struct
	ObjInfo->privObj->funcTypeInfoAry is an array of *funcTyp with objInfo->parentObj->methodCount elements...

	.text:00401C98 CTest_Priv_FuncTypInf   dd offset CTest_PubFuncTypDesc_test
	.text:00401C9C                         dd offset CTest_PubFuncTypDesc_a
	.text:00401CA0                         dd 0

	if funcTyp->memberID = IDisp->GetIDsOfNames(methodName) then return entry at [vTable + funcTyp->vOff]
*/
int manualParse(IDispatch *IDisp, int membID, int *funcAddress, int *funcOffset){
	
	void *objInfo;
	void *parentObj;
	void *privObj;
	int methodCount=0;
	int funcTypeInfoAry = 0;
	int voff=0;
	int vTable;

	_asm{
		mov eax, IDisp
		mov ebx, [eax]    //vtable
		mov vTable, ebx;

		mov ecx, [ebx-4]  //vtable - 4 = class *ObjInfo
		mov objInfo, ecx  //todo check if null / confirm vb6 class 

		mov edx, [ecx+0x18] //objInfo->parentObj
		mov parentObj, edx

        mov eax, [edx+0x1C]   //objInfo->parentObj->methodCount
		mov methodCount, eax

		mov edx, [ecx+0x0C];  //objInfo->privObj
		mov privObj, edx

		mov eax, [edx+0x18]      //objInfo->privObj->funcTypeInfoAry
		mov funcTypeInfoAry, eax
	}

	msg("objInfo = %x", (int)objInfo);
	msg("parentObj = %x methodCnt = %d", (int)parentObj, methodCount);
	msg("privObj = %x", (int)privObj);
	msg("funcTypeInfoAry = %x", funcTypeInfoAry);

	funcTyp **fta = (funcTyp**)funcTypeInfoAry; //array of funcTyp pointers
	
	for(int i=0; i < methodCount; i++){

		msg("funcTyp[%d] = 0x%x", i, (int)fta[i]);

		if((int)fta[i] != 0){

			voff = fta[i]->vOff;
			if(voff % 4 != 0) voff -= 1; //they set bit 1 as a flag (Invoke kind?)

			msg("memberID 0x%x  vTableOffset: 0x%x", fta[i]->memberID, voff);

			if(fta[i]->memberID == membID){
			    int funcAddr = *pPlus(vTable,voff);
				msg("class function implementation: 0x%x", funcAddr);
				*funcAddress = funcAddr;
				*funcOffset = vTable + voff;
				return 1;
			}

		}
	}

	return 0;
}

int __stdcall ComFuncAddr(IDispatch* IDisp, OLECHAR* sMethodName, int *funcAddr, int *funcOffset){
	
	int tmp,i;
	HRESULT hr;	
	DISPID  dispid;  
	ITypeInfo  *IType;
	TYPEATTR *ta;
	FUNCDESC *fd; //https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-oaut/d3349d25-e11d-4095-ba86-de3fda178c4e

	 msg("<cls>");
	 *funcAddr = 0;
	 *funcOffset = 0;

	 if((int)IDisp==0) return 0;

	 int vTable = *(int*)IDisp;
	 msg("vtable = %x", vTable);

	 hr = IDisp->GetTypeInfo(0,NULL,&IType);
	 if ( hr != S_OK )
	 {
	   msg("GetTypeInfo failed");
	   return 0;
	 }

	 msg("IDisp=%x\n IType = %x",(int)IDisp, (int)IType); //text:66018B98 ; const CEcTypeInfo::`vftable'

	 // Get the Dispatch ID for the method name
	 hr=IDisp->GetIDsOfNames(IID_NULL,&sMethodName,1,LOCALE_USER_DEFAULT,&dispid);
	 if( FAILED(hr) ){
	    msg("GetIDsOfNames failed");
		return 0;
	 }

	 msg("MemberID = %x", dispid);

	 hr= IType->GetTypeAttr(&ta); 
	 if( FAILED(hr) ){
		 if(hr == 0x80004001){
			msg("IType->GetTypeAttr Not implemented trying manual parse..."); 
			return manualParse(IDisp, dispid, funcAddr, funcOffset);
		 }else{
			msg("IType->GetTypeAttr failed %x", hr); 
			return 0;
		 }
	 }

	 msg("Full ITypeInfo found (IDE/ActiveX), scanning for member id");
	 
	 for(i=0; i < ta->cFuncs; i++){
		hr = IType->GetFuncDesc(i, &fd);
		if( !FAILED(hr) ){
			if(fd->memid == dispid && fd->invkind == INVOKE_FUNC){
				int tmp = *pPlus(vTable, fd->oVft);
				*funcAddr = tmp;
				*funcOffset = vTable + fd->oVft;
				return 1;
			}
		}
	 }

	 return 0;

}
 
//to verbose
		/*
		
		tmp = fd->oVft;
		_asm{ //some things are cleaner in asm...
			mov eax, IDisp   ; pointer to vtable
			mov ebx, [eax]   ; actual vtable
			add ebx, tmp     ; vtable offset
			mov ecx, [ebx]   ; func address at offset
			mov tmp, ecx
		}
		//int t2 = *pPlus( (void*)*(int*)IDisp,fd->oVft);msg("t2 = %x", t2);
		-------------------
		
		ft = (funcTyp*)*(int*)(funcTypeInfoAry+(i*4));
		//msg("funcTyp[%d] = 0x%x", i, (int)ft);

		//just preference...easier to control/debug
		_asm{
			mov edi, funcTypeInfoAry
			mov eax, i
			mov ecx, 4
			mul ecx
			add edi, eax
			mov esi, [edi]
			mov ft, esi
		}
        msg("funcTyp[%d] = %x", i, (int)ft);
		-------------------
		
		_asm{
			mov eax, vTable
			mov ebx, voff
			add eax, ebx
			mov ecx, [eax]
			mov funcAddr, ecx
		}
		
		
		*/