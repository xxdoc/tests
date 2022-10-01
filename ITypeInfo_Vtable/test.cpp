#include <windows.h>
#include <stdio.h>
#include <conio.h>
#include <ole2.h>

IDispatch        *IDisp;
ITypeInfo        *IType;
ITypeInfo2       *IType2;

int CreateObj(){

    //Create an instance of our VB COM object, and execute
	//one of its methods so that it will load up and show a UI
	//for us, then it uses our other exports to access olly plugin API
	//methods

	CLSID      clsid;
	HRESULT	   hr;
    LPOLESTR   p = OLESTR("vbSample.CTest");

    hr = CoInitialize(NULL);

	 hr = CLSIDFromProgID( p , &clsid);
	 if( hr != S_OK  ){
		 MessageBox(0,"Failed to get Clsid from string\n","",0);
		 return 0;
	 }

	 // create an instance and get IDispatch pointer
	 hr =  CoCreateInstance( clsid,
							 NULL,
							 CLSCTX_INPROC_SERVER,
							 IID_IDispatch  ,
							 (void**) &IDisp
						   );

	 if ( hr != S_OK )
	 {
	   MessageBox(0,"CoCreate failed","",0);
	   return 0;
	 }

	 hr = IDisp->GetTypeInfo(0,NULL,&IType);
	 if ( hr != S_OK )
	 {
	   MessageBox(0,"GetTypeInfo failed","",0);
	   return 0;
	 }
		
	 hr = IType->QueryInterface(IID_ITypeInfo2, (void**)&IType2);
	 if ( hr != S_OK )
	 {
	   MessageBox(0,"IID_ITypeInfo2 failed","",0);
	   return 0;
	 }

	 return 1;
}

int test(int arg){
	
	 HRESULT	   hr;

	 OLECHAR *sMethodName = OLESTR("test");
	 DISPID  dispid; // long integer containing the dispatch ID
     char buf[2000];

	 //sprintf(&buf[0], "ITypeInfo implementation: %x", (int)IType);
	 //MessageBoxA(0,buf,"",0); //OLEAUT32.75C6463C

	 // Get the Dispatch ID for the method name
	 hr=IDisp->GetIDsOfNames(IID_NULL,&sMethodName,1,LOCALE_USER_DEFAULT,&dispid);
	 if( FAILED(hr) ){
	    MessageBox(0,"GetIDS failed","",0);
		return 0;
	 }

	 DISPPARAMS dispparams;
	 VARIANTARG vararg[1]; //function takes one argument
	 VARIANT    retVal;

	 VariantInit(&vararg[0]);

	 vararg[0].vt = VT_I4 ;
	 vararg[0].intVal = arg;

	 dispparams.rgvarg = &vararg[0];
	 dispparams.cArgs = 1;  // num of args function takes
	 dispparams.cNamedArgs = 0;

	 PVOID lpfn;
	 hr = IType->AddressOfMember(dispid, INVOKE_FUNC, &lpfn);
	 //0x800288BD = TYPE_E_BADMODULEKIND 
	 printf("%X = IType->AddressOfMember(%x) = %x\n", hr, dispid, (int)lpfn);

	 TYPEATTR *ta;
	 hr= IType->GetTypeAttr(&ta);
	 printf("%X = IType2->GetTypeAttr.cfuncs = %x\n", hr, ta->cFuncs);
	 
	 int vtable=0, tmp=0;
	 FUNCDESC *fd; //https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-oaut/d3349d25-e11d-4095-ba86-de3fda178c4e
	 for(int i=0; i < ta->cFuncs;i++){
		hr = IType->GetFuncDesc(i, &fd);
		if( !FAILED(hr) ){
			printf("\t i=%d memid=%x  vft=%x\n", i, fd->memid, fd->oVft);
			if(fd->memid == dispid && fd->invkind == INVOKE_FUNC){
				printf("Found it IDisp=%x + %x\n", (int)IDisp, fd->oVft);
				tmp = fd->oVft;
				_asm{ //some things are cleaner in asm...
					mov eax, IDisp ; pointer to vtable
					mov ebx, [eax] ; actual vtable
					add ebx, tmp     ; vtable offset
					mov ecx, [ebx]   ; func address at offset
					mov tmp, ecx
				}
				printf("Final address: %x\n", tmp);
				getch();
				break;
			}
		}
	 }

	 //printf("%X = IType2->AddressOfMember(%x) = %x\n", hr, dispid, (int)lpfn);

	 hr = IType->Invoke(IDisp, dispid,DISPATCH_METHOD, &dispparams, &retVal, NULL, NULL);
     printf("%X\n", hr);

	 // and invoke the method
	 //DebugBreak();
	 //hr=IDisp->Invoke( dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dispparams, &retVal, NULL, NULL);

	 return 0;
}



void main(void)
{
	if(CreateObj() < 1 ) return;

	test(21);
	getch();

}


/*
004FD974  |. 8B41 18        MOV EAX,DWORD PTR DS:[ECX+18]            ;  MSVBVM60.BASIC_CLASS_Invoke

__int32 __stdcall EpiInvokeMethod(
	void *a1, 
	struct epiModule *a2, 
	__int32 a3, 
	int a4, 
	int a5, 
	int a6, 
	struct tagDISPPARAMS *a7, 
	VARIANTARG *pvarg, 
	IErrorInfo *pperrinfo, 
	unsigned int *a10
)

6600BB7F   FF15 24071166    CALL DWORD PTR DS:[66110724]             ; OLEAUT32.DispCallFunc

75C98A5D   FF15 5CA4CE75    CALL DWORD PTR DS:[75CEA45C]             ; OLEAUT32.75C85580  returns function offset in ecx
75C98A63   8BCB             MOV ECX,EBX                              ; vbSample.11001398
75C98A65   64:800D CA0F0000>OR BYTE PTR FS:[FCA],1
75C98A6D   FFD1             CALL ECX

11001398   E9 E3040000      JMP vbSample.11001880

https://learn.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-dispcallfunc

HRESULT DispCallFunc(
  void       *pvInstance,
  ULONG_PTR  oVft,
  CALLCONV   cc,
  VARTYPE    vtReturn,
  UINT       cActuals,
  VARTYPE    *prgvt,
  VARIANTARG **prgpvarg,
  VARIANT    *pvargResult
);

https://github.com/reactos/reactos/blob/master/dll/win32/oleaut32/typelib.c

DispCallFunc: https://github.com/reactos/reactos/blob/3fa57b8ff7fcee47b8e2ed869aecaf4515603f3f/dll/win32/oleaut32/typelib.c#L6446
*/