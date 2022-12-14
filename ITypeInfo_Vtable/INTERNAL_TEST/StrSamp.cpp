#include <windows.h>
#include <OLEAUTO.H> //only needed for SAFEARRAY structure
#include <stdio.h>
 
//COM error codes: https://learn.microsoft.com/en-us/windows/win32/com/com-error-codes-1
//                 https://learn.microsoft.com/en-us/windows/win32/com/com-error-codes-2


int __stdcall ComFuncAddr(IDispatch* IDisp, OLECHAR* sMethodName){
	
	int tmp,i;
	HRESULT hr;	
	DISPID  dispid;  
	//IDispatch* IDisp;
	ITypeInfo  *IType;
	//ITypeInfo2 *IType2;
	TYPEATTR *ta;
	FUNCDESC *fd; //https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-oaut/d3349d25-e11d-4095-ba86-de3fda178c4e
    char buf[2000];

	 if((int)IDisp==0) return 0;

	 hr = IDisp->GetTypeInfo(0,NULL,&IType);
	 if ( hr != S_OK )
	 {
	   MessageBox(0,"GetTypeInfo failed","",0);
	   return 0;
	 }

	 sprintf(&buf[0], "IDisp=%x\n IType = %x",(int)IDisp, (int)IType); //text:66018B98 ; const CEcTypeInfo::`vftable'
	 MessageBox(0,buf,"",0);

	 /*hr = IType->QueryInterface(IID_ITypeInfo2, (void**)&IType2);
	 if ( hr != S_OK )
	 {
	   MessageBox(0,"IID_ITypeInfo2 failed","",0);
	   return 0;
	 }*/

	 // Get the Dispatch ID for the method name
	 hr=IDisp->GetIDsOfNames(IID_NULL,&sMethodName,1,LOCALE_USER_DEFAULT,&dispid);
	 if( FAILED(hr) ){
	    MessageBox(0,"GetIDsOfNames failed","",0);
		return 0;
	 }

	 hr= IType->GetTypeAttr(&ta); //0x80004001: Not implemented: 
	 if( FAILED(hr) ){
		 if(hr == 0x80004001){
			MessageBox(0,"IType->GetTypeAttr failed 0x80004001: Not implemented","",0);
		 }else{
			sprintf(&buf[0], "IType->GetTypeAttr failed %x", hr); 
			MessageBox(0,buf,"",0);
		 }
		//return 0;
		 ta = new TYPEATTR;
		 ta->cFuncs = 100;
	 }
	 


	 for(i=0; i < ta->cFuncs;i++){
		hr = IType->GetFuncDesc(i, &fd);
		if( !FAILED(hr) ){
			if(fd->memid == dispid && fd->invkind == INVOKE_FUNC){
				tmp = fd->oVft;
				_asm{ //some things are cleaner in asm...todo: add null pointer err checking...
					mov eax, IDisp ; pointer to vtable
					mov ebx, [eax] ; actual vtable
					add ebx, tmp     ; vtable offset
					mov ecx, [ebx]   ; func address at offset
					mov tmp, ecx
				}
				return tmp;
			}
		}
	 }

	 return 0;

}

/*

00403360 >6605F210  MSVBVM60.BASIC_CLASS_QueryInterface  
   .text:6605F210 ; __int32 __stdcall CDeskFrame::QueryInterface(CDeskFrame *this, const struct _GUID *, void **)

00403364  66001B58  MSVBVM60.Zombie_AddRef
00403368  66001B68  MSVBVM60.Zombie_Release
0040336C  660E26F7  MSVBVM60.BASIC_DISPINTERFACE_GetTICount
00403370  66041E5E  MSVBVM60.BASIC_DISPINTERFACE_GetTypeInfo
00403374  6600BDE7  MSVBVM60.BASIC_CLASS_GetIDsOfNames
00403378  6600B841  MSVBVM60.BASIC_CLASS_Invoke  
   .text:6600B841 ; int __stdcall BASIC_CLASS_Invoke(BASIC_DISPINTERFACE *, int, int, int, int, int, int, int, int)

0040337C  004014F4  vb.___vba@03E5D300
00403380  00401503  vb.00401503
00403384  00401512  vb.00401512
00403388  00401529  vb.00401529
0040338C  00401536  vb.00401536
00403390  00401543  vb.00401543

iTypeInfo in IDE is in vba.dll
-------------------------------------------
0FAA0780  0FAFDA64  VBA6.0FAFDA64
0FAA0784  0FAA074B  VBA6.0FAA074B
0FAA0788  0FAA0660  VBA6.0FAA0660
0FAA078C  0FAFCF8C  VBA6.0FAFCF8C   GetTypeAttr
0FAA0790  0FAFDEA8  VBA6.0FAFDEA8
0FAA0794  0FAFE7CC  VBA6.0FAFE7CC
0FAA0798  0FB6C397  VBA6.0FB6C397
0FAA079C  0FAFDDE7  VBA6.0FAFDDE7
0FAA07A0  0FB01E0C  VBA6.0FB01E0C
0FAA07A4  0FB6C40F  VBA6.0FB6C40F
0FAA07A8  0FB6C457  VBA6.0FB6C457
0FAA07AC  0FB6C48B  VBA6.0FB6C48B
0FAA07B0  0FAA0904  VBA6.0FAA0904
0FAA07B4  0FB6C4CB  VBA6.0FB6C4CB
0FAA07B8  0FAFEA76  VBA6.0FAFEA76
0FAA07BC  0FB6C4D3  VBA6.0FB6C4D3
0FAA07C0  0FB6C502  VBA6.0FB6C502
0FAA07C4  0FB6C58E  VBA6.0FB6C58E
0FAA07C8  0FB6C5B9  VBA6.0FB6C5B9
0FAA07CC  0FAFD23E  VBA6.0FAFD23E
0FAA07D0  0FAFEA45  VBA6.0FAFEA45
0FAA07D4  0FB6C747  VBA6.0FB6C747



.text:66018B98                 dd offset ?QueryInterface@CEcTypeInfo@@UAGJABU_GUID@@PAPAX@Z
.text:66018B9C                 dd offset ?AddRef@CEcTypeInfo@@UAGKXZ ; CEcTypeInfo::AddRef(void)
.text:66018BA0                 dd offset ?Release@CEcTypeInfo@@UAGKXZ ; CEcTypeInfo::Release(void)
.text:66018BA4                 dd offset ?GetTypeAttr@CEcTypeInfo@@UAGJPAPAUtagTYPEATTR@@@Z ; CEcTypeInfo::GetTypeAttr(tagTYPEATTR * *)
.text:66018BA8                 dd offset ?GetTypeComp@CEcTypeInfo@@UAGJPAPAUITypeComp@@@Z ; CEcTypeInfo::GetTypeComp(ITypeComp * *)
.text:66018BAC                 dd offset ?GetFuncDesc@CEcTypeInfo@@UAGJIPAPAUtagFUNCDESC@@@Z ; CEcTypeInfo::GetFuncDesc(uint,tagFUNCDESC * *)
.text:66018BB0                 dd offset ?GetVarDesc@CEcTypeInfo@@UAGJIPAPAUtagVARDESC@@@Z ; CEcTypeInfo::GetVarDesc(uint,tagVARDESC * *)
.text:66018BB4                 dd offset ?GetNames@CEcTypeInfo@@UAGJJPAPAGIPAI@Z ; CEcTypeInfo::GetNames(long,ushort * *,uint,uint *)
.text:66018BB8                 dd offset ?SetMoniker@DESKOLE@@UAGJKPAUIMoniker@@@Z ; DESKOLE::SetMoniker(ulong,IMoniker *)
.text:66018BBC                 dd offset ?SetMoniker@DESKOLE@@UAGJKPAUIMoniker@@@Z ; DESKOLE::SetMoniker(ulong,IMoniker *)
.text:66018BC0                 dd offset ?GetIDsOfNames@CEcTypeInfo@@UAGJPAPAGIPAJ@Z ; CEcTypeInfo::GetIDsOfNames(ushort * *,uint,long *)
.text:66018BC4                 dd offset ?Invoke@CEcTypeInfo@@UAGJPAXJGPAUtagDISPPARAMS@@PAUtagVARIANT@@PAUtagEXCEPINFO@@PAI@Z ; CEcTypeInfo::Invoke(void *,long,ushort,tagDISPPARAMS *,tagVARIANT *,tagEXCEPINFO *,uint *)
.text:66018BC8                 dd offset ?GetDocumentation@CEcTypeInfo@@UAGJJPAPAG0PAK0@Z ; CEcTypeInfo::GetDocumentation(long,ushort * *,ushort * *,ulong *,ushort * *)
.text:66018BCC                 dd offset ?GetIDsOfNames@CEventSink@COcx@@UAGJABU_GUID@@PAPAGIKPAJ@Z ; COcx::CEventSink::GetIDsOfNames(_GUID const &,ushort * *,uint,ulong,long *)
.text:66018BD0                 dd offset ?GetRefTypeInfo@CEcTypeInfo@@UAGJKPAPAUITypeInfo@@@Z ; CEcTypeInfo::GetRefTypeInfo(ulong,ITypeInfo * *)
.text:66018BD4                 dd offset ?GetIDsOfNames@StubTypeInfo@@UAGJPAPAGIPAJ@Z ; StubTypeInfo::GetIDsOfNames(ushort * *,uint,long *)
.text:66018BD8                 dd offset ?GetIDsOfNames@StubTypeInfo@@UAGJPAPAGIPAJ@Z ; StubTypeInfo::GetIDsOfNames(ushort * *,uint,long *)
.text:66018BDC                 dd offset ?SetMoniker@DESKOLE@@UAGJKPAUIMoniker@@@Z ; DESKOLE::SetMoniker(ulong,IMoniker *)
.text:66018BE0                 dd offset ?SetMoniker@DESKOLE@@UAGJKPAUIMoniker@@@Z ; DESKOLE::SetMoniker(ulong,IMoniker *)
.text:66018BE4                 dd offset ?__trapFreeMarked@@YGXPAX0@Z ; __trapFreeMarked(void *,void *)
.text:66018BE8                 dd offset ?__trapFreeMarked@@YGXPAX0@Z ; __trapFreeMarked(void *,void *)
.text:66018BEC                 dd offset ?__trapFreeMarked@@YGXPAX0@Z ; __trapFreeMarked(void *,void *)
.text:66018BF0                 dd offset ??_GCEcTypeInfo@@UAEPAXI@Z ; CEcTypeInfo::`scalar deleting destructor'(uint)
.text:66018BF4                 align 8
.text:66018BF8                 public ??_7CEcTypeComp@@6B@
*/