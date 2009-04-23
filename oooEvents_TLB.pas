unit oooEvents_TLB;

// ************************************************************************ //
// WARNING                                                                    
// -------                                                                    
// The types declared in this file were generated from data read from a       
// Type Library. If this type library is explicitly or indirectly (via        
// another type library referring to this type library) re-imported, or the   
// 'Refresh' command of the Type Library Editor activated while editing the   
// Type Library, the contents of this file will be regenerated and all        
// manual modifications will be lost.                                         
// ************************************************************************ //

// PASTLWTR : 1.2
// File generated on 22/04/2009 11:51:45 from Type Library described below.

// ************************************************************************  //
// Type Lib: C:\brant\sistemas\sis-gds\CT\Programa\sct\delphi\oooEvents\oooEvents.tlb (1)
// LIBID: {313E866F-0192-4F25-8231-71EEB5D4C9B0}
// LCID: 0
// Helpfile: 
// HelpString: oooEvents Library
// DepndLst: 
//   (1) v2.0 stdole, (C:\WINDOWS\system32\STDOLE2.TLB)
//   (2) v4.0 StdVCL, (C:\WINDOWS\system32\stdvcl40.dll)
// ************************************************************************ //
{$TYPEDADDRESS OFF} // Unit must be compiled without type-checked pointers. 
{$WARN SYMBOL_PLATFORM OFF}
{$WRITEABLECONST ON}
{$VARPROPSETTER ON}
interface

uses Windows, ActiveX, Classes, Graphics, StdVCL, Variants;
  

// *********************************************************************//
// GUIDS declared in the TypeLibrary. Following prefixes are used:        
//   Type Libraries     : LIBID_xxxx                                      
//   CoClasses          : CLASS_xxxx                                      
//   DISPInterfaces     : DIID_xxxx                                       
//   Non-DISP interfaces: IID_xxxx                                        
// *********************************************************************//
const
  // TypeLibrary Major and minor versions
  oooEventsMajorVersion = 1;
  oooEventsMinorVersion = 0;

  LIBID_oooEvents: TGUID = '{313E866F-0192-4F25-8231-71EEB5D4C9B0}';

  IID_IListener: TGUID = '{9810785B-C98F-4ECD-A38C-DABF8248B9BA}';
  DIID_IListenerEvents: TGUID = '{7FF12C06-6303-4ADE-ABAD-41208AE4F24D}';
  CLASS_Listener: TGUID = '{F614DF6A-7D45-4A9A-B6B2-DCD0614F0895}';
type

// *********************************************************************//
// Forward declaration of types defined in TypeLibrary                    
// *********************************************************************//
  IListener = interface;
  IListenerDisp = dispinterface;
  IListenerEvents = dispinterface;

// *********************************************************************//
// Declaration of CoClasses defined in Type Library                       
// (NOTE: Here we map each CoClass to its Default Interface)              
// *********************************************************************//
  Listener = IListener;


// *********************************************************************//
// Interface: IListener
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {9810785B-C98F-4ECD-A38C-DABF8248B9BA}
// *********************************************************************//
  IListener = interface(IDispatch)
    ['{9810785B-C98F-4ECD-A38C-DABF8248B9BA}']
    procedure disposing; safecall;
    procedure notifyEvent(const event: IDispatch); safecall;
  end;

// *********************************************************************//
// DispIntf:  IListenerDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {9810785B-C98F-4ECD-A38C-DABF8248B9BA}
// *********************************************************************//
  IListenerDisp = dispinterface
    ['{9810785B-C98F-4ECD-A38C-DABF8248B9BA}']
    procedure disposing; dispid 201;
    procedure notifyEvent(const event: IDispatch); dispid 202;
  end;

// *********************************************************************//
// DispIntf:  IListenerEvents
// Flags:     (0)
// GUID:      {7FF12C06-6303-4ADE-ABAD-41208AE4F24D}
// *********************************************************************//
  IListenerEvents = dispinterface
    ['{7FF12C06-6303-4ADE-ABAD-41208AE4F24D}']
  end;

// *********************************************************************//
// The Class CoListener provides a Create and CreateRemote method to          
// create instances of the default interface IListener exposed by              
// the CoClass Listener. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoListener = class
    class function Create: IListener;
    class function CreateRemote(const MachineName: string): IListener;
  end;

implementation

uses ComObj;

class function CoListener.Create: IListener;
begin
  Result := CreateComObject(CLASS_Listener) as IListener;
end;

class function CoListener.CreateRemote(const MachineName: string): IListener;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_Listener) as IListener;
end;

end.
