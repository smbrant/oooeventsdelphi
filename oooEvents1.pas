{
   Basic ooo listener for Delphi

   This program was created with the following script:
   - Create a normal delphi application (code saved as oooEvents.dpr and oooEvents1.pas)
   - Add an ActiveX automation object to the project (oooEvents_TLB.pas e oooEvents2.pas)
       (I used the CoClass Name 'TListener'
       and checked the option 'Generate Event support code'.
       You can do this using the menu File/New/Other/ActiveX/Automation Object)
   - Add the methods 'disposing' and 'notifyEvent' to the default interface IListener
   - Insert code to treat events in the notifyEvent method
   - Register the event listener to ooo using the method addListener (in
     the ooo object of interest, normally the document, see code associated to the button)
   - Remember to register the automation object in Windows with
       oooListener /regserver

   Use:
   - Open a document in ooo
   - Run this program and click in the button 'Connect to current document'
   - Make changes to the ooo document and see the events being captured.

   Author: smbrant, apr/2009, compiled with Delphi7.

}
unit oooEvents1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComObj;

type
  TForm1 = class(TForm)
    Button1: TButton;
    Memo1: TMemo;
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

uses oooEvents2, oooEvents_TLB;

{$R *.dfm}

procedure TForm1.Button1Click(Sender: TObject);
var
  serviceManager, desktop, document, controller, frame, window, visible, oListener: OleVariant;
  myListener: Listener;
begin
    serviceManager:= CreateOleObject('com.sun.star.ServiceManager');
    desktop:= serviceManager.createInstance('com.sun.star.frame.Desktop');
    document:= desktop.getCurrentComponent;

    controller := document.getCurrentController;
    frame := controller.getFrame;
    window :=  frame.getContainerWindow;

    myListener:= TListener.Create;

    document.addEventListener(myListener);
end;

end.
