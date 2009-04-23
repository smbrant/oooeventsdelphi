unit oooEvents2;

{$WARN SYMBOL_PLATFORM OFF}

interface

uses
  ComObj, ActiveX, AxCtrls, Classes, oooEvents_TLB, StdVcl;

type
  TListener = class(TAutoObject, IConnectionPointContainer, IListener)
  private
    { Private declarations }
    FConnectionPoints: TConnectionPoints;
    FConnectionPoint: TConnectionPoint;
    FEvents: IListenerEvents;
    { note: FEvents maintains a *single* event sink. For access to more
      than one event sink, use FConnectionPoint.SinkList, and iterate
      through the list of sinks. }
  public
    procedure Initialize; override;
  protected
    { Protected declarations }
    property ConnectionPoints: TConnectionPoints read FConnectionPoints
      implements IConnectionPointContainer;
    procedure EventSinkChanged(const EventSink: IUnknown); override;
    procedure disposing; safecall;
    procedure notifyEvent(const event: IDispatch); safecall;
  end;

implementation

uses ComServ, oooEvents1;

procedure TListener.EventSinkChanged(const EventSink: IUnknown);
begin
  FEvents := EventSink as IListenerEvents;
end;

procedure TListener.Initialize;
begin
  inherited Initialize;
  FConnectionPoints := TConnectionPoints.Create(Self);
  if AutoFactory.EventTypeInfo <> nil then
    FConnectionPoint := FConnectionPoints.CreateConnectionPoint(
      AutoFactory.EventIID, ckSingle, EventConnect)
  else FConnectionPoint := nil;
end;


procedure TListener.disposing;
begin

end;

procedure TListener.notifyEvent(const event: IDispatch);
var
  e: OleVariant;
begin
  e:= event;
  form1.Memo1.Lines.Add('Event: '+e.EventName);
end;

initialization
  TAutoObjectFactory.Create(ComServer, TListener, Class_Listener,
    ciMultiInstance, tmApartment);
end.
