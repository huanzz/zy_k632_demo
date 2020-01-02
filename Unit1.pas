unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, Spin, SPComm;

type
  TForm1 = class(TForm)
    ComboBox1: TComboBox;
    Button1: TButton;
    SpinEdit1: TSpinEdit;
    Label1: TLabel;
    Label2: TLabel;
    Button2: TButton;
    Memo1: TMemo;
    Panel1: TPanel;
    Panel2: TPanel;
    Panel3: TPanel;
    Panel4: TPanel;
    Panel5: TPanel;
    Panel6: TPanel;
    Panel7: TPanel;
    Panel8: TPanel;
    Panel9: TPanel;
    Panel10: TPanel;
    Panel11: TPanel;
    Panel12: TPanel;
    Panel13: TPanel;
    Panel14: TPanel;
    Panel15: TPanel;
    Panel16: TPanel;
    Panel17: TPanel;
    Panel18: TPanel;
    Panel19: TPanel;
    Panel20: TPanel;
    Panel21: TPanel;
    Panel22: TPanel;
    Panel23: TPanel;
    Panel24: TPanel;
    Panel25: TPanel;
    Panel26: TPanel;
    Panel27: TPanel;
    Panel28: TPanel;
    Panel29: TPanel;
    Panel30: TPanel;
    Panel31: TPanel;
    Panel32: TPanel;
    Comm1: TComm;
    Button3: TButton;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    Label15: TLabel;
    Label16: TLabel;
    Label17: TLabel;
    Label18: TLabel;
    Label19: TLabel;
    Label20: TLabel;
    Label21: TLabel;
    Label22: TLabel;
    Label23: TLabel;
    Label24: TLabel;
    Label25: TLabel;
    Label26: TLabel;
    Label27: TLabel;
    Label28: TLabel;
    Label29: TLabel;
    Label30: TLabel;
    Label31: TLabel;
    Label32: TLabel;
    Label33: TLabel;
    Label34: TLabel;
    Button4: TButton;
    procedure EnumCOmPorts(Ports:TStrings);
    procedure ComConnect();
    procedure ComDisconnect();
    function SetComPort(conn:string):Boolean;
    function HexstrToInt(str:String):Integer;
    function CheckSum(Bt:array of Byte):Integer;
    procedure FormShow(Sender: TObject);
    procedure Button1Click(Sender: TObject); 
    function ComOpenOne(addr,num:Integer):Boolean;     //吸合one
    function ComCloseOne(addr,num:Integer):Boolean;   //断开one
    function ComReadAll(addr:Integer):Boolean;
    function ComOpenAll(addr:Integer):Boolean;
    function ComCloseAll(addr:Integer):Boolean;
    procedure Button3Click(Sender: TObject);
    procedure ShowMsg(str:string;sendstr:array of Byte);
    procedure ShowStstus(str:string);
    procedure DealData(str:string);
    function Split(source:string;dot:char):TStringList;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Comm1ReceiveData(Sender: TObject; Buffer: Pointer;
      BufferLength: Word);
    procedure Button4Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);

    function OpenOne(addr,num:Integer):Boolean;     //吸合one
    function CloseOne(addr,num:Integer):Boolean;   //断开one

  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  Connected:Boolean=False;


implementation

{$R *.dfm}

{ TForm1 }


function HextoBinary(hex: string): string;
const
  BOX: array [0 .. 15] of string =
    ('0000', '0001', '0010', '0011',
    '0100', '0101', '0110', '0111',
    '1000', '1001', '1010', '1011',
    '1100', '1101', '1110', '1111');
var
  i: integer;
begin
  for i := length(hex) downto 1 do
    Result := BOX[StrToInt('$' + hex[i])] + Result;
end;


function TForm1.HexstrToInt(str:String):Integer;
begin
  val('$'+str,Result,Result);
end;

function TForm1.CheckSum(Bt: array of Byte): Integer;
var
  i,sum : Integer;
begin
  sum:=0;
  for i:=0 to Length(Bt)-2 do
  begin
    sum:= sum + Bt[i];
  end;
  Result := HexstrToInt(IntToHex(sum,2));
end;

procedure TForm1.ComConnect;
begin
  try
    if not Connected then
    begin
      if Comm1.CommName <> '' then
      begin
        Comm1.StartComm;
        Connected:=True;
      end;
    end;
  except
    Connected:=False;
  end;
end;

procedure TForm1.ComDisconnect;
begin
  Comm1.StopComm;
  Connected:=False;
end;

procedure TForm1.EnumCOmPorts(Ports: TStrings);
var
  KeyHandle:HKEY;
  ErrCode,Index:Integer;
  ValueName,Data:String;
  ValueLen,DataLen,ValueType:DWORD;
  TmpPorts:TStringList;
begin
  ErrCode:=RegOpenKeyEx(HKEY_LOCAL_MACHINE, 'HARDWARE\DEVICEMAP\SERIALCOMM', 0,
    KEY_READ,KeyHandle);

  if ErrCode <> ERROR_SUCCESS then
    begin
      ShowMessage('Can not get port name');
      Exit;
    end;
  TmpPorts:=TStringList.Create;
  try
    Index:=0;
    repeat
      ValueLen:=256;
      DataLen:=256;
      SetLength(ValueName,ValueLen);
      SetLength(Data, DataLen);
      ErrCode:=RegEnumValue(KeyHandle, Index, PChar(ValueName),
        Cardinal(ValueLen),nil,@ValueType,PByte(PChar(Data)),@DataLen);

      if ErrCode = ERROR_SUCCESS then
      begin
        SetLength(Data,DataLen);
        TmpPorts.Add(Data);
        Inc(Index);
      end
      else if ErrCode <> ERROR_NO_MORE_ITEMS then
      begin
        ShowMessage('Can not get port name');
        Exit;
      end;
    Until (ErrCode <> ERROR_SUCCESS);
    TmpPorts.Sort;
    Ports.Assign(TmpPorts);
  Finally;
    RegCloseKey(KeyHandle);
  end;
end;

function TForm1.SetComPort(conn: string): Boolean;
begin
  Result := False;
  if ((conn <> '')) then
  begin
    try
      Comm1.StopComm;
      Comm1.CommName := conn;
      Comm1.BaudRate := 9600;
      Comm1.StopBits := _1;
      Comm1.ByteSize := _8;
      Comm1.Parity := None;
      Comm1.ParityCheck := False;
      Result := True;
    except
      Result := False;
    end;
  end;
end;

procedure TForm1.FormShow(Sender: TObject);
begin
  EnumCOmPorts(ComboBox1.Items);
end;

procedure TForm1.Button1Click(Sender: TObject);
begin
  if not Connected then
  begin
    if SetComPort(ComboBox1.Text) then
    begin
      ComConnect();
      Button1.Caption:='关闭串口';
      Sleep(500);
      //ComReadAll(SpinEdit1.Value);
      Button4.Click;
    end;
  end
  else
  begin
    ComDisconnect();
    Button1.Caption:='打开串口';
  end;
end;

function TForm1.ComCloseOne(addr, num: Integer): Boolean;
var
  Bt:array [0..7] of Byte;
begin
  Result:=False;
  if Connected then
  begin
    Bt[0]:=Byte($55);
    Bt[1]:=Byte(addr);         //模块组地址
    Bt[2]:=Byte($31);          //功能码：断开某一路，无返回值
    Bt[3]:=Byte($00);
    Bt[4]:=Byte($00);
    Bt[5]:=Byte($00);
    Bt[6]:=Byte(num);          //继电器地址
    Bt[7]:=Byte(CheckSum(Bt));
    Result:=Comm1.WriteCommData(@Bt,8);
    ShowMsg('Close:',Bt);
    Sleep(50);
  end;
end;

function TForm1.ComOpenOne(addr, num: Integer): Boolean;
var
  Bt:array [0..7] of Byte;
begin
  Result:=False;
  if Connected then
  begin
    Bt[0]:=Byte($55);
    Bt[1]:=Byte(addr);              //模块组地址
    Bt[2]:=Byte($32);               //功能码：吸合某一路，无返回值
    Bt[3]:=Byte($00);
    Bt[4]:=Byte($00);
    Bt[5]:=Byte($00);
    Bt[6]:=Byte(num);                 //继电器地址
    Bt[7]:=Byte(CheckSum(Bt));
    Result:=Comm1.WriteCommData(@Bt,8);
    ShowMsg('Open:',Bt);
    Sleep(50);
  end;

end;

procedure TForm1.Button3Click(Sender: TObject);
var
  Panel: TPanel;
  addr, num : Integer;
begin
  if Connected then
  begin
    Panel:=TPanel(Sender);
    addr := SpinEdit1.value;
    num := Panel.Tag;
    if  Panel.Caption = '断开' then
    begin
      //ComOpenOne(addr,num);
      OpenOne(addr,num);
      Panel.Caption := '闭合';
      Panel.Color := clGreen;
    end
    else
    begin
      //ComCloseOne(addr,num);
      CloseOne(addr,num);
      Panel.Caption := '断开';
      Panel.Color := clRed;
    end;

  end
  else
  begin
    Application.MessageBox('请打开串口再试','提示',64);
  end;


end;

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
    Comm1.StopComm;
end;

function TForm1.ComReadAll(addr: Integer): Boolean;
var
  Bt:array [0..7] of Byte;
begin
  Result:=False;
  if Connected then
  begin
    Bt[0]:=Byte($55);
    Bt[1]:=Byte(addr);              //模块组地址
    Bt[2]:=Byte($10);               //功能码：读当前继电器版状态
    Bt[3]:=Byte($FF);
    Bt[4]:=Byte($FF);
    Bt[5]:=Byte($FF);
    Bt[6]:=Byte($FF);                 //继电器地址
    Bt[7]:=Byte(CheckSum(Bt));
    Result:=Comm1.WriteCommData(@Bt,8);
    ShowMsg('Read:',Bt);
    Sleep(50);
  end;

end;

procedure TForm1.ShowMsg(str: string; sendstr: array of Byte);
var
  ii : Integer;
  Msg : string;
begin
 for ii:=0 to Length(sendstr)-1 do
    begin
      Msg:=Msg+IntToHex(sendstr[ii],2)+' ';
    end;
  Memo1.Lines.Add(str+':'+Msg);
end;

procedure TForm1.ShowStstus(str: string);
var
  charArr : array[0..31] of Char;
  ii: Integer;
  panel: TPanel;
begin
  StrCopy(@charArr,PChar(str));
  for ii:=0 to Length(charArr)-1 do
  begin
    panel:=TPanel(FindComponent('Panel'+IntToStr(32-ii)));
    if charArr[ii] = '1' then
      begin
        panel.Caption:='吸合';
        panel.Color:=clGreen;
      end
    else
      begin
        panel.Caption:='断开';
        panel.Color:=clRed;
      end;
  end;
end;

procedure TForm1.Comm1ReceiveData(Sender: TObject; Buffer: Pointer;
  BufferLength: Word);
var
  Rec:array[0..7] of byte;
  aa,bb,ii: Integer;
  pp:PChar;
  msg:String;
begin
  msg:='';
  pp:=Buffer;
  bb:=BufferLength;
  if bb>8 then bb:=8;        //单条数据长度
  for aa:=0 to bb-1 do
  begin
    Rec[aa]:=Integer(pp[aa]);
  end;
  if bb < 1 then Exit;

  for ii:=0 to Length(Rec)-1 do
    begin
      msg:=msg+IntToHex(Rec[ii],2)+' ';
    end;
  Memo1.Lines.Add('Receice:'+msg);
  DealData(msg);

end;

procedure TForm1.DealData(str: string);
var
  dataList : TStringList;
  showstr:string;
begin
  dataList := TStringList.Create;
  dataList := Split(str,' ');
  if dataList[0] = '22' then
  begin
    if dataList[2]= '10' then
    begin
      showstr:= HextoBinary(dataList[3]) + HextoBinary(dataList[4])+
        HextoBinary(dataList[5]) + HextoBinary(dataList[6]);
      ShowStstus(showstr);
    end;
  end;
end;

function TForm1.Split(source: string; dot: char): TStringList;
var
  strList:TStringList;
begin
  strList:=TStringList.Create;
  source := Trim(source);
  strList.Delimiter:=dot;
  strList.DelimitedText:=source;
  Result:= strList;
end;

procedure TForm1.Button4Click(Sender: TObject);
begin
  ComReadAll(SpinEdit1.Value);
end;

function TForm1.ComCloseAll(addr: Integer): Boolean;
var
  Bt:array [0..7] of Byte;
begin
  Result:=False;
  if Connected then
  begin
    Bt[0]:=Byte($55);
    Bt[1]:=Byte(addr);              //模块组地址
    Bt[2]:=Byte($34);               //功能码：组断开
    Bt[3]:=Byte($FF);
    Bt[4]:=Byte($FF);
    Bt[5]:=Byte($FF);
    Bt[6]:=Byte($FF);                 //继电器地址
    Bt[7]:=Byte(CheckSum(Bt));
    Result:=Comm1.WriteCommData(@Bt,8);
    ShowMsg('CloseAll:',Bt);
    Sleep(200);
  end;

end;

function TForm1.ComOpenAll(addr: Integer): Boolean;
var
  Bt:array [0..7] of Byte;
begin
  Result:=False;
  if Connected then
  begin
    Bt[0]:=Byte($55);
    Bt[1]:=Byte(addr);              //模块组地址
    Bt[2]:=Byte($35);               //功能码：组吸合
    Bt[3]:=Byte($FF);
    Bt[4]:=Byte($FF);
    Bt[5]:=Byte($FF);
    Bt[6]:=Byte($FF);                 //继电器地址
    Bt[7]:=Byte(CheckSum(Bt));
    Result:=Comm1.WriteCommData(@Bt,8);
    ShowMsg('OpenAll:',Bt);
    Sleep(200);
  end;

end;

procedure TForm1.Button2Click(Sender: TObject);
begin
  if Connected then
  begin
    if (Button2.Caption='吸合全部') then
    begin
      ComOpenAll(SpinEdit1.Value);
      ComReadAll(SpinEdit1.Value);
      Button2.Caption:='断开全部'
    end
    else
    begin
      ComCloseAll(SpinEdit1.Value);
      ComReadAll(SpinEdit1.Value);
      Button2.Caption:='吸合全部'
    end;
  end
  else
  begin
    Application.MessageBox('请打开串口再试','提示',64);
  end;

end;

function TForm1.CloseOne(addr, num: Integer): Boolean;
var
  Bt:array [0..7] of Byte;
begin
  Result:=False;
  if Connected then
  begin
    Bt[0]:=Byte($55);
    Bt[1]:=Byte(addr);              //模块组地址
    Bt[2]:=Byte($11);               //功能码：断开某一路，有返回值
    Bt[3]:=Byte($00);
    Bt[4]:=Byte($00);
    Bt[5]:=Byte($00);
    Bt[6]:=Byte(num);                 //继电器地址
    Bt[7]:=Byte(CheckSum(Bt));
    Result:=Comm1.WriteCommData(@Bt,8);
    ShowMsg('Open:',Bt);
    Sleep(50);
  end;

end;

function TForm1.OpenOne(addr, num: Integer): Boolean;
var
  Bt:array [0..7] of Byte;
begin
  Result:=False;
  if Connected then
  begin
    Bt[0]:=Byte($55);
    Bt[1]:=Byte(addr);              //模块组地址
    Bt[2]:=Byte($12);               //功能码：吸合某一路，有返回值
    Bt[3]:=Byte($00);
    Bt[4]:=Byte($00);
    Bt[5]:=Byte($00);
    Bt[6]:=Byte(num);                 //继电器地址
    Bt[7]:=Byte(CheckSum(Bt));
    Result:=Comm1.WriteCommData(@Bt,8);
    ShowMsg('Open:',Bt);
    Sleep(50);
  end;

end;

end.
