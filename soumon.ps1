#requires -version 2
<#@'
<?xml version ="1.0" encoding="utf-8"?>
<configuration>
  <startup useLegacyV2RuntimeActivationPolicy="true">
    <supportedRuntime version="v4.0"/>
  </startup>
</configuration>
'@|sc 'C:\EXE\powershell.exe.activation_config' -en ascii
cp 'C:\EXE\powershell.exe.activation_config' `
   'C:\EXE\powershell_ise.exe.activation_config'
[environment]::setenvironmentvariable(
  'COMPLUS_ApplicationMigrationRuntimeActivationConfigPath',
  'C:\EXE','process'#'machine' 'user'
)
sp 'registry::hklm\software\microsoft\.netframework' `
  OnlyUseLatestCLR 1 -t dword
sp 'registry::hklm\software\wow6432node\microsoft\.netframework' `
  OnlyUseLatestCLR 1 -t dword#>
param([validaterange(1,9)][int]$serv=6,[validaterange(60,300)][int]$refr=90)
set-strictmode -v latest
&{
  $erroractionpreference='stop'
  [runtime.gcsettings]::latencymode='batch'
  function dispose-after{
    param([validatenotnull()][object]$obj,[validatenotnull()][scriptblock]$sb)
    try{&$sb}
    finally{
      if($obj -as [idisposable]){
        [void][idisposable].getmethod('Dispose').invoke($obj,$null)
      }
    }
  }
  function outx{param($x)$x|out-string|write-warning}
  function json{
    add-type -a system.web.extensions
    $my.jss=new-object web.script.serialization.javascriptserializer
  }
  function rqini{
    $my.serv=$serv
    $my.sou="http://s39websouits$($my.serv).sou.local"
    $my.ars='s39ars{0:d2}' -f $my.serv
    $my.HRQH=[net.httprequestheader]
    [net.servicepointmanager]|%{
      $_::expect100continue=$false
      $_::securityprotocol='tls'
      $_::servercertificatevalidationcallback={
        param($send,$cert,$chain,$cerr)
        $true
      }
    }
    $my.u8=[text.encoding]::utf8
    $my.jl=($my.jsb='this.result={').length-1
    $my.jsf=@'
\}};;this\.cacheIDRefCount=\{{"{0}":\d+\}};;
if\(getCurWFC_NS\(this\.windowID\)!=null\)
 getCurWFC_NS\(this\.windowID\)\.status\(\[\]\);$
'@ -f $my.ars -replace '\r\n'
    $my.inc=new-object psobject -pr @{
      '№'=$null
      'Инц.'=$null
      'Ст.'=$null
      'Отв.'=$null
      'Срок'=$null
      'Кл.'=$null
      'Сопр.'=$null
      SLM=$null
      'Создан'=$null
      'Содер.'=$null
      #'Серв.'=$null;'Приор.'=$null;'ЗНО'=$null
      Id3=$null
    }
  }
  function conio{
    add-type @'
using System;
using System.Text;
using System.Runtime.InteropServices;
public static class PSConIO{
  public const uint STD_INPUT_HANDLE=unchecked((uint)(-10));
  public const uint STD_OUTPUT_HANDLE=unchecked((uint)(-11));
  //public const uint STD_ERROR_HANDLE=unchecked((uint)(-12));
  //public static readonly IntPtr INVALID_HANDLE=unchecked((IntPtr)(-1));
  public const uint ENABLE_WINDOW_INPUT=0x8;
  public const uint ENABLE_MOUSE_INPUT=0x10;
  public const uint ENABLE_QUICK_EDIT_MODE=0x40;
  public const uint ENABLE_EXTENDED_FLAGS=0x80;

  public const ushort KEY_EVENT=0x1;
  public const ushort MOUSE_EVENT=0x2;
  public const ushort WINDOW_BUFFER_SIZE_EVENT=0x4;
  public const ushort MENU_EVENT=0x8;
  public const ushort FOCUS_EVENT=0x10;

  [StructLayout(LayoutKind.Explicit,CharSet=CharSet.Unicode)]
  public struct KEY_EVENT_RECORD{
    [FieldOffset(0),MarshalAs(UnmanagedType.Bool)]
    public bool KeyDown;
    [FieldOffset(4),MarshalAs(UnmanagedType.U2)]
    public ushort RepeatCount;
    [FieldOffset(6),MarshalAs(UnmanagedType.U2)]
    public ushort VirtualKeyCode;
    [FieldOffset(8),MarshalAs(UnmanagedType.U2)]
    public ushort VirtualScanCode;
    [FieldOffset(10)]
    public char UnicodeChar;
    [FieldOffset(12),MarshalAs(UnmanagedType.U4)]
    public uint ControlKeyState;
  }
  public const ushort VK_F5=0x74;
  public const uint SHIFT_PRESSED=0x10;
  public const uint NUMLOCK_ON=0x20;
  public const uint SCROLLOCK_ON=0x40;
  public const uint CAPSLOCK_ON=0x80;

  [StructLayout(LayoutKind.Sequential)]
  public struct COORD{public short X;public short Y;}

  [StructLayout(LayoutKind.Sequential)]
  public struct MOUSE_EVENT_RECORD{
    public COORD MousePosition;
    public uint ButtonState;
    public uint ControlKeyState;
    public uint EventFlags;
  }
  public const uint FROM_LEFT_1ST_BUTTON_PRESSED=0x1;
  public const uint DOUBLE_CLICK=0x2;
  public const uint MOUSE_WHEELED=0x4;

  [StructLayout(LayoutKind.Sequential)]
  public struct WINDOW_BUFFER_SIZE_RECORD{public COORD Size;}

  [StructLayout(LayoutKind.Sequential)]
  public struct MENU_EVENT_RECORD{public uint CommandId;}

  [StructLayout(LayoutKind.Sequential)]
  public struct FOCUS_EVENT_RECORD{public int SetFocus;}

  [StructLayout(LayoutKind.Explicit)]
  public struct INPUT_RECORD{
    [FieldOffset(0)]public ushort EventType;
    [FieldOffset(4)]public KEY_EVENT_RECORD KeyEvent;
    [FieldOffset(4)]public MOUSE_EVENT_RECORD MouseEvent;
    [FieldOffset(4)]public WINDOW_BUFFER_SIZE_RECORD WindowBufferSizeEvent;
    [FieldOffset(4)]public MENU_EVENT_RECORD MenuEvent;
    [FieldOffset(4)]public FOCUS_EVENT_RECORD FocusEvent;
  }

  [DllImport("kernel32.dll",SetLastError=true)]
  public static extern IntPtr GetStdHandle(uint HandleId);

  [DllImport("kernel32.dll",SetLastError=true)]
  public static extern bool GetConsoleMode(IntPtr con,out uint Mode);

  [DllImport("kernel32.dll",SetLastError=true)]
  public static extern bool SetConsoleMode(IntPtr con,uint Mode);

  [DllImport("kernel32.dll",SetLastError=true)]
  public static extern bool FlushConsoleInputBuffer(IntPtr ConsoleInput);

  [DllImport("kernel32.dll",SetLastError=true)]
  public static extern uint WaitForSingleObject(
    IntPtr Handle,uint Milliseconds
  );
  public const uint INFINITE=0xFFFFFFFF;
  public const uint WAIT_OBJECT_0=0x0;
  public const uint WAIT_ABANDONED=0x80;
  public const uint WAIT_TIMEOUT=0x102;
  public const uint WAIT_FAILED=0xFFFFFFFF;

  [DllImport("kernel32.dll",EntryPoint="ReadConsoleInputW",
             CharSet=CharSet.Unicode,SetLastError=true)]
  public static extern bool ReadConsoleInput(
    IntPtr ConsoleInput,
    [Out] INPUT_RECORD[] Buffer,
    uint Length,
    out uint NumberOfEventsRead
  );

  [DllImport("kernel32.dll",EntryPoint="ReadConsoleOutputCharacterW",
             CharSet=CharSet.Unicode,SetLastError=true)]
  public static extern bool ReadConsoleOutputCharacter(
    IntPtr ConsoleOutput,
    [Out] StringBuilder Character,
    uint Length,
    COORD ReadCoord,
    out uint NumberOfCharsRead
  );

  public struct SMALL_RECT{
    public short Left;
    public short Top;
    public short Right;
    public short Bottom;
  }

  public struct CONSOLE_SCREEN_BUFFER_INFO{
    public COORD Size;
    public COORD CursorPosition;
    public short Attributes;
    public SMALL_RECT Window;
    public COORD MaximumWindowSize;
  }

  [DllImport("kernel32.dll",SetLastError=true)]
  public static extern bool GetConsoleScreenBufferInfo(
    IntPtr ConsoleOutput,
    out CONSOLE_SCREEN_BUFFER_INFO ConsoleScreenBufferInfo
  );

  [DllImport("kernel32.dll",SetLastError=true)]
  public static extern bool SetConsoleWindowInfo(
    IntPtr ConsoleOutput,
    bool Absolute,
    [In] ref SMALL_RECT ConsoleWindow
  );
  /*[DllImport("kernel32.dll",SetLastError=true)]
  public static extern bool ReadConsoleOutputAttribute(
    IntPtr ConsoleOutput,
    [Out] ushort[] Attribute,
    uint Length,
    COORD ReadCoord,
    out uint NumberOfAttrsRead
  );
  [DllImport("kernel32.dll",SetLastError=true)]
  public static extern bool WriteConsoleOutputAttribute(
    IntPtr ConsoleOutput,
    ushort[] Attribute,
    uint Length,
    COORD WriteCoord,
    out uint NumberOfAttrsWritten
  );
  //Text color contains blue.
  public const ushort FOREGROUND_BLUE=0x1;
  //Text color contains green.
  public const ushort FOREGROUND_GREEN=0x2;
  //Text color contains red.
  public const ushort FOREGROUND_RED=0x4;
  //Text color is intensified.
  public const ushort FOREGROUND_INTENSITY=0x8;
  //Background color contains blue.
  public const ushort BACKGROUND_BLUE=0x10;
  //Background color contains green.
  public const ushort BACKGROUND_GREEN=0x20;
  //Background color contains red.
  public const ushort BACKGROUND_RED=0x40;
  //Background color is intensified.
  public const ushort BACKGROUND_INTENSITY=0x80;
  //Leading byte.
  public const ushort COMMON_LVB_LEADING_BYTE=0x100;
  //Trailing byte.
  public const ushort COMMON_LVB_TRAILING_BYTE=0x200;
  //Top horizontal.
  public const ushort COMMON_LVB_GRID_HORIZONTAL=0x400;
  //Left vertical.
  public const ushort COMMON_LVB_GRID_LVERTICAL=0x800;
  //Right vertical.
  public const ushort COMMON_LVB_GRID_RVERTICAL=0x1000;
  //Reverse foreground and background attribute.
  public const ushort COMMON_LVB_REVERSE_VIDEO=0x4000;
  //Underscore.
  public const ushort COMMON_LVB_UNDERSCORE=0x8000;*/
}
'@
    $my.cin=[PSConIO]::getstdhandle([PSConIO]::STD_INPUT_HANDLE)
    $my.cout=[PSConIO]::getstdhandle([PSConIO]::STD_OUTPUT_HANDLE)
    $mod=0
    [void][PSConIO]::getconsolemode($my.cin,[ref]$mod)
    $mod=$mod -band (-bnot [PSConIO]::ENABLE_QUICK_EDIT_MODE)
    $mod=$mod -band (-bnot [PSConIO]::ENABLE_WINDOW_INPUT)
    $mod=$mod -bor [PSConIO]::ENABLE_MOUSE_INPUT
    $mod=$mod -bor [PSConIO]::ENABLE_EXTENDED_FLAGS
    [void][PSConIO]::setconsolemode($my.cin,$mod)
    $my.lck3=-bnot([PSConIO]::NUMLOCK_ON -bor
                   [PSConIO]::CAPSLOCK_ON -bor
                   [PSConIO]::SCROLLOCK_ON)
    [void][PSConIO]::flushconsoleinputbuffer($my.cin)
  }
  function synth{
    $err=$true
    try{
      add-type -a system.speech
      $my.syn=new-object speech.synthesis.speechsynthesizer
      foreach($iv in $my.syn.getinstalledvoices()){
        foreach($vi in $iv.voiceinfo){
          if($vi.culture -eq 'ru-ru'){
            $my.syn.selectvoice($vi.name)
            $my.syn.volume=75
            $err=$false
            break
          }
        }
      }
    }catch{outx $_}
    if($err -and $my.syn){$my.syn.dispose();$my.syn=$null}
  }
  function setbsw{param($bsw)
    if(!$my.rui.windowsize -or $my.rui.windowsize.width -le $bsw){
      $bsw,$bs.width=($bs=$my.rui.buffersize).width,$bsw
      $my.rui.buffersize=$bs
      $my.bsw=$bsw
    }
  }
  function init{
    json
    rqini
    conio
    synth
    $my.t0=[datetime]'1970-01-01'
    $my.gu='ГУ'
    $my.tu='ТУ'
    $my.ft='dd-MM-yy HH:mm'
    $my.rui=$host.ui.rawui
    setbsw 512
  }
  function mkrq{param($mth,$pth='',$acc='text/html,application/xhtml+xml,*/*')
    $rq=[net.httpwebrequest]::create("$($my.sou)/arsys/${pth}")
    $rq.accept=$acc
    $rq.allowautoredirect=$true
    $rq.headers[$my.HRQH::acceptlanguage]='ru-RU'
    $rq.automaticdecompression='gzip,deflate'
    $rq.cookiecontainer=$my.cc
    $rq.headers['DNT']=1
    $rq.method=$mth
    $rq.timeout=300000
    $rq.useragent='Mozilla/5.0 (Windows NT 6.1; Trident/7.0; rv:11.0) like Gecko'
    $rq
  }
  function sesbeg{
    $my.cc=new-object net.cookiecontainer
    $rq=mkrq 'GET'
    $rq.authenticationlevel='mutualauthrequired'
    $rq.credentials=[net.credentialcache]::defaultcredentials
    $rq.preauthenticate=$true
    dispose-after($rq.getresponse()){}
    $rq=mkrq 'POST'
    $rq.headers[$my.HRQH::pragma]='no-cache'
    $post=[byte[]][char[]]'timezone=use_server&goto=null&tzind=1'
    $rq.contentlength=$c=$post.count
    $rq.contenttype='application/x-www-form-urlencoded'
    dispose-after($st=$rq.getrequeststream()){
      $st.write($post,0,$c)
      dispose-after($rq.getresponse()){}
    }
    $ui=$my.cc.getcookies("$($my.sou)/arsys/")['MJUID'].value
$get=@'
forms/{0}/HPD:Incident%20Management%20Console/Default%
20User%20View%20(Manager)/udd.js?ui={1}&format=html&w=
'@ -f $my.ars,$ui -replace '\r\n'
    $rq=mkrq 'GET' $get 'application/javascript,*/*;q=0.8'
    dispose-after($rs=$rq.getresponse()){
      dispose-after($st=$rs.getresponsestream()){
        dispose-after($sr=new-object io.streamreader $st,$my.u8,$false){
          $ct=$sr.readtoend()
          if($ct -match ';window\.sTok="([^"]+)";'){$my.tok=$matches[1]}
        }
      }
    }
    $get="BackChannel/?param=15%2FSetOverride%2F1%2F1&sToken=$($my.tok)"
    $rq=mkrq 'GET' $get '*/*'
    $rq.contenttype='text/plain;charset=UTF-8'
    dispose-after($rq.getresponse()){}
  }
  function rqdat{param($get)
    $get=$get -f $my.ars,$my.gu,$my.tu,$my.tok -replace '\r\n'
    $get=[uri]::escapeuristring($get)
    $rq=mkrq 'GET' $get '*/*'
    $rq.ifmodifiedsince=[datetime]::utcnow
    $rq.contenttype='text/plain;charset=UTF-8'
    $t=@{}
    dispose-after($rs=$rq.getresponse()){
      dispose-after($st=$rs.getresponsestream()){
        dispose-after($sr=new-object io.streamreader $st,$my.u8,$false){
          $t.js=$sr.readtoend()
        }
      }
    }
    if($t.js.startswith($my.jsb) -and $t.js -match $my.jsf){
      $my.jss.deserializeobject(
        $t.js.substring($my.jl,$t.js.length-$my.jl-$matches[0].length+1)
      ).r
    }
  }
  <#function rqtrc{param($inc)
    $get=@'
BackChannel/?param=259/GetTableEntryList/8/{0}13/HPD:Help Desk18/Best Practice V
iew9/3013896148/{0}11/HPD:WorkLog0/1/03/1009/2/1/52/-145/1\4\1\1\1000000161\99\1
000000161\5\536870914\26/2/9/53687091410/100000016127/2/5/(1=1)15/{1}8/2/1/41/48
/2/1/11/01/12/0/2/0/&sToken={2}
'@ -f $my.ars,$inc,$my.tok -replace '\r\n'
    rqdat $get
  }#>
  function rqasgn{
    $get=@'
BackChannel/?param=441/GetTableEntryList/8/{0}31/HPD:Incident Management Console
27/Default User View (Manager)9/3020872008/{0}13/HPD:Help Desk0/1/02/-18/2/1/01/
1176/1\1\1\4\4\1\7\2\6\5\4\3\1\1\2\2\0\4\1\2\2\1\2\2\1\1\1\2\2\2\4\1\1\7\2\6\3\4
\1\1\7\2\6\4\4\1\1\7\2\6\2\4\1\1\7\2\6\1\4\1\1\1000000251\99\536870917\4\1\1\100
0000014\99\536870918\24/2/9/5368709179/53687091852/2/29/{1}15/{2}8/2/1/41/48/2/1
/01/01/02/0/2/0/&sToken={3}
'@
    rqdat $get
  }
  function rqidnt{
    $get=@'
BackChannel/?param=441/GetTableEntryList/8/{0}31/HPD:Incident Management Console
27/Default User View (Manager)9/3020872008/{0}13/HPD:Help Desk0/1/02/-18/2/1/01/
1176/1\1\1\4\4\1\7\2\6\5\4\3\1\1\2\2\0\4\1\2\2\1\2\2\1\1\1\2\2\2\4\1\1\7\2\6\3\4
\1\1\7\2\6\4\4\1\1\7\2\6\2\4\1\1\7\2\6\1\4\1\1\1000000082\99\301352400\4\1\1\100
0000010\99\301531200\24/2/9/3013524009/30153120052/2/29/{1}15/{2}8/2/1/41/48/2/1
/01/01/02/0/2/0/&sToken={3}
'@
    rqdat $get
  }
  function rqincs{
    rqasgn
    rqidnt
  }
  function sesfin{
    $rq=mkrq 'POST' 'servlet/LogoutServlet'
    $rq.keepalive=$false
    $rq.headers[$my.HRQH::pragma]='no-cache'
    $post=[byte[]][char[]]"sToken=$($my.tok)"
    $rq.contentlength=$c=$post.count
    $rq.contenttype='application/x-www-form-urlencoded'
    dispose-after($st=$rq.getrequeststream()){
      $st.write($post,0,$c)
      dispose-after($rq.getresponse()){$my.tok=$null}
    }
  }
  function fullname{param($fnm)
    $f=''
    [regex]::matches($fnm,'(\w+)')|%{$_.groups[1].value}|
    %{$f=if($f){"${f}$($_[0])."}else{"$_ "}}
    $f
  }
  function outincs{param($incs)
    cls
    $i=0
    $beep=$false
    $f="$($incs.count)".length
    $f="{0,${f}} {1}"
    $inc=$my.inc
    $ss=$dd=$rs=$a=$b=$uc=$tu=$fl=$null
    $id3=''
    $incs|sort -prop @{e={$_.i}}|?{$_.i -ne $id3}|%{
      $i++
      $ss=$_.d[7]
      $dd=$_.d[10].p
      $rs=$_.d[38].p
      $uc=([datetime]::utcnow-$my.t0).totalseconds
      $tu=$_.d[56].p
      $fl=''
      if($ss.p -ne 4 -and $dd -le $uc+60*60){
        $fl='>'
        if($tu -eq $my.tu){$beep=$true}
      }
      if($ss.p -eq 1){
        $fl="${fl}!"
        if($tu -eq $my.tu){$beep=$true}
      }
      $inc.'№'=$f -f $i,$fl
      $inc.'Инц.'=$_.d[0].p
      $inc.'Ст.'=$ss.v
      $inc.'Отв.'=if($_.d[9].p){fullname $_.d[9].p}else{$_.d[30].p}
      $inc.'Срок'=$my.t0.addseconds($dd).tolocaltime().tostring($my.ft)
      $inc.'Кл.'=fullname (($_.d[33,13,27]|%{$_.p}) -join ' ')
      $inc.'Сопр.'=$my.t0.addseconds($rs).tolocaltime().tostring($my.ft)
      $inc.SLM=$_.d[12].v
      $inc.'Создан'=$my.t0.addseconds($_.d[60].p).tolocaltime().tostring($my.ft)
      $inc.'Содер.'=$_.d[4].p
      #$inc.'Серв.'=$_.d[5].p;$inc.'Приор.'=$_.d[6].v;$inc.'ЗНО'=$_.d[3].p
      $inc.Id3=$id3=$_.i
      $inc
    }|ft '№','Инц.','Ст.','Отв.','Срок','Кл.',
         'Сопр.',SLM,'Создан','Содер.',Id3 -a
    if($beep){
      [console]::beep(1000,100)
      if($my.syn){try{$my.syn.speak('угроза премированию')}catch{}}
    }
  }
  function shincs{
    write-host 'Обновление...'
    try{
      sesbeg
      $incs=rqincs
      sesfin
      if($incs){outincs $incs}else{write-host 'Данные не получены...'}
    }catch{outx $_}
  }
  function cin{
    $shi=$true
    $wo0=[PSConIO]::WAIT_FAILED
    $ir=[PSConIO+INPUT_RECORD[]](new-object PSConIO+INPUT_RECORD)
    $csbi=new-object PSConIO+CONSOLE_SCREEN_BUFFER_INFO
    $sr=new-object PSConIO+SMALL_RECT
    $refr=1000*$refr
    $ior=$false
    $cnt=$cap=0
    $sb=new-object text.stringbuilder
    $ev=$coord=$url=$null
    $sw=[diagnostics.stopwatch]::startnew()
    for(;;){
      $my.rui.windowtitle="СОУ Монитор (23) - $(date -f 'HH:mm:ss')"
      if($shi){$sw.reset();shincs;$shi=$false;$sw.start()}
      $wo0=[PSConIO]::waitforsingleobject($my.cin,300)
      if($wo0 -ne [PSConIO]::WAIT_OBJECT_0){
        $shi=$sw.elapsedmilliseconds -ge $refr
        continue
      }
      $cnt=0
      $ior=[PSConIO]::readconsoleinput($my.cin,$ir,1,[ref]$cnt)
      if(!$ior -or !$cnt){continue}
      $ev=$ir[0]
      if($ev.eventtype -ne [PSConIO]::KEY_EVENT -and
         $ev.eventtype -ne [PSConIO]::MOUSE_EVENT){continue}
      if($ev.eventtype -eq [PSConIO]::KEY_EVENT){
        $ev=$ev.keyevent
        if($ev.keydown){continue}
        if($ev.controlkeystate -band $my.lck3){continue}
        $shi=$ev.virtualkeycode -eq [PSConIO]::VK_F5
      }else{
        $ev=$ev.mouseevent
        if($ev.eventflags -ne [PSConIO]::MOUSE_WHEELED -and
           ($ev.buttonstate -ne [PSConIO]::FROM_LEFT_1ST_BUTTON_PRESSED -or
            $ev.eventflags -ne [PSConIO]::DOUBLE_CLICK)){continue}
        if($ev.controlkeystate -band $my.lck3){continue}
        if($ev.eventflags -eq [PSConIO]::MOUSE_WHEELED){
          $ior=[PSConIO]::getconsolescreenbufferinfo($my.cout,[ref]$csbi)
          if(!$ior){continue}
          if($ev.buttonstate -band 0x80000000){
            if($csbi.window.bottom -eq $csbi.size.y-1){continue}
            $cnt=1
          }else{
            if(!$csbi.window.top){continue}
            $cnt=-1
          }
          $sr.top=$sr.bottom=$cnt
          $sr.left=$sr.right=0
          [void][PSConIO]::setconsolewindowinfo($my.cout,$false,[ref]$sr)
        }else{
          $coord=$ev.mouseposition
          $coord.x=0
          $cap=$my.rui.buffersize.width
          [void]$sb.ensurecapacity($cap)
          $ior=[PSConIO]::readconsoleoutputcharacter(
            $my.cout,$sb,$cap,$coord,[ref]$cnt
          )
          if(!$ior -or !$cnt){continue}
          [void]$sb.remove($cnt,$sb.length-$cnt)
          if("${sb}" -notmatch '\s+(INC\d+)\s*$'){continue}
          $url=@'
{0}/arsys/servlet/ViewFormServlet?form=HPD%3AHelp+Desk&server={1}&eid={2}
'@ -f $my.sou,$my.ars,$matches[1] -replace '\r\n'
          &{(new-object -c shell.application).open($url)}
        }
      }
    }
  }
  function fin{
    if($my.tok){try{sesfin}catch{outx $_}}
    if($my.syn){$my.syn.dispose()}
    if($my.bsw){setbsw $my.bsw}
  }
  $my=@{}
  $my.tok=$my.syn=$my.bsw=$null
  try{init;cin}catch{outx $_}finally{fin}
}