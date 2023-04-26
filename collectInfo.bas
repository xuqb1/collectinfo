#DIM ALL

#RESOURCE "collectInfo.pbr"
#INCLUDE "Encapsule.inc"
#INCLUDE "SQLlite.inc"
'
' --------------------------------------------------
#PBFORMS BEGIN INCLUDES
#IF NOT %DEF(%WINAPI)
  #INCLUDE "WIN32API.INC"
#ENDIF
#PBFORMS END INCLUDES
' --------------------------------------------------
'------------------------------------------------------------------------------
'   ** 窗体及控件ID声明 **
'------------------------------------------------------------------------------
#PBFORMS BEGIN CONSTANTS
%IDD_DIALOG1      = 101
%IDC_TYPECB       = 102
%IDC_SEARCHTB     = 103
%IDC_SEARCHBTN    = 104
%IDC_INFOLT       = 105
%IDC_ADDBTN       = 106
%IDC_EDITBTN      = 107
%IDC_DELBTN       = 108
%IDC_CLOSEBTN     = 109
%IDC_PAGEINFOLB   = 110
%IDC_PREPAGEBTN   = 111
%IDC_CURRPAGETB   = 112
%IDC_GOPAGEBTN    = 113
%IDC_NEXTPAGEBTN  = 114
%IDC_SETTOPBTN    = 115
%IDC_TYPEBTN      = 116

%IDC_TYPELB       = 120
%IDC_TITLELB      = 121
%IDC_TITLETB      = 122
%IDC_CONTENTLB    = 123
%IDC_CONTENTTB    = 124

%IDC_EDIT101      = 101
%IDC_LABEL102     = 102

%IDC_TYPELT       = 130

%WM_TRAY          = 205

#PBFORMS END CONSTANTS
'------------------------------------------------------------------------------
GLOBAL dbname   AS STRING
GLOBAL hDB      AS LONG
GLOBAL pSize    AS LONG
GLOBAL hMenu    AS LONG
GLOBAL msgTitle AS STRING
'------------------------------------------------------------------------------

FUNCTION PBMAIN()
  ShowDIALOG1 %HWND_DESKTOP
END FUNCTION
'------------------------------------------------------------------------------
'   ** 对话框1 **
'------------------------------------------------------------------------------
FUNCTION ShowDIALOG1(BYVAL hParent AS DWORD) AS LONG
  LOCAL lRslt     AS LONG
  LOCAL tmpStr    AS STRING
  LOCAL typeStr() AS STRING

  InitGlobalVars

  #PBFORMS BEGIN DIALOG %IDD_DIALOG1->->
    LOCAL hDlg  AS DWORD

    '%WS_MAXIMIZEBOX OR
    DIALOG NEW hParent, "收集信息", , , 260, 240, %WS_POPUP OR _
                        %WS_BORDER OR %WS_DLGFRAME OR %WS_SYSMENU OR %WS_MINIMIZEBOX OR _
                        %WS_CLIPSIBLINGS OR %WS_VISIBLE OR %DS_MODALFRAME OR %WS_THICKFRAME _
                        OR %DS_3DLOOK OR %DS_NOFAILCREATE OR %DS_SETFONT, _
                        %WS_EX_CONTROLPARENT OR %WS_EX_LEFT OR %WS_EX_LTRREADING OR _
                        %WS_EX_RIGHTSCROLLBAR, TO hDlg
    DIALOG SET ICON hDlg,"APPICON"
    tmpStr="全部," & GetAllTypeName(hDlg)
    REDIM typeStr(PARSECOUNT(tmpStr,",")-1)
    PARSE tmpStr,typeStr(),","

    CONTROL ADD COMBOBOX,hDlg,%IDC_TYPECB,typeStr(),5,5,60,60
    CONTROL SET TEXT hDlg,%IDC_TYPECB,"全部"
    CONTROL ADD TEXTBOX,hDlg, %IDC_SEARCHTB,"",66,5,150,12
    CONTROL ADD BUTTON, hDlg, %IDC_SEARCHBTN,   "查询",217,5,35,12,%BS_CENTER OR %BS_DEFAULT OR %BS_VCENTER _
        OR %WS_TABSTOP,%WS_EX_LEFT

    CONTROL ADD LISTBOX,hDlg, %IDC_INFOLT,,5,19,211,200,%LBS_EXTENDEDSEL OR %LBS_MULTIPLESEL _
        OR %LBS_NOINTEGRALHEIGHT OR %LBS_NOTIFY OR %LBS_USETABSTOPS OR %WS_TABSTOP OR %WS_VSCROLL, _
        %WS_EX_CLIENTEDGE

    CONTROL ADD BUTTON, hDlg, %IDC_ADDBTN,      "增加",217,30,35,25
    CONTROL ADD BUTTON, hDlg, %IDC_EDITBTN,     "编辑",217,65,35,25
    CONTROL ADD BUTTON, hDlg, %IDC_DELBTN,      "删除",217,100,35,25
    CONTROL ADD BUTTON, hDlg, %IDC_TYPEBTN,     "分类",217,100,35,25
    CONTROL ADD CHECKBOX, hDlg, %IDC_SETTOPBTN, "置顶",217,140,35,25,%BS_PUSHLIKE OR %BS_CENTER OR %BS_VCENTER _
        OR %WS_TABSTOP,%WS_EX_LEFT
    CONTROL ADD BUTTON, hDlg, %IDCANCEL,        "关闭",217,180,35,25

    CONTROL ADD LABEL,  hDlg, %IDC_PAGEINFOLB,  "",152,225,100,12
    CONTROL ADD BUTTON, hDlg, %IDC_PREPAGEBTN,  "< 上一页",5,222,40,14
    CONTROL ADD TEXTBOX,hDlg, %IDC_CURRPAGETB,  "1",53,222,25,14
    CONTROL ADD BUTTON, hDlg, %IDC_GOPAGEBTN,   "转到",79,222,25,14
    CONTROL ADD BUTTON, hDlg, %IDC_NEXTPAGEBTN, "下一页 >",110,222,40,14
    MENU NEW POPUP TO hMenu
    MENU ADD STRING, hMenu, "主窗口", 401, %MF_ENABLED
    MENU ADD STRING, hMenu, "关于",   402, %MF_ENABLED
    MENU ADD STRING, hMenu, "退出",   403, %MF_ENABLED
  #PBFORMS END DIALOG

  DIALOG SHOW MODAL hDlg, CALL ShowDIALOG1Proc TO lRslt

  #PBFORMS BEGIN CLEANUP %IDD_DIALOG1
  #PBFORMS END CLEANUP

  FUNCTION = lRslt
END FUNCTION
'------------------------------------------------------------------------------
'   ** 对话框回调 **
'------------------------------------------------------------------------------
CALLBACK FUNCTION ShowDIALOG1Proc()
  LOCAL   tmpLng  AS LONG
  LOCAL   tmpStr  AS STRING
  LOCAL   xx,yy   AS LONG
  STATIC  hIcon1  AS LONG
  STATIC  ti      AS NOTIFYICONDATA
  STATIC  p       AS POINTAPI

  SELECT CASE AS LONG CBMSG
    CASE %WM_INITDIALOG
      ' Initialization handler
      hIcon1=LoadImage(0, "icon/list16.ico", %IMAGE_ICON, 0, 0, %LR_LOADFROMFILE)
      ti.cbSize = SIZEOF(ti)
      ti.hWnd = CB.HNDL
      ti.uID = GetWindowLong(CB.HNDL,%GWL_HINSTANCE)   'hInst
      ti.uFlags = %NIF_ICON OR %NIF_MESSAGE OR %NIF_TIP
      ti.uCallbackMessage = %WM_TRAY
      ti.hIcon = LoadIcon(GetWindowLong(CB.HNDL,%GWL_HINSTANCE),"APPICON")  'hIcon1'
      ti.szTip = "收集信息"
      Shell_NotifyIcon %NIM_ADD, ti
      DoQuery CB.HNDL
      DIALOG SHOW STATE CB.HNDL,%SW_HIDE
    CASE %WM_TRAY
      SELECT CASE AS LONG LOWRD(CB.LPARAM)
        CASE %WM_RBUTTONDOWN
          SetForegroundWindow CB.HNDL
          GetCursorPos p
          TrackPopupMenu hMenu, %TPM_BOTTOMALIGN OR %TPM_RIGHTALIGN,p.x,p.y, 0, CB.HNDL, BYVAL %NULL

        CASE %WM_LBUTTONDBLCLK
          DIALOG SHOW STATE CB.HNDL,%SW_RESTORE

      END SELECT
    CASE %WM_NCACTIVATE
      STATIC hWndSaveFocus AS DWORD
      IF ISFALSE CBWPARAM THEN
        ' Save control focus
        hWndSaveFocus = GetFocus()
      ELSEIF hWndSaveFocus THEN
        ' Restore control focus
        SetFocus(hWndSaveFocus)
        hWndSaveFocus = 0
      END IF
    CASE %WM_SIZE
      IF CBWPARAM = %SIZE_MINIMIZED THEN EXIT FUNCTION
      DIALOG GET CLIENT CB.HNDL TO xx,yy
      IF xx<250 THEN
        xx=250
        DIALOG SET CLIENT CB.HNDL,xx,yy
      END IF
      IF yy<220 THEN
        yy=220
        DIALOG SET CLIENT CB.HNDL,xx,yy
      END IF
      CONTROL SET SIZE  CB.HNDL, %IDC_SEARCHTB,   xx-105, 12
      CONTROL SET LOC   CB.HNDL, %IDC_SEARCHBTN,  xx-38,4
      CONTROL SET SIZE  CB.HNDL, %IDC_INFOLT,     xx-45, yy-40
      CONTROL SET LOC   CB.HNDL, %IDC_ADDBTN,     xx-38, 25
      CONTROL SET LOC   CB.HNDL, %IDC_EDITBTN,    xx-38, 60
      CONTROL SET LOC   CB.HNDL, %IDC_DELBTN,     xx-38, 90
      CONTROL SET LOC   CB.HNDL, %IDC_TYPEBTN,    xx-38, 120
      CONTROL SET LOC   CB.HNDL, %IDC_SETTOPBTN,  xx-38, 150
      CONTROL SET LOC   CB.HNDL, %IDCANCEL,       xx-38, 180
      CONTROL SET LOC   CB.HNDL, %IDC_PREPAGEBTN, 5,yy-18
      CONTROL SET LOC   CB.HNDL, %IDC_CURRPAGETB, 53,yy-18
      CONTROL SET LOC   CB.HNDL, %IDC_GOPAGEBTN,  79,yy-18
      CONTROL SET LOC   CB.HNDL, %IDC_NEXTPAGEBTN,110,yy-18
      CONTROL SET LOC   CB.HNDL, %IDC_PAGEINFOLB, 155,yy-15
      CONTROL SET SIZE  CB.HNDL, %IDC_PAGEINFOLB, xx-160,12
      DIALOG REDRAW CB.HNDL

    CASE %WM_COMMAND
      IF CBCTLMSG <> %BN_CLICKED AND CBCTLMSG <> 1 AND CBCTLMSG<>%LBN_DBLCLK THEN
        EXIT FUNCTION
      END IF
      SELECT CASE AS LONG CBCTL

        CASE 401 '主窗口菜单
          ' First, update the dialog message with latest received message
          'CONTROL SET TEXT hDlg&, 110, UDPMessage
          DIALOG SHOW STATE CB.HNDL,%SW_RESTORE 'hDlg& CALL CB_Dlg

        ' About
        CASE 402 '关于菜单
            MSGBOX "收集信息 by xsoft" _
            & CHR$(13,10,10) & "本程序可自由复制使用。" _
            & CHR$(13,10) & "Enjoy!", %MB_OK + %MB_ICONINFORMATION,"关于"
        ' Quit
        CASE 403 '退出菜单
          DIALOG END CB.HNDL,0

        CASE %IDC_SEARCHBTN   '查询按钮
          DoQuery CB.HNDL

        CASE %IDC_ADDBTN      '增加按钮
          AddDlg CB.HNDL
          ResetType CB.HNDL
          DoQuery CB.HNDL
        CASE %IDC_EDITBTN     '编辑按钮
          LISTBOX GET SELECT CB.HNDL,%IDC_INFOLT TO tmpLng
          IF tmpLng=0 THEN
            InfoBox CB.HNDL,"请先选择要编辑的条目"
            EXIT FUNCTION
          END IF
          EditDlg CB.HNDL
          ResetType CB.HNDL
          DoQuery CB.HNDL

        CASE %IDC_DELBTN      '删除按钮
          LISTBOX GET SELECT CB.HNDL,%IDC_INFOLT TO tmpLng
          IF tmpLng=0 THEN
            InfoBox CB.HNDL,"请先选择要删除的条目",%MB_OK OR %MB_ICONWARNING
            EXIT FUNCTION
          END IF
          IF InfoBox(CB.HNDL,"删除将不可恢复。你确定要删除吗？",%MB_OKCANCEL OR %MB_ICONQUESTION)=%IDOK THEN
            DoDelete CB.HNDL
            DoQuery CB.HNDL
          END IF

        CASE %IDC_TYPEBTN     '分类按钮
          TypeDlg CB.HNDL
          ResetType CB.HNDL
          DoQuery CB.HNDL

        CASE %IDC_INFOLT      '列表双击
          IF CBCTLMSG=%LBN_DBLCLK THEN
            EditDlg CB.HNDL
            ResetType CB.HNDL
            DoQuery CB.HNDL
          END IF

        CASE %IDC_PREPAGEBTN  '上一页按钮
          CONTROL GET TEXT CB.HNDL,%IDC_CURRPAGETB TO tmpStr
          tmpStr=TRIM$(tmpStr)
          IF tmpStr="" THEN
            tmpStr="1"
          END IF
          IF VAL(tmpStr)<=1 THEN
            tmpStr="1"
            CONTROL SET TEXT CB.HNDL,%IDC_CURRPAGETB,tmpStr
            EXIT FUNCTION
          END IF
          tmpStr=FORMAT$(VAL(tmpStr)-1)
          CONTROL SET TEXT CB.HNDL,%IDC_CURRPAGETB,tmpStr
          DoQuery CB.HNDL

        CASE %IDC_GOPAGEBTN   '跳转按钮
          CONTROL GET TEXT CB.HNDL,%IDC_CURRPAGETB TO tmpStr
          tmpStr=TRIM$(tmpStr)
          IF tmpStr="" THEN
            tmpStr="1"
          END IF
          IF VAL(tmpStr)<=1 THEN
            tmpStr="1"
            CONTROL SET TEXT CB.HNDL,%IDC_CURRPAGETB,tmpStr
            'EXIT FUNCTION
          END IF
          'tmpStr=FORMAT$(VAL(tmpStr)+1)
          CONTROL SET TEXT CB.HNDL,%IDC_CURRPAGETB,tmpStr
          DoQuery CB.HNDL

        CASE %IDC_NEXTPAGEBTN '下一页按钮
          CONTROL GET TEXT CB.HNDL,%IDC_CURRPAGETB TO tmpStr
          tmpStr=TRIM$(tmpStr)
          IF tmpStr="" THEN
            tmpStr="1"
          END IF
          IF VAL(tmpStr)<=1 THEN
            tmpStr="1"
            CONTROL SET TEXT CB.HNDL,%IDC_CURRPAGETB,tmpStr
            'EXIT FUNCTION
          END IF
          tmpStr=FORMAT$(VAL(tmpStr)+1)
          CONTROL SET TEXT CB.HNDL,%IDC_CURRPAGETB,tmpStr
          DoQuery CB.HNDL

        CASE %IDC_SETTOPBTN   '切换置顶按钮
          CONTROL GET CHECK CB.HNDL,%IDC_SETTOPBTN TO tmpLng
          IF ISTRUE tmpLng THEN
            SetWindowPos CB.HNDL,%HWND_TOPMOST,0,0,0,0,%SWP_NOMOVE OR %SWP_NOSIZE OR %SWP_SHOWWINDOW
          ELSE
            SetWindowPos CB.HNDL,%HWND_NOTOPMOST,0,0,0,0,%SWP_NOMOVE OR %SWP_NOSIZE OR %SWP_SHOWWINDOW
          END IF

        CASE %IDCANCEL        '关闭按钮
          ShowWindow CB.HNDL, %SW_HIDE
          FUNCTION = 1
          EXIT FUNCTION

      END SELECT
    CASE %WM_DESTROY
      Shell_NotifyIcon %NIM_DELETE, ti
    CASE %WM_SYSCOMMAND
      SELECT CASE (CB.WPARAM AND &H0FFF0)
        CASE %SC_CLOSE
          ShowWindow CB.HNDL, %SW_HIDE
          FUNCTION = 1
          EXIT FUNCTION
        CASE %SC_MINIMIZE
          ShowWindow CB.HNDL, %SW_HIDE
          FUNCTION = 1
          EXIT FUNCTION
      END SELECT
  END SELECT
END FUNCTION
'------------------------------------------------------------------------------
FUNCTION AddDlg(BYVAL hParent AS DWORD)AS LONG
  LOCAL hDlg      AS DWORD
  LOCAL tmpStr    AS STRING
  LOCAL typeStr() AS STRING
  LOCAL lRslt     AS LONG

  DIALOG NEW hParent,"增加条目",,,200,153,%WS_POPUP OR _
                        %WS_BORDER OR %WS_DLGFRAME OR %WS_SYSMENU OR %WS_MINIMIZEBOX OR _
                         %WS_CLIPSIBLINGS OR %WS_VISIBLE OR %DS_MODALFRAME OR %WS_THICKFRAME _
                        OR %DS_3DLOOK OR %DS_NOFAILCREATE OR %DS_SETFONT, _
                        %WS_EX_CONTROLPARENT OR %WS_EX_LEFT OR %WS_EX_LTRREADING OR _
                        %WS_EX_RIGHTSCROLLBAR, TO hDlg
    CONTROL ADD LABEL,  hDlg,%IDC_TYPELB,"分类:",5,6,25,12
    tmpStr="全部," & GetAllTypeName(hDlg)
    REDIM typeStr(PARSECOUNT(tmpStr,",")-1)
    PARSE tmpStr,typeStr(),","
    CONTROL ADD COMBOBOX,hDlg,%IDC_TYPECB,typeStr(),30,5,60,60
    CONTROL GET TEXT hParent,%IDC_TYPECB TO tmpStr
    tmpStr=TRIM$(tmpStr)
    CONTROL SET TEXT hDlg,%IDC_TYPECB,tmpStr

    CONTROL ADD LABEL,  hDlg, %IDC_TITLELB,   "标题:",5,20,25,12
    CONTROL ADD TEXTBOX,hDlg, %IDC_TITLETB,   "",30,19,150,12

    CONTROL ADD LABEL,  hDlg, %IDC_CONTENTLB, "内容:",5,34,25,12
    CONTROL ADD TEXTBOX,hDlg, %IDC_CONTENTTB, "",30,33,150,80,%ES_LEFT OR %ES_MULTILINE _
        OR %ES_AUTOHSCROLL OR %WS_HSCROLL OR %ES_AUTOVSCROLL OR %WS_VSCROLL _
        OR %ES_WANTRETURN  OR %WS_TABSTOP,%WS_EX_CLIENTEDGE

    CONTROL ADD BUTTON, hDlg, %IDC_ADDBTN,    "增加",40,123,40,20
    CONTROL ADD BUTTON, hDlg, %IDOK,          "确定",90,123,40,20
    CONTROL ADD BUTTON, hDlg, %IDCANCEL,      "关闭",140,123,40,20

    DIALOG SHOW MODAL hDlg, CALL AddDlgProc TO lRslt
END FUNCTION
'------------------------------------------------------------------------------
CALLBACK FUNCTION AddDlgProc()AS LONG
  STATIC hParent  AS DWORD
  LOCAL  tmpStr   AS STRING
  LOCAL  xx,yy    AS LONG
  LOCAL typeStr   AS STRING

  SELECT CASE AS LONG CBMSG
    CASE %WM_INITDIALOG
      ' Initialization handler
      WINDOW GET PARENT CB.HNDL TO hParent
      CONTROL GET TEXT hParent,%IDC_SEARCHTB TO tmpStr
      CONTROL SET TEXT CB.HNDL,%IDC_TITLETB,tmpStr

      LISTBOX GET TEXT hParent,%IDC_INFOLT TO tmpStr
      typeStr=TRIM$(PARSE$(tmpStr,",",3))
      CONTROL SET TEXT CB.HNDL,%IDC_TYPECB,typeStr
    CASE %WM_NCACTIVATE
      STATIC hWndSaveFocus AS DWORD
      IF ISFALSE CBWPARAM THEN
        ' Save control focus
        hWndSaveFocus = GetFocus()
      ELSEIF hWndSaveFocus THEN
        ' Restore control focus
        SetFocus(hWndSaveFocus)
        hWndSaveFocus = 0
      END IF
    CASE %WM_SIZE
      IF CBWPARAM = %SIZE_MINIMIZED THEN EXIT FUNCTION
      DIALOG GET CLIENT CB.HNDL TO xx,yy
      IF xx<200 THEN
        xx=200
        DIALOG SET CLIENT CB.HNDL,xx,yy
      END IF
      IF yy<153 THEN
        yy=153
        DIALOG SET CLIENT CB.HNDL,xx,yy
      END IF
      CONTROL SET SIZE  CB.HNDL, %IDC_TITLETB,    xx-50, 12
      CONTROL SET SIZE  CB.HNDL, %IDC_CONTENTTB,  xx-50, yy-73
      CONTROL SET LOC   CB.HNDL, %IDC_ADDBTN,     xx-160, yy-30
      CONTROL SET LOC   CB.HNDL, %IDOK,           xx-110, yy-30
      CONTROL SET LOC   CB.HNDL, %IDCANCEL,       xx-60, yy-30
      DIALOG REDRAW     CB.HNDL
    CASE %WM_COMMAND
      ' Process control notifications
      IF CBCTLMSG <> %BN_CLICKED AND CBCTLMSG <> 1 THEN
        EXIT FUNCTION
      END IF
      SELECT CASE AS LONG CBCTL
        CASE %IDC_ADDBTN    '增加按钮
          CONTROL GET TEXT CB.HNDL,%IDC_TYPECB TO tmpStr
          tmpStr=TRIM$(tmpStr)
          IF tmpStr="" THEN
            InfoBox CB.HNDL,"请选择或填写分类"
            CONTROL SET FOCUS CB.HNDL,%IDC_TYPECB
            EXIT FUNCTION
          END IF
          CONTROL GET TEXT CB.HNDL,%IDC_TITLETB TO tmpStr
          tmpStr=TRIM$(tmpStr)
          IF tmpStr="" THEN
            InfoBox CB.HNDL,"请填写标题"
            CONTROL SET FOCUS CB.HNDL,%IDC_TITLETB
            EXIT FUNCTION
          END IF
          CONTROL GET TEXT CB.HNDL,%IDC_CONTENTTB TO tmpStr
          tmpStr=TRIM$(tmpStr)
          IF tmpStr="" THEN
            InfoBox CB.HNDL,"请填写内容"
            CONTROL SET FOCUS CB.HNDL,%IDC_CONTENTTB
            EXIT FUNCTION
          END IF
          DoAdd CB.HNDL
        CASE %IDOK          '确定按钮
          CONTROL GET TEXT CB.HNDL,%IDC_TYPECB TO tmpStr
          tmpStr=TRIM$(tmpStr)
          IF tmpStr="" THEN
            InfoBox CB.HNDL,"请选择或填写分类"
            CONTROL SET FOCUS CB.HNDL,%IDC_TYPECB
            EXIT FUNCTION
          END IF
          CONTROL GET TEXT CB.HNDL,%IDC_TITLETB TO tmpStr
          tmpStr=TRIM$(tmpStr)
          IF tmpStr="" THEN
            InfoBox CB.HNDL,"请填写标题"
            CONTROL SET FOCUS CB.HNDL,%IDC_TITLETB
            EXIT FUNCTION
          END IF
          CONTROL GET TEXT CB.HNDL,%IDC_CONTENTTB TO tmpStr
          tmpStr=TRIM$(tmpStr)
          IF tmpStr="" THEN
            InfoBox CB.HNDL,"请填写内容"
            CONTROL SET FOCUS CB.HNDL,%IDC_CONTENTTB
            EXIT FUNCTION
          END IF
          IF DoAdd(CB.HNDL)>0 THEN
            DIALOG END CB.HNDL,0
          END IF
        CASE %IDCANCEL      '取消按钮
          DIALOG END CB.HNDL,0
      END SELECT
  END SELECT
END FUNCTION
'------------------------------------------------------------------------------
FUNCTION EditDlg(BYVAL hParent AS DWORD)AS LONG
  LOCAL hDlg      AS DWORD
  LOCAL tmpStr    AS STRING
  LOCAL typeStr() AS STRING
  LOCAL lRslt     AS LONG

  DIALOG NEW hParent,"编辑条目",,,200,153,%WS_POPUP OR _
                        %WS_BORDER OR %WS_DLGFRAME OR %WS_SYSMENU OR %WS_MINIMIZEBOX OR _
                         %WS_CLIPSIBLINGS OR %WS_VISIBLE OR %DS_MODALFRAME OR %WS_THICKFRAME _
                        OR %DS_3DLOOK OR %DS_NOFAILCREATE OR %DS_SETFONT, _
                        %WS_EX_CONTROLPARENT OR %WS_EX_LEFT OR %WS_EX_LTRREADING OR _
                        %WS_EX_RIGHTSCROLLBAR, TO hDlg
    CONTROL ADD LABEL,  hDlg,%IDC_TYPELB,"分类:",5,6,25,12
    tmpStr="全部," & GetAllTypeName(hDlg)
    REDIM typeStr(PARSECOUNT(tmpStr,",")-1)
    PARSE tmpStr,typeStr(),","
    CONTROL ADD COMBOBOX,hDlg,%IDC_TYPECB,typeStr(),30,5,60,60
    CONTROL GET TEXT hParent,%IDC_TYPECB TO tmpStr
    tmpStr=TRIM$(tmpStr)
    CONTROL SET TEXT hDlg,%IDC_TYPECB,tmpStr

    CONTROL ADD LABEL,  hDlg, %IDC_TITLELB,   "标题:",5,20,25,12
    CONTROL ADD TEXTBOX,hDlg, %IDC_TITLETB,   "",30,19,150,12

    CONTROL ADD LABEL,  hDlg, %IDC_CONTENTLB, "内容:",5,34,25,12
    CONTROL ADD TEXTBOX,hDlg, %IDC_CONTENTTB, "",30,33,150,80,%ES_LEFT OR %ES_MULTILINE _
        OR %ES_AUTOHSCROLL OR %WS_HSCROLL OR %ES_AUTOVSCROLL OR %WS_VSCROLL _
        OR %ES_WANTRETURN  OR %WS_TABSTOP,%WS_EX_CLIENTEDGE

    CONTROL ADD BUTTON, hDlg, %IDOK,          "确定",90,123,40,20
    CONTROL ADD BUTTON, hDlg, %IDCANCEL,      "关闭",140,123,40,20

    DIALOG SHOW MODAL hDlg, CALL EditDlgProc TO lRslt
END FUNCTION
'------------------------------------------------------------------------------
CALLBACK FUNCTION EditDlgProc()AS LONG
  STATIC hParent  AS DWORD
  STATIC idStr    AS STRING
  LOCAL  tmpStr   AS STRING
  LOCAL  typeStr  AS STRING
  LOCAL  xx,yy    AS LONG

  SELECT CASE AS LONG CBMSG
    CASE %WM_INITDIALOG
      ' Initialization handler
      WINDOW GET PARENT CB.HNDL TO hParent
      LISTBOX GET TEXT hParent,%IDC_INFOLT TO tmpStr
      IF tmpStr = "" THEN
        AddDlg hParent
        DIALOG END CB.HNDL,0
        EXIT FUNCTION
      END IF
      idStr=TRIM$(PARSE$(tmpStr,",",1))
      typeStr=TRIM$(PARSE$(tmpStr,",",3))
      CONTROL SET TEXT CB.HNDL,%IDC_TYPECB,typeStr
      CONTROL SET TEXT CB.HNDL,%IDC_TITLETB,TRIM$(PARSE$(tmpStr,",",2))
      tmpStr=MID$(tmpStr,INSTR(tmpStr,typeStr+",")+LEN(typeStr)+2)
      REPLACE "\n" WITH $CRLF IN tmpStr
      IF RIGHT$(tmpStr,2)=", " THEN
        tmpStr=MID$(tmpStr,1,LEN(tmpStr)-2)
      END IF
      CONTROL SET TEXT CB.HNDL,%IDC_CONTENTTB,tmpStr
    CASE %WM_NCACTIVATE
      STATIC hWndSaveFocus AS DWORD
      IF ISFALSE CBWPARAM THEN
        ' Save control focus
        hWndSaveFocus = GetFocus()
      ELSEIF hWndSaveFocus THEN
        ' Restore control focus
        SetFocus(hWndSaveFocus)
        hWndSaveFocus = 0
      END IF

    CASE %WM_SIZE
      IF CBWPARAM = %SIZE_MINIMIZED THEN EXIT FUNCTION
      DIALOG GET CLIENT CB.HNDL TO xx,yy
      IF xx<200 THEN
        xx=200
        DIALOG SET CLIENT CB.HNDL,xx,yy
      END IF
      IF yy<153 THEN
        yy=153
        DIALOG SET CLIENT CB.HNDL,xx,yy
      END IF
      CONTROL SET SIZE  CB.HNDL,  %IDC_TITLETB,   xx-50, 12
      CONTROL SET SIZE  CB.HNDL,  %IDC_CONTENTTB, xx-50, yy-73
      CONTROL SET LOC   CB.HNDL,  %IDOK,          xx-110, yy-30
      CONTROL SET LOC   CB.HNDL,  %IDCANCEL,      xx-60, yy-30
      DIALOG REDRAW     CB.HNDL

    CASE %WM_COMMAND
      ' Process control notifications
      IF CBCTLMSG <> %BN_CLICKED AND CBCTLMSG <> 1 THEN
        EXIT FUNCTION
      END IF
      SELECT CASE AS LONG CBCTL
        CASE %IDOK          '确定按钮
          CONTROL GET TEXT CB.HNDL,%IDC_TYPECB TO tmpStr
          tmpStr=TRIM$(tmpStr)
          IF tmpStr="" THEN
            InfoBox CB.HNDL,"请选择或填写分类"
            CONTROL SET FOCUS CB.HNDL,%IDC_TYPECB
            EXIT FUNCTION
          END IF
          CONTROL GET TEXT CB.HNDL,%IDC_TITLETB TO tmpStr
          tmpStr=TRIM$(tmpStr)
          IF tmpStr="" THEN
            InfoBox CB.HNDL,"请填写标题"
            CONTROL SET FOCUS CB.HNDL,%IDC_TITLETB
            EXIT FUNCTION
          END IF
          CONTROL GET TEXT CB.HNDL,%IDC_CONTENTTB TO tmpStr
          tmpStr=TRIM$(tmpStr)
          IF tmpStr="" THEN
            InfoBox CB.HNDL,"请填写内容"
            CONTROL SET FOCUS CB.HNDL,%IDC_CONTENTTB
            EXIT FUNCTION
          END IF
          IF DoUpdate(CB.HNDL,idStr)>0 THEN
            DIALOG END CB.HNDL,0
          END IF
        CASE %IDCANCEL      '取消按钮
          DIALOG END CB.HNDL,0
      END SELECT
  END SELECT
END FUNCTION
'------------------------------------------------------------------------------
FUNCTION TypeDlg(BYVAL hParent AS DWORD)AS LONG
  LOCAL hDlg  AS DWORD
  LOCAL lRslt AS LONG

  DIALOG NEW hParent,"管理分类",,,200,153,%WS_POPUP OR _
                        %WS_BORDER OR %WS_DLGFRAME OR %WS_SYSMENU OR %WS_MINIMIZEBOX OR _
                         %WS_CLIPSIBLINGS OR %WS_VISIBLE OR %DS_MODALFRAME OR %WS_THICKFRAME _
                        OR %DS_3DLOOK OR %DS_NOFAILCREATE OR %DS_SETFONT, _
                        %WS_EX_CONTROLPARENT OR %WS_EX_LEFT OR %WS_EX_LTRREADING OR _
                        %WS_EX_RIGHTSCROLLBAR, TO hDlg
  CONTROL ADD LISTBOX,hDlg, %IDC_TYPELT,,5,5,145,143,%LBS_EXTENDEDSEL OR %LBS_MULTIPLESEL _
        OR %LBS_NOINTEGRALHEIGHT OR %LBS_NOTIFY OR %LBS_USETABSTOPS OR %WS_TABSTOP OR %WS_VSCROLL, _
        %WS_EX_CLIENTEDGE
  CONTROL ADD BUTTON, hDlg, %IDC_ADDBTN,  "增加",40,20,40,20
  CONTROL ADD BUTTON, hDlg, %IDC_EDITBTN, "编辑",40,50,40,20
  CONTROL ADD BUTTON, hDlg, %IDC_DELBTN,  "删除",40,80,40,20
  CONTROL ADD BUTTON, hDlg, %IDCANCEL,    "关闭",140,120,40,20

  DIALOG SHOW MODAL hDlg, CALL TypeDlgProc TO lRslt
END FUNCTION
'------------------------------------------------------------------------------
CALLBACK FUNCTION TypeDlgProc()AS LONG
  STATIC hParent      AS DWORD
  LOCAL  tmpStr       AS STRING
  LOCAL  xx,yy        AS LONG
  LOCAL  rsStr        AS STRING
  LOCAL  currTypeName AS STRING
  LOCAL  i            AS LONG

  SELECT CASE AS LONG CBMSG
    CASE %WM_INITDIALOG
      ' Initialization handler
      WINDOW GET PARENT CB.HNDL TO hParent
      'CONTROL GET TEXT hParent,%IDC_SEARCHTB TO tmpStr
      'CONTROL SET TEXT CB.HNDL,%IDC_TITLETB,tmpStr
      rsStr=GetAllType(CB.HNDL)
      LISTBOX RESET CB.HNDL,%IDC_TYPELT
      FOR i=1 TO PARSECOUNT(rsStr,$CRLF)
        LISTBOX ADD CB.HNDL,%IDC_TYPELT,PARSE$(rsStr,$CRLF,i)
      NEXT i
    CASE %WM_NCACTIVATE
      STATIC hWndSaveFocus AS DWORD
      IF ISFALSE CBWPARAM THEN
        ' Save control focus
        hWndSaveFocus = GetFocus()
      ELSEIF hWndSaveFocus THEN
        ' Restore control focus
        SetFocus(hWndSaveFocus)
        hWndSaveFocus = 0
      END IF
    CASE %WM_SIZE
      IF CBWPARAM = %SIZE_MINIMIZED THEN EXIT FUNCTION
      DIALOG GET CLIENT CB.HNDL TO xx,yy
      IF xx<200 THEN
        xx=200
        DIALOG SET CLIENT CB.HNDL,xx,yy
      END IF
      IF yy<153 THEN
        yy=153
        DIALOG SET CLIENT CB.HNDL,xx,yy
      END IF
      CONTROL SET SIZE  CB.HNDL, %IDC_TYPELT,   xx-55, yy-10
      CONTROL SET LOC   CB.HNDL, %IDC_ADDBTN,   xx-45, 10
      CONTROL SET LOC   CB.HNDL, %IDC_EDITBTN,  xx-45, 40
      CONTROL SET LOC   CB.HNDL, %IDC_DELBTN,   xx-45, 70
      CONTROL SET LOC   CB.HNDL, %IDCANCEL,     xx-45, yy-25
      DIALOG REDRAW     CB.HNDL
    CASE %WM_COMMAND
      ' Process control notifications
      IF CBCTLMSG <> %BN_CLICKED AND CBCTLMSG <> 1 THEN
        EXIT FUNCTION
      END IF
      SELECT CASE AS LONG CBCTL
        CASE %IDC_ADDBTN    '增加按钮
          IF InputDlg(CB.HNDL,"请输入分类名",tmpStr,"新建分类","")<=0 THEN
            EXIT FUNCTION
          END IF
          tmpStr=TRIM$(tmpStr)
          IF tmpStr="" THEN
            InfoBox CB.HNDL,"分类名不能为空"
            EXIT FUNCTION
          END IF
          AddType CB.HNDL,tmpStr

          rsStr=GetAllType(CB.HNDL)
          LISTBOX RESET CB.HNDL,%IDC_TYPELT
          FOR i=1 TO PARSECOUNT(rsStr,$CRLF)
            LISTBOX ADD CB.HNDL,%IDC_TYPELT,PARSE$(rsStr,$CRLF,i)
          NEXT i
        CASE %IDC_EDITBTN   '编辑按钮
          LISTBOX GET TEXT CB.HNDL,%IDC_TYPELT TO tmpStr
          tmpStr=TRIM$(tmpStr)
          IF tmpStr="" THEN
            InfoBox CB.HNDL,"请选择要编辑的分类名"
            EXIT FUNCTION
          END IF
          currTypeName=PARSE$(tmpStr,",",2)
          IF InputDlg(CB.HNDL,"请输入分类"+$DQ+currTypeName+$DQ+"的新分类名",tmpStr,"编辑分类",currTypeName)<=0 THEN
            EXIT FUNCTION
          END IF
          tmpStr=TRIM$(tmpStr)
          IF tmpStr="" THEN
            InfoBox CB.HNDL,"新分类名不能为空"
            EXIT FUNCTION
          END IF
          IF tmpStr=currTypeName THEN
            InfoBox CB.HNDL,"新分类名与当前分类名一致，不需要修改"
            EXIT FUNCTION
          END IF
          UpdateType CB.HNDL,tmpStr

          rsStr=GetAllType(CB.HNDL)
          LISTBOX RESET CB.HNDL,%IDC_TYPELT
          FOR i=1 TO PARSECOUNT(rsStr,$CRLF)
            LISTBOX ADD CB.HNDL,%IDC_TYPELT,PARSE$(rsStr,$CRLF,i)
          NEXT i
        CASE %IDC_DELBTN    '删除按钮
          LISTBOX GET TEXT CB.HNDL,%IDC_TYPELT TO tmpStr
          tmpStr=TRIM$(tmpStr)
          IF tmpStr="" THEN
            InfoBox CB.HNDL,"请选择要删除的分类名"
            EXIT FUNCTION
          END IF
          IF VAL(PARSE$(tmpStr,3))>0 THEN
            InfoBox CB.HNDL,"该分类已被使用，不能删除"
            EXIT FUNCTION
          END IF
          DeleteType CB.HNDL

          rsStr=GetAllType(CB.HNDL)
          LISTBOX RESET CB.HNDL,%IDC_TYPELT
          FOR i=1 TO PARSECOUNT(rsStr,$CRLF)
            LISTBOX ADD CB.HNDL,%IDC_TYPELT,PARSE$(rsStr,$CRLF,i)
          NEXT i
        CASE %IDCANCEL
          DIALOG END CB.HNDL,0
      END SELECT
  END SELECT
END FUNCTION
'------------------------------------------------------------------------------
FUNCTION InfoBox(BYVAL hParent AS DWORD,BYVAL promptStr AS STRING,OPT styleLng AS LONG,titleStr AS STRING)AS LONG
  LOCAL hasPic  AS LONG
  LOCAL rsLng   AS LONG
  LOCAL hInst   AS LONG
  LOCAL x,y     AS LONG
  LOCAL sLng    AS LONG
  LOCAL tStr    AS ASCIIZ * 100
  LOCAL pStr    AS ASCIIZ * 500

  pStr=promptStr

  IF ISMISSING(styleLng) THEN
    sLng=%MB_OK
  ELSE
    sLng=styleLng
  END IF
  IF ISMISSING(titleStr) THEN
    tStr=msgTitle
  ELSE
    tStr=titleStr
  END IF

  MessageBox hParent, pStr, tStr, sLng
END FUNCTION
'------------------------------------------------------------------------------
FUNCTION InputDlg(BYVAL hParent AS DWORD,BYVAL promptStr AS STRING,BYREF rsStr AS STRING,OPT titleStr AS STRING,_
          defaultStr AS STRING,xx AS LONG,yy AS LONG)AS LONG
  LOCAL hDlg        AS DWORD
  LOCAL hCtl        AS DWORD
  LOCAL ttStr       AS STRING
  LOCAL orgiStr     AS STRING
  LOCAL xpos        AS LONG
  LOCAL ypos        AS LONG
  LOCAL desktopw    AS LONG
  LOCAL desktoph    AS LONG
  LOCAL rsLng       AS LONG
  LOCAL dwStyle     AS LONG
  LOCAL pX,pY,pW,pH AS LONG
  LOCAL tmpLng      AS LONG
  LOCAL rc          AS RECT
  LOCAL pt          AS POINTAPI

  ttStr="提示"
  orgiStr=""

  DESKTOP GET CLIENT TO desktopw,desktoph
  IF ISFALSE ISMISSING(titleStr) THEN
    ttStr=titleStr
  END IF
  IF ISFALSE ISMISSING(defaultStr) THEN
    orgiStr=defaultStr
  END IF
  IF ISFALSE ISMISSING(xx) THEN
    xpos=xx
  END IF
  IF ISFALSE ISMISSING(yy) THEN
    ypos=yy
  END IF
  DIALOG NEW hParent, ttStr,xpos,ypos , 200, 123, _
              %WS_POPUP OR %WS_CAPTION OR %WS_THICKFRAME,%WS_EX_TOOLWINDOW TO hDlg
  IF ISMISSING(xx) OR ISMISSING(yy) THEN
    IF hParent=0 THEN
      dwStyle=GetWindowLong(hDlg,%GWL_STYLE)
      SetWindowLong(hDlg,%GWL_STYLE,dwStyle OR %DS_CENTER)
    ELSE
      CenterDialog(hParent,hDlg,200,123)
    END IF
  END IF
  DIALOG GET LOC hDlg TO pX,pY
  DIALOG SET USER hDlg,1,VARPTR(rsStr)
  DIALOG SEND hDlg, %WM_SETICON, %ICON_SMALL, LoadIcon(%NULL, BYVAL %IDI_APPLICATION)
  DIALOG SEND hDlg, %WM_SETICON, %ICON_BIG, LoadIcon(%NULL, BYVAL %IDI_APPLICATION)
  CONTROL ADD TEXTBOX, hDlg, %IDC_EDIT101,orgiStr, 5, 101, 187, 15, _
              %WS_CHILD OR %WS_VISIBLE OR %WS_TABSTOP OR %ES_AUTOHSCROLL, _
              %WS_EX_CLIENTEDGE OR %WS_EX_NOPARENTNOTIFY
  CONTROL ADD BUTTON, hDlg, %IDOK,        "确定", 139, 6, 53, 12, _
              %WS_CHILD OR %WS_VISIBLE OR %WS_TABSTOP, _
              %WS_EX_NOPARENTNOTIFY
  CONTROL ADD BUTTON, hDlg, %IDCANCEL,    "取消", 139, 23, 53, 12, _
              %WS_CHILD OR %WS_VISIBLE OR %WS_TABSTOP, _
              %WS_EX_NOPARENTNOTIFY
  CONTROL ADD LABEL, hDlg, %IDC_LABEL102, promptStr, 7, 6, 127, 89, _
              %WS_CHILD OR %WS_VISIBLE, _
              %WS_EX_NOPARENTNOTIFY
  DIALOG SHOW MODAL hDlg, CALL InputDlgProc TO rsLng
  IF rsLng=0 THEN
    rsStr=""
    FUNCTION=0
  ELSE
    FUNCTION=1
  END IF
END FUNCTION
CALLBACK FUNCTION InputDlgProc()AS LONG
  LOCAL tmpStr  AS STRING
  LOCAL tmpLng  AS LONG
  LOCAL pRsStr  AS STRING PTR
  LOCAL xx,yy   AS LONG

  SELECT CASE CB.MSG
    CASE %WM_SIZE
      IF CBWPARAM = %SIZE_MINIMIZED THEN EXIT FUNCTION
      DIALOG GET CLIENT CB.HNDL TO xx,yy
      IF xx<120 THEN
        xx=120
        DIALOG SET CLIENT CB.HNDL,xx,yy
      END IF
      IF yy<120 THEN
        yy=120
        DIALOG SET CLIENT CB.HNDL,xx,yy
      END IF
      CONTROL SET SIZE  CB.HNDL,  %IDC_EDIT101,   xx-15, 15
      CONTROL SET LOC   CB.HNDL,  %IDC_EDIT101,   5, yy-25
      CONTROL SET SIZE  CB.HNDL,  %IDC_LABEL102,  xx-70, yy-36
      CONTROL SET LOC   CB.HNDL,  %IDOK,          xx-58, 10
      CONTROL SET LOC   CB.HNDL,  %IDCANCEL,      xx-58, 30
      DIALOG REDRAW     CB.HNDL
    CASE %WM_COMMAND
      'Messages from controls and menu items are handled here.
      '-------------------------------------------------------
      IF CB.CTLMSG <> %BN_CLICKED THEN EXIT FUNCTION
      SELECT CASE CBCTL
       CASE %IDCANCEL
         DIALOG END CBHNDL, 0
       CASE %IDOK
         CONTROL GET TEXT CB.HNDL,%IDC_EDIT101 TO tmpStr
         DIALOG GET USER CB.HNDL,1 TO tmpLng
         pRsStr=tmpLng
         @pRsStr=tmpStr
         'MSGBOX tmpStr
         DIALOG END CB.HNDL,1
      END SELECT
  END SELECT
END FUNCTION
'------------------------------------------------------------------------------
FUNCTION CenterDialog(BYVAL hParent AS DWORD,BYVAL hDlg AS DWORD,BYVAL dlgWidth AS LONG,BYVAL dlgHeight AS LONG)AS LONG
  LOCAL x,y   AS LONG
  LOCAL rc    AS RECT
  LOCAL pt    AS POINTAPI
  GetWindowRect hParent,rc
  DIALOG PIXELS hDlg,rc.nLeft,rc.nTop TO UNITS rc.nLeft,rc.nTop
  DIALOG PIXELS hDlg,rc.nRight,rc.nBottom TO UNITS rc.nRight,rc.nBottom

  IF ISFALSE IsWindow(hParent) THEN
    hParent=%HWND_DESKTOP
  ELSE
    IF ISFALSE iswindowVisible(hParent) THEN
      hParent=%HWND_DESKTOP
    END IF
  END IF

  IF hParent<>%HWND_DESKTOP THEN
    'DIALOG SET LOC hDlg, rc.nLeft+(rc.nRight-rc.nLeft-dlgWidth)\2, rc.nTop+(rc.nBottom-rc.nTop-dlgHeight)\2
    DIALOG SET LOC hDlg, (rc.nRight-rc.nLeft-dlgWidth)\2, (rc.nBottom-rc.nTop-dlgHeight)\2
  END IF
END FUNCTION
'------------------------------------------------------------------------------
FUNCTION InitGlobalVars() AS LONG
  dbname="infos.db"
  pSize=20
  msgTitle="提示"
END FUNCTION
'------------------------------------------------------------------------------
FUNCTION GetAllTypeName(BYVAL hWnd AS DWORD) AS STRING
  LOCAL sqlStr  AS STRING
  LOCAL rsStr   AS STRING

  IF ISFALSE sqlOpen(dbname, hDB) THEN
    InfoBox hWnd,"unable to open db"
    GOTO HERE
  END IF
  IF ISFALSE sqlTableExist(hDB, "t_type") THEN
    sqlStr="CREATE TABLE IF NOT EXISTS t_type(id integer primary key autoincrement,name varchar(40))"
    IF ISFALSE sqlExe(hDB, sqlStr) THEN
      InfoBox hWnd,sqlErrMsg(hDB)
      GOTO HERE
    END IF
    FUNCTION=""
    EXIT FUNCTION
  END IF
  rsStr=ExecuteQuery(hWnd,hDB,"select name from t_type order by id")
  REPLACE $CRLF WITH "," IN rsStr
  REPLACE ", ," WITH "," IN rsStr
  FUNCTION=rsStr
HERE:
  sqlClose hDB
END FUNCTION
'------------------------------------------------------------------------------
FUNCTION ResetType(BYVAL hWnd AS DWORD)AS LONG
  LOCAL selType   AS STRING
  LOCAL sqlStr    AS STRING
  LOCAL rsStr     AS STRING
  LOCAL typeArr() AS STRING
  LOCAL i         AS LONG

  CONTROL GET TEXT hWnd,%IDC_TYPECB TO selType
  selType=TRIM$(selType)
  COMBOBOX RESET hWnd,%IDC_TYPECB
  rsStr="全部,"+GetAllTypeName(hWnd)
  REDIM typeArr(PARSECOUNT(rsStr,",")-1)
  PARSE rsStr,typeArr(),","
  FOR i=0 TO UBOUND(typeArr())
    COMBOBOX ADD hWnd,%IDC_TYPECB,TRIM$(typeArr(i))
  NEXT i
  CONTROL SET TEXT hWnd,%IDC_TYPECB,selType
END FUNCTION
'------------------------------------------------------------------------------
FUNCTION DoQuery(BYVAL hWnd AS DWORD)AS LONG
  LOCAL sqlStr    AS STRING
  LOCAL whereStr  AS STRING
  LOCAL rsStr     AS STRING
  LOCAL tmpStr    AS STRING
  LOCAL findStr   AS STRING
  LOCAL typeStr   AS STRING
  LOCAL pageStr   AS STRING
  LOCAL recCount  AS LONG
  LOCAL pageCount AS LONG
  LOCAL i         AS LONG

  IF ISFALSE sqlOpen(dbname, hDB) THEN
    InfoBox hWnd,"unable to open db"
    GOTO HERE
  END IF
  IF ISFALSE sqlTableExist(hDB, "t_info") THEN
    sqlStr="CREATE TABLE IF NOT EXISTS t_info(id integer primary key autoincrement," _
            +"title varchar(40),type varchar(40),content varchar(500))"
    IF ISFALSE sqlExe(hDB, sqlStr) THEN
      InfoBox hWnd,sqlErrMsg(hDB)
      GOTO HERE
    END IF
    CONTROL SET TEXT hWnd,%IDC_PAGEINFOLB,"无记录"
    FUNCTION=0
    EXIT FUNCTION
  END IF

  CONTROL GET TEXT hWnd,%IDC_CURRPAGETB TO pageStr
  pageStr=TRIM$(pageStr)
  IF pageStr="" THEN
    pageStr="1"
  END IF
  LISTBOX RESET hWnd,%IDC_INFOLT
  sqlStr="select count(*) from t_info"

  CONTROL GET TEXT hWnd,%IDC_TYPECB TO typeStr
  typeStr=TRIM$(typeStr)
  IF typeStr="" THEN
    typeStr="全部"
  END IF
  IF typeStr<>"全部" THEN
    whereStr=" type='" & typeStr & "'"
  END IF

  CONTROL GET TEXT hWnd,%IDC_SEARCHTB TO findStr
  findStr=TRIM$(findStr)

  IF findStr<>"" THEN
    IF whereStr<>"" THEN
      whereStr = whereStr + " and"
    END IF
    whereStr=whereStr + " (title like '%" + findStr + "%' or content like '%" + findStr + "%')"
  END IF
  'msgbox "test1 whereStr=" & whereStr
  IF whereStr="" THEN
    rsStr=ExecuteQuery(hWnd,hDB,sqlStr)
  ELSE
    rsStr=ExecuteQuery(hWnd,hDB,sqlStr & " where" & whereStr)
  END IF
  IF RIGHT$(rsStr,1)="," THEN
    rsStr=MID$(rsStr,1,LEN(rsStr)-1)
  END IF
  recCount=VAL(rsStr)
  IF recCount<=0 THEN
    CONTROL SET TEXT hWnd,%IDC_PAGEINFOLB,"无记录"
    FUNCTION=0
    EXIT FUNCTION
  END IF
  IF recCount<=pSize THEN
    pageStr="1"
    pageCount=1
  ELSE
    pageCount=recCount\pSize
    IF (recCount MOD pSize)>0 THEN
      INCR pageCount
    END IF
    'MSGBOX "test pageCount="+STR$(pageCount)
  END IF
  IF pageCount<VAL(pageStr) THEN
    pageStr=FORMAT$(pageCount)
  END IF

  CONTROL SET TEXT hWnd,%IDC_CURRPAGETB,pageStr
  CONTROL SET TEXT hWnd,%IDC_PAGEINFOLB,"共" & STR$(recCount) & " 条  第 " & pageStr & " 页 共" _
      & STR$(pageCount) & " 页"

  sqlStr="select * from t_info"
  IF whereStr<>"" THEN
    sqlStr=sqlStr & " where " & whereStr
  END IF
  sqlStr=sqlStr & " limit " & STR$(pSize) & " offset " & STR$((VAL(pageStr)-1)*pSize)
  rsStr=ExecuteQuery(hWnd,hDB,sqlStr)

  FOR i=1 TO PARSECOUNT(rsStr,$CRLF)
    LISTBOX ADD hWnd,%IDC_INFOLT,PARSE$(rsStr,$CRLF,i)
  NEXT i
  FUNCTION=recCount
HERE:
  sqlClose hDB
END FUNCTION
'------------------------------------------------------------------------------
FUNCTION DoAdd(BYVAL hWnd AS DWORD)AS LONG
  LOCAL sqlStr      AS STRING
  LOCAL whereStr    AS STRING
  LOCAL rsStr       AS STRING
  LOCAL tmpStr      AS STRING
  LOCAL titleStr    AS STRING
  LOCAL typeStr     AS STRING
  LOCAL contentStr  AS STRING
  LOCAL i           AS LONG

  IF ISFALSE sqlOpen(dbname, hDB) THEN
    InfoBox hWnd,"unable to open db"
    GOTO HERE
  END IF
  IF ISFALSE sqlTableExist(hDB, "t_info") THEN
    sqlStr="CREATE TABLE IF NOT EXISTS t_info(id integer primary key autoincrement," _
            +"title varchar(40),type varchar(40),content varchar(500))"
    IF ISFALSE sqlExe(hDB, sqlStr) THEN
      InfoBox hWnd,sqlErrMsg(hDB)
      FUNCTION=0
      GOTO HERE
    END IF
  END IF
  CONTROL GET TEXT hWnd,%IDC_TYPECB TO typeStr
  typeStr=TRIM$(typeStr)
  IF ISFALSE sqlTableExist(hDB, "t_type") THEN
    sqlStr="CREATE TABLE IF NOT EXISTS t_type(id integer primary key autoincrement,name varchar(40))"
    IF ISFALSE sqlExe(hDB, sqlStr) THEN
      InfoBox hWnd,sqlErrMsg(hDB)
      FUNCTION=0
      GOTO HERE
    END IF
    sqlExe(hDB, "insert into t_type (name)values('"+typeStr+"')")
  ELSE
    rsStr=ExecuteQuery(hWnd,hDB,"select name from t_type where name='"+typeStr+"'")
    IF rsStr="" THEN
      sqlExe(hDB, "insert into t_type (name)values('"+typeStr+"')")
    END IF
  END IF
  CONTROL GET TEXT hWnd,%IDC_TITLETB TO titleStr
  titleStr=TRIM$(titleStr)
  REPLACE $CRLF WITH "\n" IN titleStr
  REPLACE $CR WITH "\n" IN titleStr
  REPLACE $LF WITH "\n" IN titleStr
  WHILE INSTR(titleStr,"\n\n")>0
    REPLACE "\n\n" WITH "\n" IN titleStr
  WEND
  REPLACE "'" WITH "''" IN titleStr
  sqlStr="select * from t_info where title='"+titleStr+"'"
  rsStr=ExecuteQuery(hWnd,hDB,sqlStr)
  IF rsStr<>"" THEN
    InfoBox hWnd,"已存在"
    FUNCTION=0
    GOTO HERE
  END IF
  CONTROL GET TEXT hWnd,%IDC_CONTENTTB TO contentStr
  contentStr=TRIM$(contentStr)
  REPLACE $CRLF WITH "\n" IN contentStr
  REPLACE $CR WITH "\n" IN contentStr
  REPLACE $LF WITH "\n" IN contentStr
  WHILE INSTR(contentStr,"\n\n")>0
    REPLACE "\n\n" WITH "\n" IN contentStr
  WEND
  REPLACE "'" WITH "''" IN contentStr
  sqlStr="insert into t_info (type,title,content)values('"+typeStr+"','"+titleStr+"','"+contentStr+"')"
  sqlExe(hDB,sqlStr)
  InfoBox hWnd,"创建成功"
  FUNCTION=1
HERE:
  sqlClose hDB
END FUNCTION
'------------------------------------------------------------------------------
FUNCTION DoUpdate(BYVAL hWnd AS DWORD,BYVAL idStr AS STRING)AS LONG
  LOCAL sqlStr      AS STRING
  LOCAL whereStr    AS STRING
  LOCAL rsStr       AS STRING
  LOCAL tmpStr      AS STRING
  LOCAL titleStr    AS STRING
  LOCAL typeStr     AS STRING
  LOCAL contentStr  AS STRING
  LOCAL i           AS LONG

  IF ISFALSE sqlOpen(dbname, hDB) THEN
    InfoBox hWnd,"unable to open db"
    GOTO HERE
  END IF
  IF ISFALSE sqlTableExist(hDB, "t_info") THEN
    sqlStr="CREATE TABLE IF NOT EXISTS t_info(id integer primary key autoincrement," _
            +"title varchar(40),type varchar(40),content varchar(500))"
    IF ISFALSE sqlExe(hDB, sqlStr) THEN
      InfoBox hWnd,sqlErrMsg(hDB)
      FUNCTION=0
      GOTO HERE
    END IF
  END IF
  CONTROL GET TEXT hWnd,%IDC_TYPECB TO typeStr
  typeStr=TRIM$(typeStr)
  IF ISFALSE sqlTableExist(hDB, "t_type") THEN
    sqlStr="CREATE TABLE IF NOT EXISTS t_type(id integer primary key autoincrement,name varchar(40))"
    IF ISFALSE sqlExe(hDB, sqlStr) THEN
      InfoBox hWnd,sqlErrMsg(hDB)
      FUNCTION=0
      GOTO HERE
    END IF
    sqlExe(hDB, "insert into t_type (name)values('"+typeStr+"')")
  ELSE
    rsStr=ExecuteQuery(hWnd,hDB,"select name from t_type where name='"+typeStr+"'")
    IF rsStr="" THEN
      sqlExe(hDB, "insert into t_type (name)values('"+typeStr+"')")
    END IF
  END IF
  CONTROL GET TEXT hWnd,%IDC_TITLETB TO titleStr
  titleStr=TRIM$(titleStr)
  REPLACE $CRLF WITH "\n" IN titleStr
  REPLACE $CR WITH "\n" IN titleStr
  REPLACE $LF WITH "\n" IN titleStr
  WHILE INSTR(titleStr,"\n\n")>0
    REPLACE "\n\n" WITH "\n" IN titleStr
  WEND
  REPLACE "'" WITH "''" IN titleStr

  CONTROL GET TEXT hWnd,%IDC_CONTENTTB TO contentStr
  contentStr=TRIM$(contentStr)
  REPLACE $CRLF WITH "\n" IN contentStr
  REPLACE $CR WITH "\n" IN contentStr
  REPLACE $LF WITH "\n" IN contentStr
  WHILE INSTR(contentStr,"\n\n")>0
    REPLACE "\n\n" WITH "\n" IN contentStr
  WEND
  REPLACE "'" WITH "''" IN contentStr
  sqlStr="update t_info set type='"+typeStr+"',title='"+titleStr+"',content='"+contentStr+"' where id=" + idStr
  sqlExe(hDB,sqlStr)
  InfoBox hWnd,"更新成功"
  FUNCTION=1
HERE:
  sqlClose hDB
END FUNCTION
'------------------------------------------------------------------------------
FUNCTION DoDelete(BYVAL hWnd AS DWORD)AS LONG
  LOCAL sqlStr      AS STRING
  LOCAL whereStr    AS STRING
  LOCAL rsStr       AS STRING
  LOCAL tmpStr      AS STRING
  LOCAL idsStr      AS STRING
  LOCAL typeStr     AS STRING
  LOCAL contentStr  AS STRING
  LOCAL i           AS LONG
  LOCAL tmpLng      AS LONG
  LOCAL tmpLng1     AS LONG

  IF ISFALSE sqlOpen(dbname, hDB) THEN
    InfoBox hWnd,"unable to open db"
    GOTO HERE
  END IF
  IF ISFALSE sqlTableExist(hDB, "t_info") THEN
    sqlStr="CREATE TABLE IF NOT EXISTS t_info(id integer primary key autoincrement," _
            +"title varchar(40),type varchar(40),content varchar(500))"
    IF ISFALSE sqlExe(hDB, sqlStr) THEN
      InfoBox hWnd,sqlErrMsg(hDB)
      FUNCTION=0
      GOTO HERE
    END IF
    FUNCTION = 0
    GOTO HERE
  END IF
  LISTBOX GET COUNT hWnd,%IDC_INFOLT TO tmpLng
  FOR i=1 TO tmpLng
    LISTBOX GET STATE hWnd,%IDC_INFOLT,i TO tmpLng1
    IF ISTRUE tmpLng1 THEN
      LISTBOX GET TEXT hWnd,%IDC_INFOLT,i TO tmpStr
      idsStr=idsStr+PARSE$(tmpStr,",",1)+","
    END IF
  NEXT i
  IF RIGHT$(idsStr,1)="," THEN
    idsStr=MID$(idsStr,1,LEN(idsStr)-1)
  END IF
  sqlStr="delete from t_info where id in("+idsStr+")"
  sqlExe(hDB,sqlStr)
  FUNCTION=PARSECOUNT(idsStr,",")
HERE:
  sqlClose hDB
END FUNCTION
'------------------------------------------------------------------------------
FUNCTION ExecuteQuery(BYVAL hWnd AS DWORD,hDB AS LONG,BYVAL sqlStr AS STRING)AS STRING
  LOCAL s AS STRING
  LOCAL i AS LONG
  sqlRecSetNew(rs, hDB)
  IF ISFALSE sqlSelect( rs, sqlStr ) THEN
    InfoBox hWnd,sqlErrMsg(hDB)
    EXIT FUNCTION
  END IF
  sqlMoveFirst rs
  WHILE ISFALSE rs.IsEof
    FOR i=1 TO rs.ColCount
        s = s + sqlGetAt(rs,i) + ", "
    NEXT i
    s = s + $CRLF
    sqlMoveNext rs
  WEND
  IF RIGHT$(s,2)=$CRLF THEN
    s = MID$(s,1,LEN(s)-2)
  END IF
  FUNCTION=s
END FUNCTION
'------------------------------------------------------------------------------
FUNCTION GetAllType(BYVAL hWnd AS DWORD) AS STRING
  LOCAL sqlStr  AS STRING
  LOCAL rsStr   AS STRING

  IF ISFALSE sqlOpen(dbname, hDB) THEN
    InfoBox hWnd,"unable to open db"
    GOTO HERE
  END IF
  IF ISFALSE sqlTableExist(hDB, "t_type") THEN
    sqlStr="CREATE TABLE IF NOT EXISTS t_type(id integer primary key autoincrement,name varchar(40))"
    IF ISFALSE sqlExe(hDB, sqlStr) THEN
      InfoBox hWnd,sqlErrMsg(hDB)
      GOTO HERE
    END IF
    FUNCTION=""
    EXIT FUNCTION
  END IF
  rsStr=ExecuteQuery(hWnd,hDB,"select t1.id,t1.name,count(t2.id) as InfoCount from t_type t1 " _
        & "left join t_info t2 on t2.type=t1.name group by t1.id,t1.name")
  'REPLACE $CRLF WITH "," IN rsStr
  REPLACE ", ," WITH "," IN rsStr
  FUNCTION=rsStr
HERE:
  sqlClose hDB
END FUNCTION
'------------------------------------------------------------------------------
FUNCTION AddType(BYVAL hWnd AS DWORD,BYVAL typeName AS STRING)AS LONG
  LOCAL sqlStr  AS STRING
  LOCAL rsStr   AS STRING

  IF ISFALSE sqlOpen(dbname, hDB) THEN
    InfoBox hWnd,"unable to open db"
    GOTO HERE
  END IF
  IF ISFALSE sqlTableExist(hDB, "t_type") THEN
    sqlStr="CREATE TABLE IF NOT EXISTS t_type(id integer primary key autoincrement,name varchar(40))"
    IF ISFALSE sqlExe(hDB, sqlStr) THEN
      InfoBox hWnd,sqlErrMsg(hDB)
      GOTO HERE
    END IF
    FUNCTION=0
    sqlStr="insert into t_type (name)values('" + typeName + "')"
    sqlExe(hDB,sqlStr)
    InfoBox hWnd,"已成功增加分类"
    GOTO HERE
  END IF
  rsStr=ExecuteQuery(hWnd,hDB,"select * from t_type where name='"+typeName+"'")
  IF rsStr<>"" THEN
    InfoBox hWnd,"已存在该分类，请不要重复添加"
    FUNCTION=0
    GOTO HERE
  END IF
  sqlExe(hDB,"insert into t_type (name)values('" + typeName + "')")
  InfoBox hWnd,"已成功增加分类"
  FUNCTION=1
HERE:
  sqlClose hDB
END FUNCTION
'------------------------------------------------------------------------------
FUNCTION UpdateType(BYVAL hWnd AS DWORD,BYVAL newTypeName AS STRING)AS LONG
  LOCAL sqlStr  AS STRING
  LOCAL rsStr   AS STRING
  LOCAL tmpStr  AS STRING

  IF ISFALSE sqlOpen(dbname, hDB) THEN
    InfoBox hWnd,"unable to open db"
    GOTO HERE
  END IF
  IF ISFALSE sqlTableExist(hDB, "t_type") THEN
    sqlStr="CREATE TABLE IF NOT EXISTS t_type(id integer primary key autoincrement,name varchar(40))"
    IF ISFALSE sqlExe(hDB, sqlStr) THEN
      InfoBox hWnd,sqlErrMsg(hDB)
      GOTO HERE
    END IF
    FUNCTION=0
    sqlStr="insert into t_type (name)values('" + newTypeName + "')"
    sqlExe(hDB,sqlStr)
    InfoBox hWnd,"已成功增加分类"
    GOTO HERE
  END IF
  rsStr=ExecuteQuery(hWnd,hDB,"select * from t_type where name='"+newTypeName+"'")
  IF rsStr<>"" THEN
    InfoBox hWnd,"已存在该分类，不能修改分类"
    FUNCTION=0
    GOTO HERE
  END IF
  LISTBOX GET TEXT hWnd,%IDC_TYPELT TO tmpStr
  sqlExe(hDB,"update t_type set name='" + newTypeName + "' where id="+PARSE$(tmpStr,",",1))
  sqlExe(hDB,"update t_info set type='" + newTypeName + "' where type='"+TRIM$(PARSE$(tmpStr,",",2))+"'")
  InfoBox hWnd,"已成功修改分类"
  FUNCTION=1
HERE:
  sqlClose hDB
END FUNCTION
'------------------------------------------------------------------------------
FUNCTION DeleteType(BYVAL hWnd AS DWORD)AS LONG
  LOCAL sqlStr    AS STRING
  LOCAL rsStr     AS STRING
  LOCAL tmpStr    AS STRING
  LOCAL typeCount AS LONG
  LOCAL tmpLng    AS LONG
  LOCAL i         AS LONG
  LOCAL idsStr    AS STRING

  IF ISFALSE sqlOpen(dbname, hDB) THEN
    InfoBox hWnd,"unable to open db"
    GOTO HERE
  END IF
  IF ISFALSE sqlTableExist(hDB, "t_type") THEN
    sqlStr="CREATE TABLE IF NOT EXISTS t_type(id integer primary key autoincrement,name varchar(40))"
    IF ISFALSE sqlExe(hDB, sqlStr) THEN
      InfoBox hWnd,sqlErrMsg(hDB)
      GOTO HERE
    END IF
    FUNCTION=0
    GOTO HERE
  END IF
  LISTBOX GET COUNT hWnd,%IDC_TYPELT TO typeCount
  FOR i=1 TO typeCount
    LISTBOX GET STATE hWnd,%IDC_TYPELT,i TO tmpLng
    IF ISTRUE tmpLng THEN
      LISTBOX GET TEXT hWnd,%IDC_TYPELT,i TO tmpStr
      IF VAL(PARSE$(tmpStr,",",3))>0 THEN
        ITERATE FOR
      END IF
      idsStr=idsStr & PARSE$(tmpStr,",",1) & ","
    END IF
  NEXT i
  IF RIGHT$(idsStr,1)="," THEN
    idsStr=MID$(idsStr,1,LEN(idsStr)-1)
  END IF
  IF idsStr<>"" THEN
    sqlExe(hDB,"delete from t_type where id in("+idsStr+")")
  END IF
  FUNCTION=1
HERE:
  sqlClose hDB
END FUNCTION
'------------------------------------------------------------------------------
