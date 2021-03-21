/**************************************************
    文件名  ：AutoInput
    创建日期：2020年7月2日
    作者    ：李宇
    功能    ：自动填写数据到光标处
    支持功能：
        1、读取文本文件中存储的数据和按键序列
        2、读取Excel表格中存储的数据
        3、导入导出按键序列
        4、文本数据支持重复输入和输入前后等待的时间
    更新历史：
        2020.07.02  初版做成
        2020.11.10  整合读取文本和读取Excel表格功能
        2020.11.11  程序界面和功能优化
*/

#NoEnv      ; Recommended for performance and compatibility with future AutoHotkey releases.
;#Warn       ; Enable warnings to assist with detecting common errors.
SendMode Input                  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%     ; Ensures a consistent starting directory.

Gui, +AlwaysOnTop -MaximizeBox
Gui, Font, S10, 微软雅黑
Gui, Margin , 5, 5

Menu, FileMenu, Add, 导入文本文件`tCtrl+T, MenuHandler
Menu, FileMenu, Add, 导入按键序列文件`tCtrl+I, MenuHandler
Menu, FileMenu, Add, 导出按键序列文件`tCtrl+E, MenuHandler
Menu, FileMenu, Add
Menu, FileMenu, Add, 退出(&X), GuiClose
Menu, SettingMenu, Add, 插入序列后自动读取下一行`tCtrl+P, MenuHandler
Menu, SettingMenu, Add, 插入序列后自动读取上一行`tCtrl+N, MenuHandler
Menu, ViewMenu, Add, 显示数据信息`tCtrl+S, MenuHandler
Menu, HelpMenu, Add, 关于, MenuHandler
Menu, MenuBar, Add, 文件(&F), :FileMenu
Menu, MenuBar, Add, 设置(&S), :SettingMenu
Menu, MenuBar, Add, 视图(&V), :ViewMenu
Menu, MenuBar, Add, 帮助(&H), :HelpMenu
Menu, Tray, Tip , 自动读取工作表行并录入光标处
Gui, Menu, MenuBar
Gui, Add, StatusBar
Gui, Add, ListView, R8 W400 Grid NoSortHdr -LV0x10 GSelRow, 数据项|插入按键
Gui, Add, Tab3, W400 Vtabsel AltSubmit, 文本数据|Excel数据|按键设置

Gui, Tab, 文本数据
Gui, Add, CheckBox, Section Vbol_repeat, 重复执行
Gui, Add, Edit,     X+5 W30 Vrepeat_count R1 Center Number Limit3, 30
Gui, Add, Text,     X+5, 次
Gui, Add, Text,     XS, 前置数据
Gui, Add, Edit,     X+5 W320 Vpre_data R1, {LButton}
Gui, Add, Text,     XS, 执行前后等待时间（单位：秒）
Gui, Add, Edit,     X+5 W30 Vtime1 R1 Center Number Limit2, 2
Gui, Add, Edit,     X+5 W30 Vtime2 R1 Center Number Limit2, 8
Gui, Add, Progress, XS W385 H20 Vprogress Range0-100 -Smooth, 0

Gui, Tab, Excel数据
Gui, Add, GroupBox, W390 H60 CRed Section, 激活Excel工作表，输入数据的起始行列和项数
Gui, Add, Text,     XP5 YP25 Center, 行列项
Gui, Add, Edit,     X+5 W60 Vtxt_row R1 Center Number Limit5
Gui, Add, UpDown,   Vexcel_row 0x80 Range1-99999, 2
Gui, Add, Edit,     X+0 W60 Vtxt_col R1 Center Number Limit5
Gui, Add, UpDown,   Vexcel_col 0x80 Range1-99999, 1
Gui, Add, Edit,     X+0 W50 Vtxt_count R1 Center Number Limit3
Gui, Add, UpDown,   Vdata_count 0x80 Range1-999, 10
Gui, Add, Button,   X+5 W50 GReadExcelFile, 读取
Gui, Add, Button,   X+5 W50 GReadExcelFile, 上一行
Gui, Add, Button,   X+5 W50 GReadExcelFile, 下一行

Gui, Tab, 按键设置
Gui, Add, GroupBox, W390 H95 CRed, 双击数据行后，插入或清除按键
Gui, Add, Button,   XP5 W50 YP25 GUpdateKey, Tab
Gui, Add, Button,   X+5 W50 GUpdateKey, 空格
Gui, Add, Button,   X+5 W50 GUpdateKey, 回车
Gui, Add, Button,   X+5 W50 GUpdateKey, Home
Gui, Add, Button,   X+5 W50 GUpdateKey, End
Gui, Add, Button,   X+5 W50 GUpdateKey, PgUp
Gui, Add, Button,   X+5 W50 GUpdateKey, PgDn
Gui, Add, Button,   X15 Y+5 W50 GUpdateKey, ↑
Gui, Add, Button,   X+5 W50 GUpdateKey, ↓
Gui, Add, Button,   X+5 W50 GUpdateKey, ←
Gui, Add, Button,   X+5 W50 GUpdateKey, →
Gui, Add, Button,   X+5 W50 GUpdateKey, 清除
Gui, Add, Button,   X+5 GClearAllKey, 清除所有

Gui, Tab            ; 随后的控件不属于任何选项卡控件的一部分。
Gui, Add, Edit,     W400 VData R5 ReadOnly Hidden, 这里将显示生成的插入序列，`n按【CTRL+1】组合键可将序列插入到当前光标处。
Gui, Add, Button,   GInsDataToText VBtnInsData Hidden, 打开记事本，并在光标处插入测试数据(&I)

Gui, Show, AutoSize NoActivate, 自动填写表单，按组合键CTRL+1执行

global sheet, excel_row, excel_col, data_count
global txt_row, txt_col, txt_count
global pre_data, arrData, arrKey, send_data, selindex
global flag_auto_read   ;0:不自动读取 1:自动读取下一行 2:自动读取上一行

InitData()

^1::
    Gui, Submit , NoHide
    Switch tabsel {
        Case 1:
            if(bol_repeat==1){
                SB_SetText("重复执行开始…")
                While(A_Index<=repeat_count){

                    Send , %pre_data%
                    Sleep, time1 * 1000
                    Send , %send_data%
                    Sleep, time2 * 1000
                    GuiControl,, progress, % A_Index*100/repeat_count
                    SB_SetText("重复执行进度 " A_Index "/" repeat_count)
                }
                SB_SetText("重复执行完成，共执行 " repeat_count " 次。")
            }Else{
                Send , %send_data%
            }
        Case 2:
            InsDataPro()
    }
    Return

InitData(){
    data_count:=0
    selindex:=1
    flag_auto_read:=0   ;0:不自动读取
    arrData:=Array()
    arrKey:=Array()

    filename := SubStr(A_ScriptName, 1, InStr(A_ScriptName, ".")) "txt"
    if(FileExist(filename)){
        ReadTextFile(filename)
        UpdateListView()
        UpdateSendData()
    }
    Return   
}

InsDataToText(){
    if(!WinExist("ahk_exe notepad.exe")){
        Run, notepad.exe
        Sleep, 500
    }
    WinActivate , ahk_class Notepad
    if(WinActive("ahk_class Notepad")){
        InsDataPro()
    }   
    Return
}

InsDataPro(){
    GuiControl, Disable, BtnInsData
    Send , %send_data%
    If(flag_auto_read == 1){
;        ReadNextData()
    }Else If(flag_auto_read == 2){
;        ReadPreData()
    }
    GuiControl, Enable, BtnInsData
    Return
}

MenuHandler(ItemName, ItemPos, MenuName){
    local v:=0
    Switch MenuName {
        Case "FileMenu":
            Switch ItemPos {
                Case 1:
                    ReadTextFile()
                    UpdateListView()
                    UpdateSendData()
                Case 2:
                    ReadKeyFile()
                    UpdateListView()
                    UpdateSendData()
                Case 3:
                    WriteKeyFile()
            }
        Case "SettingMenu":
            Menu, % MenuName, Uncheck, 1&
            Menu, % MenuName, Uncheck, 2&
            If(ItemPos==flag_auto_read){
                flag_auto_read:=0
            }Else{
                Menu, % MenuName, Check, % ItemPos "&"
                flag_auto_read:=ItemPos
            }
        Case "ViewMenu":
            Switch ItemPos {
                Case 1:
                    GuiControlGet, v, Visible , Data
                    Menu, % MenuName, % v?"Uncheck":"Check", % ItemPos "&"
                    GuiControl, % v?"Hide":"Show", Data
                    GuiControl, % v?"Hide":"Show", BtnInsData
                    Gui, Show, AutoSize
            }
        Case "HelpMenu":
            Switch ItemPos {
                Case 1:
                    MsgBox, 0x1000, 关于应用, 作者：李宇`n版本：0.9.20201110
            }
    }
    Return
}

ReadTextFile(filename:=""){
    If (filename = ""){
        FileSelectFile, filename, , , 导入文本文件, 文本文件(*.txt)
    }
    If(filename = ""){
        Return False
    }
    data_count:=0
    Loop, Read, %filename%
    {
        pos := InStr(A_LoopReadLine, "`t")
        if(pos > 0){
            arrData[A_Index] := Trim(SubStr(A_LoopReadLine, 1, pos - 1))
            arrKey[A_Index] := Trim(SubStr(A_LoopReadLine, pos))
        }else{
            arrData[A_Index] := Trim(A_LoopReadLine)
            arrKey[A_Index] := ""
        }
        data_count++
    }
    Return True
}

ReadKeyFile(){
    local filename
    FileSelectFile, filename, , , 导入按键序列文件, 按键序列文件(*.key)
    If(filename = ""){
        Return False
    }
    Loop, Read, %filename%
    {
        arrKey[A_Index] := A_LoopReadLine
    }
    Return True
}

WriteKeyFile(){
    local file, filename, datastr:=""
    FileSelectFile, filename, S16, , 导出按键序列文件, 按键序列文件(*.key)
    If(filename=""){
        Return False
    }
    file:=FileOpen(filename, "w", "UTF-8")
    If(file){
        Loop, % data_count {
            datastr := datastr arrKey[A_Index] "`r`n"
        }
        file.Write(datastr)
        file.Close()
    }Else{
        Return False
    }
    Return True
}

ReadExcelFile(CtrlHwnd, GuiEvent, EventInfo, ErrLevel:=""){
    Gui, Submit , NoHide
    GuiControlGet, btntext,, % CtrlHwnd, Value
    Switch btntext {
        Case "读取":
            If (txt_row<1 || txt_col<1 || txt_count<1){
                GuiControl, Text, txt_row, % excel_row
                GuiControl, Text, txt_col, % excel_col
                GuiControl, Text, txt_count, % data_count
            }
        Case "上一行":
            If (excel_row > 1){
                GuiControl, Text, excel_row, % --excel_row
            } 
        Case "下一行":
            If (excel_row < 99999){
                GuiControl, Text, excel_row, % ++excel_row
            }
    }
    If (ReadExcelData(excel_row, excel_col, data_count)){
        UpdateListView()
        UpdateSendData()
    }Else{
        MsgBox, % 0x10+0x1000+0x40000, 读取数据出错, 请首先打开Excel文件，并激活需读取的工作表。
    }
    Return
}

ReadExcelData(row, col, count){
    local cell
    if(!sheet){
        Try {
            sheet := ComObjActive("Excel.Application").ActiveSheet
        }Catch{
            Return False
        }
    }
    cell := sheet.Cells(row, col)
    Loop, % count {
        arrData[A_Index] := cell.Offset(0, A_Index-1).Value
    }
    Return True
}

UpdateListView(){
    LV_Delete()
    Loop, % data_count {
        If (arrKey.Length()<A_Index){
            arrKey[A_Index]:="{Tab}"
        }
        LV_Add("" , arrData[A_Index], arrKey[A_Index])
    }
    LV_ModifyCol()
    Return
}

UpdateSendData(){
    send_data := ""
    Loop, % data_count {
        send_data := send_data arrData[A_Index] arrKey[A_Index]
    }
    GuiControl, Text, Data, %send_data%
    Return
}

GuiClose(GuiHwnd) {
    ExitApp , 0
}

SelRow(CtrlHwnd, GuiEvent, EventInfo, ErrLevel:=""){
    if (GuiEvent = "DoubleClick"){
        selindex := EventInfo
    }
    Return
}

UpdateKey(CtrlHwnd, GuiEvent, EventInfo, ErrLevel:=""){
    GuiControlGet, btntext,, % CtrlHwnd, Value
    Switch btntext {
        Case "Tab":     key:="{Tab}"
        Case "空格":    key:="{Space}"
        Case "回车":    key:="{Enter}"
        Case "Home":    key:="{Home}"
        Case "End":     key:="{End}"
        Case "PgUp":    key:="{PgUp}"
        Case "PgDn":    key:="{PgDn}"
        Case "↑":       key:="{Up}"
        Case "↓":       key:="{Down}"
        Case "←":       key:="{Left}"
        Case "→":       key:="{Right}"
        Case "清除":
            arrKey[selindex]:=""
            key:=""
    }
    arrKey[selindex]:=arrKey[selindex] key
    LV_Modify(selindex, , arrData[selindex], arrKey[selindex])
    LV_ModifyCol()
    UpdateSendData()
    Return
}

ClearAllKey(CtrlHwnd, GuiEvent, EventInfo, ErrLevel:=""){
    Loop, % data_count {
        arrKey[A_Index] := ""
        LV_Modify(A_Index, , arrData[A_Index], "")
        send_data := send_data arrData[A_Index] arrKey[A_Index]
    }
    LV_ModifyCol()
    UpdateSendData()
    Return
}