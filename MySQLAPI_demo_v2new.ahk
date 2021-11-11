; ======================================================================================================================
; Demo of MySQLAPI class
;
; You must have access to a running MySQL server. This demo app will create a database and a table and present
; a simple GUI to add, edit, or remove records.
;
; Programmer:     panofish (www.autohotkey.com)
; Modified by:    just me  (www.autohotkey.com)
;                 samsonwang
; AutoHotkey:     v2.0-beta.3+
; ======================================================================================================================
; REMOVED: #NoEnv
#SingleInstance Force
; REMOVED: SetBatchLines, -1
ListLines false
#Include "Class_MySQLAPI_v2new.ahk" ; pull from local directory
OnExit AppExit
Global MySQL_SUCCESS := 0
; ======================================================================================================================
; Settings
; ======================================================================================================================
UserID := "root"           ; User name - must have privileges to create databases
UserPW := ""           ; User''s password
Server := "localhost"      ; Server''s host name or IP address
Database := "Test"         ; Name of the database to work with
DropDatabase := False      ; DROP DATABASE
DropTable := False         ; DROP TABLE Address
; ======================================================================================================================
; Connect to MySQL
; ======================================================================================================================
; Instantiate a MYSQL object
; If !(My_DB := New MySQLAPI)
If !(My_DB := MySQLAPI())
   ExitApp
; Get the version of libmariadb.dll
ClientVersion := My_DB.Get_Client_Info()
; Connect to the server, Host can be a hostname or IP address
If !My_DB.Connect(Server, UserID, UserPW) {  ; Host, User, Password
   MsgBox "Connection failed!`r`n" . My_DB.ErrNo() . " - " . My_DB.Error(), "MySQL Error!", 16
   ExitApp
}
; ======================================================================================================================
; CREATE DADABASE Test
; ======================================================================================================================
If (DropDatabase)
   My_DB.Query("DROP DATABASE IF EXISTS " . DataBase)
SQL := "CREATE DATABASE IF NOT EXISTS " . Database . " DEFAULT CHARACTER SET utf8 DEFAULT COLLATE utf8_bin"
My_DB.Query(SQL)
; ======================================================================================================================
; Select the database as default
; ======================================================================================================================
My_DB.Select_DB(Database)
; ======================================================================================================================
; CREATE TABLE Address
; ======================================================================================================================
If (DropTable)
   My_DB.Query("DROP TABLE IF EXISTS Address")
SQL := "CREATE TABLE IF NOT EXISTS Address ( "
     . "Name VARCHAR(50) NULL, "
     . "Address VARCHAR(50) NULL, "
     . "City VARCHAR(50) NULL, "
     . "State VARCHAR(2) NULL, "
     . "Zip INT(5) ZEROFILL NULL, "
     . "PRIMARY KEY (Name) )"
My_DB.Query(SQL)
; ======================================================================================================================
; Build GUI
; ======================================================================================================================
Fields := ["Name", "Address", "City", "State", "Zip"]
myGui := Gui()
myGui.OnEvent("Close", GuiClose)
; myGui.Default()
myGui.MarginX := "10", myGui.MarginY := "10"
myGui.Opt("+OwnDialogs")
myGui.Add("Text", "Section Right w70", "Name")
ogcEditName := myGui.Add("Edit", "x+10 w250 vName")
ogcEditName.GetPos(&CX, &CY, &CW, &CH)
myGui.Add("Text", "xs Right w70", "Address")
ogcEditAddress := myGui.Add("Edit", "x+10 w250 vAddress")
myGui.Add("Text", "xs Right w70", "City")
ogcEditCity := myGui.Add("Edit", "x+10 w150 vCity")
myGui.Add("Text", "xs Right w70", "State")
ogcEditState := myGui.Add("Edit", "x+10 w30 Uppercase Limit2 vState")
myGui.Add("Text", "xs Right w70", "Zip")
ogcEditZip := myGui.Add("Edit", "x+10 w60  Number Limit5 vZip")
ogcButtonBtnAddUpd := myGui.Add("Button", "ys x410 w100 h" . CH . " vBtnAddUpd  Default", "Add")
ogcButtonBtnAddUpd.OnEvent("Click", SubBtnAction)
ogcButtonClear := myGui.Add("Button", "wp hp", "Clear")
ogcButtonClear.OnEvent("Click", SubBtnClear)
ogcButtonBtnDelete := myGui.Add("Button", "wp hp vBtnDelete", "Delete")
ogcButtonBtnDelete.OnEvent("Click", SubBtnAction)
ogcButtonReload := myGui.Add("Button", "wp hp", "Reload")
ogcButtonReload.OnEvent("Click", UpdateListView)
; To increase performance use Count option if you know the max number of lines
; LV0x00010000 = LV_EX_DOUBLEBUFFER
ogcList1 := myGui.Add("ListView", "xs r10 w500 AltSubmit vList1 Grid -Multi +LV0x00010000")
ogcList1.OnEvent("DoubleClick", SubListView)
SB := myGui.Add("StatusBar")
myGui.Title := "MySQLAPI Demo - Client version: " . ClientVersion
myGui.Show()
; UpdateListView()
ControlClick ogcButtonReload
; ======================================================================================================================
; Make a first query and get the result "manually"
; ======================================================================================================================
SQL := "SELECT COUNT(*) FROM Address"
If (My_DB.Query(SQL) = MySQL_SUCCESS) {
   My_Result := My_DB.Store_Result()
   My_Field := My_DB.Fetch_Field(My_Result)
   FieldName := StrGet(NumGet(My_Field + 0, 0, "UPtr"), "UTF-8")
   My_Row := My_DB.Fetch_Row(My_Result)
   FieldValue := StrGet(NumGet(My_Row + 0, 0, "UPtr"), "UTF-8")
   My_DB.Free_Result(My_Result)
}
MsgBox "Query:`r`n" . SQL . "`r`n`r`nResult:`r`nName = " . FieldName . "`r`nValue = " . FieldValue, "MySQLAPI Demo", 0
Return


; ======================================================================================================================
; ListView event handler
; DoubleClick a row to delete / edit the entry
; ======================================================================================================================
SubListView(GuiCtrlObj, Info){
   ; If (A_GuiEvent = "DoubleClick") {
      ; CurrentRow := A_EventInfo
      CurrentRow := Info
	  if CurrentRow = 0
	  {
         ControlClick ogcButtonClear
	  }
	  else
	  {
         Loop GuiCtrlObj.GetCount("Column") {
            Value := GuiCtrlObj.GetText(CurrentRow,A_Index)
            If (Fields[A_Index] = "Zip")
               Value := SubStr("0000" . Value, -4)
            ogcEdit%Fields[A_Index]%.Value := Value
		 }
         ogcEditName.Opt("+ReadOnly")
         ogcButtonBtnAddUpd.Text := "Update"
         ogcEditAddress.Focus()
      }
   ; }
   Return
}

; ======================================================================================================================
; Perform the requested action
; ======================================================================================================================
SubBtnAction(GuiCtrlObj, Info){
   myGui.Opt("+OwnDialogs")
   oSaved := myGui.Submit(0)
   Name := oSaved.Name
   Address := oSaved.Address
   City := oSaved.City
   State := oSaved.State
   Zip := oSaved.Zip
   
   Name := ogcEditName.Text
   If !Trim(Name, " `t`r`n")
      Return
   ; Escape mysql special characters in case user entered them
   V1 := My_DB.Real_Escape_String(&Name)
   V2 := My_DB.Real_Escape_String(&Address)
   V3 := My_DB.Real_Escape_String(&City)
   V4 := My_DB.Real_Escape_String(&State)
   V5 := My_DB.Real_Escape_String(&Zip)
   ; Get the action
   ; Action := ogc%A_GuiControl%.Text
   Action := GuiCtrlObj.Text
   SQL := ""
   If (Action = "Add") {
      ;-----------------------------------------------------------------------------------------------------------------
      ; Insert new record
      ;-----------------------------------------------------------------------------------------------------------------
      SB.SetText("Inserting new record!")
      SQL := "INSERT INTO Address ( Name, Address, City, State, Zip) "
           . "VALUES ( '" . V1 . "', '" . V2 . "', '" . V3 . "', '" . V4 . "', '" . V5 . "')"
      Done := "inserted!"
   }
   Else If (Action = "Delete") {
      ;-----------------------------------------------------------------------------------------------------------------
      ; Delete record
      ;-----------------------------------------------------------------------------------------------------------------
      msgResult := MsgBox("Do you really want to delete '" . Name . "'?", "Delete", 36)
      if (msgResult = "Yes")
      {      
         SB.SetText("Deleting record!")
         SQL := "DELETE FROM Address WHERE Name = '" . Name . "'"
         Done := "deleted!"
      }
   }
   Else If (Action = "Update") {
      ;-----------------------------------------------------------------------------------------------------------------
      ; Update record
      ;-----------------------------------------------------------------------------------------------------------------
      SB.SetText("Updating record!")
      SQL := "UPDATE Address SET Address = '" . V2 . "', City = '" . V3 . "', State='" . V4 . "', Zip='" . V5 . "' "
           . "WHERE Name = '" . V1 . "'"
      Done := "updated!"
   }
   If (SQL) {
      If (My_DB.Query(SQL) = MySQL_SUCCESS) {
         Rows := My_DB.Affected_Rows()
		 ; UpdateListView()
         ; SubBtnClear()
		 Sleep 100						; Need to wait for some time for adding transaction to complete.
		 ControlClick ogcButtonReload
		 ControlClick ogcButtonClear
         SB.SetText(Rows . " row(s) " . Done)
      } Else {
		 MsgBox My_DB.ErrNo() . ": " . My_DB.Error(), "MySQL Error!", 16
      }
   }
   Return
}
; ======================================================================================================================
; Clear Edits
; ======================================================================================================================
SubBtnClear(GuiCtrlObj, Info){
   For Each, Ctrl In Fields
      ogcEdit%Fields[A_Index]%.Value := ""
   ogcEditName.Opt("-Readonly")
   ogcButtonBtnAddUpd.Text := "Add"
   ogcEditName.Focus()
   SB.SetText("")
   Return
}
; ======================================================================================================================
; Fill ListView with existing addresses from database
; ======================================================================================================================
UpdateListView(GuiCtrlObj, Info){
   SQL := "SELECT Name, Address, City, State, Zip FROM Address ORDER BY Name"
   If (My_DB.Query(SQL) = MySQL_SUCCESS) {
      Result := My_DB.GetResult()
      LV_Fill(Result, "List1")
      SB.SetText("ListView has been updated: " . Result.Columns . " columns - " . Result.Rows . " rows.")
   }
   Return
}
; ======================================================================================================================
; GUI was closed
; ======================================================================================================================
GuiClose(thisGui){
   ExitApp
}
AppExit(ExitReason, ExitCode){
   My_DB := ""
   ExitApp
}
; ======================================================================================================================
; Fill ListView with the result of a query.
; Note: The current data in the ListView are replaced with the new data.
; Parameters:
;    Result       -  Result object returned from MySQLDAPI.Query()
;    ListViewName -  Name of the ListView''s asociated variable
; ======================================================================================================================
LV_Fill(Result, ListViewName) {
   ;--------------------------------------------------------------------------------------------------------------------
   ; Delete all rows and columns of the ListView
   ;--------------------------------------------------------------------------------------------------------------------
   ogc%ListViewName%.Opt("-Redraw")             ; to improve performance, turn off redraw then turn back on at end
   ; myGui.ListView(ListViewName)                   ; specify which listview will be updated with LV commands
   ogc%ListViewName%.Delete()                                     ; delete all rows in the listview
   Loop ogc%ListViewName%.GetCount("Column")                   ; delete all columns of the listview
      ogc%ListViewName%.DeleteCol(1)
   ;--------------------------------------------------------------------------------------------------------------------
   ; Parse field names
   ;--------------------------------------------------------------------------------------------------------------------
   Loop Result.Fields.Length {
      ogc%ListViewName%.InsertCol(A_Index, "", Result.Fields[A_Index].Name)
   }
   ;--------------------------------------------------------------------------------------------------------------------
   ; Parse rows
   ;--------------------------------------------------------------------------------------------------------------------
   Count := 0
   Loop Result.Records.Length {
      RowNum := ogc%ListViewName%.Add("")                         ; add a blank row to the listview
      Row := Result.Records[A_Index]                                ; extract the row from the result
      Loop Row.Length {                              ; populate the columns of the current row
         Value := Row[A_Index]
         If (A_Index = 5)
            Value := SubStr("0000" . Value, -4)
         ogc%ListViewName%.Modify(RowNum, "Col" . A_Index, Value) ; update current column of current row
      }
   }
   ;--------------------------------------------------------------------------------------------------------------------
   ; Autosize columns: should be done outside the row loop to improve performance
   ;--------------------------------------------------------------------------------------------------------------------
   Loop ogc%ListViewName%.GetCount("Column")
      ogc%ListViewName%.ModifyCol(A_Index, "AutoHdr") ; Autosize header.
   ogc%ListViewName%.ModifyCol(1, "Sort Logical")
   ogc%ListViewName%.Opt("+Redraw") ; to improve performance, turn off redraw at beginning then turn back on at end
   Return
}
