Attribute VB_Name = "modConstant"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3B3781ED01EA"
Option Explicit
'**********************************************************************
'
'           Module Name: modAPIConst.bas
'
'           Purpose    : Declare various constant used internally by
'                        the application
'
'           Author     : Joyprakash Saikia
'

'
'
'
'**********************************************************************
'##ModelId=3B3781ED0294
Public Const APP_NAME = "JxDBReport" 'application name




'Application Specified Error
'The Corresponding String is Stored on the Resource File

Public Const ERR_REPORT_SIZE_UNDEFINED = 513

Public Const ERR_POS_EXCEED_SIZE = 515 'xpos or ypos exceed printable area
Public Const ERR_INVALID_PROP_VALUE = 516


'DataSource Related Error
Public Const ERR_NO_DATASOURCE = 600
Public Const ERR_NO_SUCH_FIELD = 601 ' not field exist in recordset
Public Const ERR_NO_RECORD = 602 ' no record in datasource

Public Const ERR_NO_PRINTER = 999 'no printer installed
Public Const ERR_CREATING_RPT_FILE = 1000
Public Const ERR_OPENING_RPT_FILE = 1001

'preview related msg
Public Const MSG_PAGE_NOTEXIST = 2000 'no page exist to be view/print

