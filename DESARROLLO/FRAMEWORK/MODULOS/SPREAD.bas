Attribute VB_Name = "SPREAD"
Option Explicit

Global Const SS_CELL_TYPE_INTEGER As Integer = 3
Global Const SS_CELL_TYPE_EDIT As Integer = 1
Global Const SS_CELL_TYPE_FLOAT As Integer = 2
Global Const SS_CELL_TYPE_DATE As Integer = 0
'Private Const SS_CELL_TYPE_DATE As Integer = 0
Global Const SS_CELL_TYPE_CHECKBOX As Integer = 10
Global Const SS_CELL_TIPE_COMBOBOX As Integer = 8
Global Const SS_ACTION_INSERT_ROW As Integer = 7

' SelBackColor property settings
Global Const SPREAD_COLOR_NONE = &H8000000B
Global Const SS_ACTION_SORT = 25
Global Const SS_SORT_BY_ROW = 0
Global Const SS_SORT_BY_COL = 1
Global Const SS_SORT_ORDER_ASCENDING = 1
