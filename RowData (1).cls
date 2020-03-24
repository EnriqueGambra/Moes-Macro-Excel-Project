VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RowData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class module:  RowData
Private obj_row_name As String
Private obj_account_name As String
Private obj_name As String

Private obj_is_debit As Boolean

Private obj_index As Integer

Private obj_currency_value As Currency

Private Sub Class_Initialize()
    'This class will contain all the data for a particular row
    obj_name = ""
    obj_is_debit = True
    obj_index = 0
End Sub

Property Let index(value As Integer)
    'Sets the index value
    obj_index = value
End Property

Property Get index() As Integer
    'Gets the index value
    index = obj_index
End Property

Property Let row_name(value As String)
    'Sets the row_name
    obj_row_name = value
End Property

Property Get row_name() As String
    'Returns the row name
    row_name = obj_row_name
End Property

Property Let account_name(value As String)
    'Sets the account name
    obj_account_name = value
End Property

Property Get account_name() As String
    'Retrieves the account name
    account_name = obj_account_name
End Property

Property Let currency_value(value As Currency)
    'Sets the currency value for the row
    obj_currency_value = value
End Property

Property Get currency_value() As Currency
    'Gets the currency value
    currency_value = obj_currency_value
End Property

Property Let is_debit(value As Boolean)
    'Sets the is_debit property as a row
    obj_is_debit = value
End Property

Property Get is_debit() As Boolean
    'Gets the property for is_debit
    is_debit = obj_is_debit
End Property

Property Let name(value As String)
    'Sets the name string for the row
    obj_name = value
End Property

Property Get name() As String
    'Gets the name string for the row
    name = obj_name
End Property




