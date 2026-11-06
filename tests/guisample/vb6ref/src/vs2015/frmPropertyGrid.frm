VERSION 5.00
Object = "{5B44EC52-B95B-45CF-98FF-A49DFEED5A92}#20.3#0"; "PropertyGrid.ocx"
Begin VB.Form frmPropertyGrid 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XtremePropertyGrid.PropertyGrid wndPropertyGrid 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      _Version        =   983040
      _ExtentX        =   7646
      _ExtentY        =   5106
      _StockProps     =   68
      ToolBarVisible  =   -1  'True
      HelpVisible     =   -1  'True
      PropertySort    =   0
      VisualTheme     =   9
   End
End
Attribute VB_Name = "frmPropertyGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    
    Dim Category As PropertyGridItem
    Dim Item As PropertyGridItem, ItemDirectory As PropertyGridItem
    
    wndPropertyGrid.ToolBarVisible = False
    
'Common Samples
    Set Category = wndPropertyGrid.AddCategory("Document Settings")
  
    Set Item = Category.AddChildItem(PropertyItemBool, "SaveOnClose", True)
    Category.AddChildItem PropertyItemFont, "WindowFont", Me.Font
    Category.AddChildItem PropertyItemString, "WindowSize", "100; 100"
    Category.Expanded = True
    Item.Selected = True
    
    
    Set Category = wndPropertyGrid.AddCategory("Global Settings")
    Category.AddChildItem PropertyItemString, "Greeting Text", "Welcome to your application!"
    Category.AddChildItem PropertyItemNumber, "ItemsInMRUList", 4
    Set Item = Category.AddChildItem(PropertyItemNumber, "MaxRepeatRate", 10)
    Item.Description = "The rate in milliseconds that the text will repeat."
    Item = Item + 6
    Category.AddChildItem PropertyItemColor, "ToolbarColor", &HAA00FF
    Category.Expanded = True
    
    Set Category = wndPropertyGrid.AddCategory("Version")
    Set Item = Category.AddChildItem(PropertyItemString, "AppVerion", "1.0")
    
    Item.ReadOnly = True
    Category.AddChildItem PropertyItemString, "Language", "Russian"
    Category.Expanded = True
    
    Item.Tag = 500
    Item = "2.0.0"
    
    Debug.Print Item.Parent.Caption
    Debug.Print Item.Tag
    Debug.Print Item
    
'Items Type Samples
    Set Category = wndPropertyGrid.AddCategory("Items")
    Set Item = Category.AddChildItem(PropertyItemString, "StringItem", "String")
    Set Item = Category.AddChildItem(PropertyItemNumber, "NumberItem", 10)
    Item.Tag = 200
    Set Item = Category.AddChildItem(PropertyItemColor, "ColorItem", RGB(10, 20, 30))
    Set Item = Category.AddChildItem(PropertyItemBool, "BoolItem", True)
    Set Item = Category.AddChildItem(PropertyItemFont, "FontItem", Me.Font)
'Flags Sample
    Set ItemDirectory = Category.AddChildItem(PropertyItemString, "Directory", "C:\")
    ItemDirectory.Flags = ItemHasEdit + ItemHasExpandButton
    ItemDirectory.Tag = ITEM_DIRECTORY
  
  
'Childs and Categories property Samples
    Dim Items As PropertyGridItems
    Set Items = wndPropertyGrid.Categories
    Debug.Print Items.Count
    Debug.Print Items(1).Caption
    Debug.Print wndPropertyGrid.Categories(2).Caption
    
    Set Category = wndPropertyGrid.Categories(3)
    Debug.Print Category.Caption
    Debug.Print Category.Childs(1).Caption; " = "; Category.Childs(1)
    
' Clear and Remove
    'Category.Childs.Remove(1)
    'Category.Childs.Clear
   ' wndPropertyGrid.Categories.Clear
    
    
'Enumeration Samples
    For Each Category In wndPropertyGrid.Categories
        
        Debug.Print "Category "; Category.Caption
        For Each Item In Category.Childs
            Debug.Print "  Item "; Item.Caption
    
        Next
    
    Next
    
'FindItem
    Set Item = wndPropertyGrid.FindItem("Version")
    If Not (Item Is Nothing) Then
        Debug.Print Item.Caption
    End If
    
    Set Item = wndPropertyGrid.FindItem(500)
    If Not (Item Is Nothing) Then
        Debug.Print Item.Caption
    End If

'Constraints Samples
    Set Item = wndPropertyGrid.FindItem("Language")
    Item.Flags = ItemHasEdit + ItemHasComboButton
    
    Item.Constraints.Add ("Neutral")
    Item.Constraints.Add ("Arabic")
    Item.Constraints.Add ("German")
    Item.Constraints.Add ("Chinese(Taiwan)")
    Item.Constraints.Add ("English (United States)")
    Item.Constraints.Add ("France")
    Item.Constraints.Add ("Russian")
    
    Dim Constraint As Variant
    
    Debug.Print Item.Constraints(3)
    For Each Constraint In Item.Constraints
        Debug.Print Constraint
    Next
    
    Exit Sub
Err:
    Debug.Print "Error: "; Err.Description
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    wndPropertyGrid.Move Me.ScaleLeft, Me.ScaleTop, Me.ScaleWidth, Me.ScaleHeight

End Sub
