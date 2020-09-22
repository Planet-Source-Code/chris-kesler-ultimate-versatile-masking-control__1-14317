VERSION 5.00
Begin VB.UserControl ucArchGenMask 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1140
   ControlContainer=   -1  'True
   ForwardFocus    =   -1  'True
   HitBehavior     =   2  'Use Paint
   PropertyPages   =   "ucNumerickMask.ctx":0000
   ScaleHeight     =   240
   ScaleWidth      =   1140
   ToolboxBitmap   =   "ucNumerickMask.ctx":0091
   Begin VB.TextBox txtCBR 
      Height          =   315
      Left            =   -30
      TabIndex        =   0
      Top             =   -30
      Width           =   1215
   End
End
Attribute VB_Name = "ucArchGenMask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
Dim intMaskType As Integer
Private objDate As clsDate
Private objPhone As clsPhone
Private objSSN As clsSSN
Private objZip As clsZip
Private objIntlZip As clsIntlZip
Private objCustom As clsCustom
Private objCurrency As clsCurrency
Private objEmail As clsEmail
Private objIP As clsIP
Private objWholeNum As clsWholeNumber

'Property Variables
Dim m_blnSign As Boolean
Dim m_blnAlpha As Boolean
Dim m_blnNumeric As Boolean
Dim m_strCharAllowed As String
Dim m_intMaxAllowed As Integer
Dim m_intDateTypeMask As Integer
Dim m_intEmailMask As Integer
Dim m_blnSpaces As Boolean
Dim m_blnPoints As Boolean
Dim m_strMask As String
Dim m_blnArea As Boolean
Dim m_blnPars As Boolean
Dim m_blnSpc As Boolean
Dim m_blnDash As Boolean
Dim m_blnExt  As Boolean
Dim m_blnDashes As Boolean
Dim m_blnAllowNegative As Boolean
Dim m_blnUsePrecision As Boolean
Dim m_intPrecision As Integer
Dim m_blnOnlyFive As Boolean
Dim m_blnAllCaps As Boolean

Dim strVersion As String
Dim m_MaskType As Integer

'Property Constants
Const m_def_blnSign = True
Const m_def_blnAlpha = True
Const m_def_blnNumeric = True
Const m_def_strCharAllowed = ""
Const m_def_intMaxAllowed = 25
Const m_def_intDateTypeMask = 0
Const m_def_intEmailMask = 0
Const m_def_blnSpaces = True
Const m_def_blnPoints = True
Const m_def_blnArea = True
Const m_def_blnPars = True
Const m_def_blnSpc = True
Const m_def_blnDash = True
Const m_def_blnExt = True
Const m_def_blnDashes = True
Const m_def_blnOnlyFive = True
Const m_def_blnAllowNegative = False
Const m_def_blnUsePrecision = True
Const m_def_intPrecision = 2
Const m_def_blnAllCaps = False

'Declarations to set Masktype via internal and external sources.
Public Enum Masktype
    DateMask = 0
    IntlZipMask = 1
    PhoneMask = 2
    SSNMask = 3
    ZipMask = 4
    CurrencyMask = 5
    EmailMask = 6
    CustomMask = 7
    IPMask = 8
    WholeNumber = 9
End Enum

'Event Declarations:
Event Click() 'MappingInfo=txtCBR,txtCBR,-1,Click
Event DblClick() 'MappingInfo=txtCBR,txtCBR,-1,DblClick
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=txtCBR,txtCBR,-1,KeyDown
Event KeyPress(KeyAscii As Integer) 'MappingInfo=txtCBR,txtCBR,-1,KeyPress
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=txtCBR,txtCBR,-1,KeyUp
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Event WriteProperties(PropBag As PropertyBag) 'MappingInfo=UserControl,UserControl,-1,WriteProperties
Event Validate(Cancel As Boolean) 'MappingInfo=txtCBR,txtCBR,-1,Validate
Event Show() 'MappingInfo=UserControl,UserControl,-1,Show
Event Resize() 'MappingInfo=UserControl,UserControl,-1,Resize
Event ReadProperties(PropBag As PropertyBag) 'MappingInfo=UserControl,UserControl,-1,ReadProperties
Event Paint() 'MappingInfo=UserControl,UserControl,-1,Paint
Event OLEStartDrag(Data As DataObject, AllowedEffects As Long) 'MappingInfo=UserControl,UserControl,-1,OLEStartDrag
Event OLESetData(Data As DataObject, DataFormat As Integer) 'MappingInfo=UserControl,UserControl,-1,OLESetData
Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean) 'MappingInfo=UserControl,UserControl,-1,OLEGiveFeedback
Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer) 'MappingInfo=UserControl,UserControl,-1,OLEDragOver
Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,OLEDragDrop
Event OLECompleteDrag(Effect As Long) 'MappingInfo=UserControl,UserControl,-1,OLECompleteDrag
Event Hide() 'MappingInfo=UserControl,UserControl,-1,Hide
Event Change() 'MappingInfo=txtCBR,txtCBR,-1,Change
Event InitProperties() 'MappingInfo=UserControl,UserControl,-1,InitProperties

'ENUMS
Public Enum mBorderStyle
    None = 0
    [Fixed Single]
End Enum

Public Enum mAppearance
   Flat = 0
   [3D]
End Enum


Private Sub txtCBR_KeyPress(KeyAscii As Integer)
    SetObject
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    Else
        If txtCBR.SelLength > 0 Then txtCBR.Text = Left(txtCBR.Text, txtCBR.SelStart)
        Select Case intMaskType
            Case 0  'date
                objDate.DateTypeMask = m_intDateTypeMask
                txtCBR.Text = objDate.Mask(KeyAscii, txtCBR.Text)
            Case 1  'International Zip (for Canada only)
                objIntlZip.IntlZipMask = m_blnSpaces
                txtCBR.Text = objIntlZip.Mask(KeyAscii, txtCBR.Text)
            Case 2  'phone
                objPhone.PhoneFormat = m_strMask
                objPhone.UseArea = m_blnArea
                objPhone.UseDash = m_blnDash
                objPhone.UseExt = m_blnExt
                objPhone.UsePars = m_blnPars
                objPhone.UseSpc = m_blnSpc
                txtCBR.Text = objPhone.Mask(KeyAscii, txtCBR.Text)
            Case 3  'ssn
                objSSN.SSNMask = m_blnDashes
                txtCBR.Text = objSSN.Mask(KeyAscii, txtCBR.Text)
            Case 4  'zip
                objZip.OnlyFive = m_blnOnlyFive
                txtCBR.Text = objZip.Mask(KeyAscii, txtCBR.Text)
            Case 5  'currency
                objCurrency.UseSign = m_blnSign
                txtCBR.Text = objCurrency.Mask(KeyAscii, txtCBR.Text)
            Case 6  'email
                objEmail.EmailMask = m_intEmailMask
                txtCBR.Text = objEmail.Mask(KeyAscii, txtCBR.Text)
            Case 7  'custom
                objCustom.AllowAlpha = m_blnAlpha
                objCustom.AllowNumeric = m_blnNumeric
                objCustom.CharAllowed = m_strCharAllowed
                objCustom.MaxAllowed = m_intMaxAllowed
                objCustom.AllCaps = m_blnAllCaps
                txtCBR.Text = objCustom.Mask(KeyAscii, txtCBR.Text)
            Case 8  'IP
                objIP.IPMask = m_blnPoints
                txtCBR.Text = objIP.Mask(KeyAscii, txtCBR.Text)
            Case 9 'Whole Numbers
                objWholeNum.AllowNegatives = m_blnAllowNegative
                objWholeNum.UsePrecision = m_blnUsePrecision
                objWholeNum.PrecisionValue = m_intPrecision
                txtCBR.Text = objWholeNum.Mask(KeyAscii, txtCBR.Text)
        End Select
        txtCBR.SelStart = Len(txtCBR.Text)
    End If
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    RaiseEvent ReadProperties(PropBag)
    strVersion = App.Major & "." & App.Revision & "." & App.Minor
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    UserControl.PaletteMode = PropBag.ReadProperty("PaletteMode", 3)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    UserControl.FontTransparent = PropBag.ReadProperty("FontTransparent", True)
    txtCBR.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.FillColor = PropBag.ReadProperty("FillColor", &H0&)
    UserControl.OLEDropMode = PropBag.ReadProperty("OLEDropMode", 0)
    txtCBR.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    txtCBR.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    txtCBR.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    txtCBR.WhatsThisHelpID = PropBag.ReadProperty("WhatsThisHelpID", 0)
    txtCBR.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
    txtCBR.Tag = PropBag.ReadProperty("Tag", "")
    txtCBR.SelText = PropBag.ReadProperty("SelText", "")
    txtCBR.SelStart = PropBag.ReadProperty("SelStart", 0)
    txtCBR.SelLength = PropBag.ReadProperty("SelLength", 0)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    txtCBR.PasswordChar = PropBag.ReadProperty("PasswordChar", "")
    Set Palette = PropBag.ReadProperty("Palette", Nothing)
    txtCBR.OLEDragMode = PropBag.ReadProperty("OLEDragMode", 0)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    txtCBR.MaxLength = PropBag.ReadProperty("MaxLength", 0)
    Set MaskPicture = PropBag.ReadProperty("MaskPicture", Nothing)
    txtCBR.Locked = PropBag.ReadProperty("Locked", False)
    txtCBR.FontUnderline = PropBag.ReadProperty("FontUnderline", 0)
    txtCBR.FontStrikethru = PropBag.ReadProperty("FontStrikethru", 0)
    txtCBR.CausesValidation = PropBag.ReadProperty("CausesValidation", True)
    txtCBR.Appearance = PropBag.ReadProperty("Appearance", 1)
    m_blnSign = PropBag.ReadProperty("CurrencySign", True) 'Currency
    m_blnAlpha = PropBag.ReadProperty("AllowAlpha", True) 'Custom
    m_blnNumeric = PropBag.ReadProperty("AllowNumeric", True) 'Custom
    m_strCharAllowed = PropBag.ReadProperty("CharAllowed", "") 'Custom
    m_blnAllCaps = PropBag.ReadProperty("AllCaps", False) 'Custom
    m_intMaxAllowed = PropBag.ReadProperty("MaxAllowed", 25) 'Custom
    m_intDateTypeMask = PropBag.ReadProperty("DateType", 0) 'Date
    m_intEmailMask = PropBag.ReadProperty("EmailType", 0) 'Email
    m_blnSpaces = PropBag.ReadProperty("IZIPAllowSpaces", True) 'IntlZip
    m_blnPoints = PropBag.ReadProperty("IPAllowPoints", True) 'IP Address
    m_strMask = PropBag.ReadProperty("PhoneMaskType", "(&&&) &&&-&&&& X&&&&") 'Phone Mask
    m_blnArea = PropBag.ReadProperty("AreaCode", True) 'Phone
    m_blnPars = PropBag.ReadProperty("Parenthesis", True) 'Phone
    m_blnSpc = PropBag.ReadProperty("PhnSpaces", True) 'Phone
    m_blnDash = PropBag.ReadProperty("PhnDashes", True) 'Phone
    m_blnExt = PropBag.ReadProperty("Extension", True) 'Phone
    m_blnDashes = PropBag.ReadProperty("SSNDashes", True) 'SSN
    m_blnOnlyFive = PropBag.ReadProperty("ZipOnlyFive", True) 'Zip
    intMaskType = PropBag.ReadProperty("MaskType", 0) 'Mask Type
    m_blnAllowNegative = PropBag.ReadProperty("AllowNegative", False) ' Whole Numbers
    m_blnUsePrecision = PropBag.ReadProperty("UsePrecision", True)
    m_intPrecision = PropBag.ReadProperty("Precision", 2)
    
End Sub
Private Sub UserControl_Resize()
    txtCBR.Width = UserControl.Width
    txtCBR.Height = UserControl.Height
End Sub
Public Property Get Text() As String
    Text = txtCBR.Text
End Property
Public Property Let Text(ByVal Value As String)
    txtCBR.Text = Value
    PropertyChanged "Text"
End Property
Public Property Get Value() As String
    Value = txtCBR.Text
End Property
Public Property Let Value(ByVal sValue As String)
    txtCBR.Text = sValue
    PropertyChanged "Text"
End Property
Public Property Get Tag() As String
    Tag = txtCBR.Tag
End Property
Public Property Let Tag(ByVal Value As String)
    txtCBR.Tag = Value
    PropertyChanged "Tag"
End Property
Public Property Get Masktype() As Masktype
Attribute Masktype.VB_ProcData.VB_Invoke_Property = "pagGeneral"
    Masktype = intMaskType
End Property
Public Property Let Masktype(ByVal Value As Masktype)
    intMaskType = Value
    'PropertyChanged "MaskType"
    SetObject
End Property

Private Sub SetObject()
    Set objDate = Nothing
    Set objPhone = Nothing
    Set objSSN = Nothing
    Set objZip = Nothing
    Set objCurrency = Nothing
    Set objEmail = Nothing
    Set objCustom = Nothing
    Set objIP = Nothing
    Set objIntlZip = Nothing
    Set objWholeNum = Nothing
    
    Select Case intMaskType
        Case 0  'date
            Set objDate = New clsDate
        Case 1  'International Zip Format (for Canada Only)
            Set objIntlZip = New clsIntlZip
        Case 2  'phone
            Set objPhone = New clsPhone
        Case 3  'SSN
            Set objSSN = New clsSSN
        Case 4  'ZIP
            Set objZip = New clsZip
        Case 5  'currency
            Set objCurrency = New clsCurrency
        Case 6  'email
            Set objEmail = New clsEmail
        Case 7  'custom
            Set objCustom = New clsCustom
        Case 8  'IP
            Set objIP = New clsIP
        Case 9
            Set objWholeNum = New clsWholeNumber
    End Select
   
End Sub

Private Sub UserControl_Terminate()
    Set objDate = Nothing
    Set objPhone = Nothing
    Set objSSN = Nothing
    Set objZip = Nothing
    Set objCurrency = Nothing
    Set objEmail = Nothing
    Set objCustom = Nothing
    Set objIP = Nothing
    Set objIntlZip = Nothing
    Set objWholeNum = Nothing
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtCBR,txtCBR,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
   BackColor = txtCBR.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
   txtCBR.BackColor() = New_BackColor
   PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtCBR,txtCBR,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
   
   ForeColor = txtCBR.ForeColor
   
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
   txtCBR.ForeColor() = New_ForeColor
   PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
  Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
   UserControl.Enabled() = New_Enabled
   PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtCBR,txtCBR,-1,Font
Public Property Get Font() As Font
   Set Font = txtCBR.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
   Set txtCBR.Font() = New_Font
   PropertyChanged "Font"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtCBR,txtCBR,-1,Refresh
Public Sub Refresh()
   txtCBR.Refresh
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtCBR,txtCBR,-1,SetFocus
Public Sub SetFocus()
    txtCBR.SetFocus
End Sub

Private Sub txtCBR_Click()
   RaiseEvent Click
End Sub
Private Sub txtCBR_DblClick()
   RaiseEvent DblClick
End Sub
Private Sub txtCBR_KeyDown(KeyCode As Integer, Shift As Integer)
   RaiseEvent KeyDown(KeyCode, Shift)
End Sub
Private Sub txtCBR_KeyUp(KeyCode As Integer, Shift As Integer)
   RaiseEvent KeyUp(KeyCode, Shift)
End Sub
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    RaiseEvent WriteProperties(PropBag)
    
    Call PropBag.WriteProperty("BackColor", txtCBR.BackColor, &H80000005)
    Call PropBag.WriteProperty("ForeColor", txtCBR.ForeColor, &H80000008)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", txtCBR.Font, Ambient.Font)
    Call PropBag.WriteProperty("WhatsThisHelpID", txtCBR.WhatsThisHelpID, 0)
    Call PropBag.WriteProperty("ToolTipText", txtCBR.ToolTipText, "")
    Call PropBag.WriteProperty("SelText", txtCBR.SelText, "")
    Call PropBag.WriteProperty("SelStart", txtCBR.SelStart, 0)
    Call PropBag.WriteProperty("SelLength", txtCBR.SelLength, 0)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("PasswordChar", txtCBR.PasswordChar, "")
    Call PropBag.WriteProperty("PaletteMode", UserControl.PaletteMode, 3)
    Call PropBag.WriteProperty("Palette", Palette, Nothing)
    Call PropBag.WriteProperty("OLEDropMode", UserControl.OLEDropMode, 0)
    Call PropBag.WriteProperty("OLEDragMode", txtCBR.OLEDragMode, 0)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MaxLength", txtCBR.MaxLength, 0)
    Call PropBag.WriteProperty("MaskPicture", MaskPicture, Nothing)
    Call PropBag.WriteProperty("Locked", txtCBR.Locked, False)
    Call PropBag.WriteProperty("FontUnderline", txtCBR.FontUnderline, 0)
    Call PropBag.WriteProperty("FontTransparent", UserControl.FontTransparent, True)
    Call PropBag.WriteProperty("FontStrikethru", txtCBR.FontStrikethru, 0)
    Call PropBag.WriteProperty("FontSize", txtCBR.FontSize, 0)
    Call PropBag.WriteProperty("FontName", txtCBR.FontName, "")
    Call PropBag.WriteProperty("FontItalic", txtCBR.FontItalic, 0)
    Call PropBag.WriteProperty("FontBold", txtCBR.FontBold, 0)
    Call PropBag.WriteProperty("FillColor", UserControl.FillColor, &H0&)
    Call PropBag.WriteProperty("CausesValidation", txtCBR.CausesValidation, True)
    Call PropBag.WriteProperty("Appearance", txtCBR.Appearance, 1)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 1)
    Call PropBag.WriteProperty("BorderStyle", txtCBR.BorderStyle, 1)
    Call PropBag.WriteProperty("Tag", txtCBR.Tag, "")
    Call PropBag.WriteProperty("CurrencySign", m_blnSign, True) 'Currency
    Call PropBag.WriteProperty("AllowAlpha", m_blnAlpha, True) 'Custom
    Call PropBag.WriteProperty("AllowNumeric", m_blnNumeric, True) 'Custom
    Call PropBag.WriteProperty("CharAllowed", m_strCharAllowed, "") 'Custom
    Call PropBag.WriteProperty("MaxAllowed", m_intMaxAllowed, 25) 'Custom
    Call PropBag.WriteProperty("AllCaps", m_blnAllCaps, False) 'Custom
    Call PropBag.WriteProperty("DateType", m_intDateTypeMask, 0) 'Date
    Call PropBag.WriteProperty("EmailType", m_intEmailMask, 0) 'Email
    Call PropBag.WriteProperty("IZIPAllowSpaces", m_blnSpaces, True) 'IntlZip
    Call PropBag.WriteProperty("IPAllowPoints", m_blnPoints, True) 'IP Address
    Call PropBag.WriteProperty("PhoneMaskType", m_strMask, "(&&&) &&&-&&&& X&&&&") 'Phone Mask
    Call PropBag.WriteProperty("AreaCode", m_blnArea, True) 'Phone
    Call PropBag.WriteProperty("Parenthesis", m_blnPars, True) 'Phone
    Call PropBag.WriteProperty("PhnSpaces", m_blnSpc, True) 'Phone
    Call PropBag.WriteProperty("PhnDashes", m_blnDash, True) 'Phone
    Call PropBag.WriteProperty("Extension", m_blnExt, True) 'Phone
    Call PropBag.WriteProperty("SSNDashes", m_blnDashes, True) 'SSN
    Call PropBag.WriteProperty("ZipOnlyFive", m_blnOnlyFive, True) 'Zip
    Call PropBag.WriteProperty("MaskType", intMaskType, 0)
    Call PropBag.WriteProperty("AllowNegative", m_blnAllowNegative, False) ' Whole Numbers
    Call PropBag.WriteProperty("UsePrecision", m_blnUsePrecision, True)
    Call PropBag.WriteProperty("Precision", m_intPrecision, 2)

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtCBR,txtCBR,-1,WhatsThisHelpID
Public Property Get WhatsThisHelpID() As Long
    WhatsThisHelpID = txtCBR.WhatsThisHelpID
End Property

Public Property Let WhatsThisHelpID(ByVal New_WhatsThisHelpID As Long)
   PropertyChanged "WhatsThisHelpID"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ValidateControls
Public Sub ValidateControls()
   UserControl.ValidateControls
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtCBR,txtCBR,-1,ToolTipText
Public Property Get ToolTipText() As String
     ToolTipText = txtCBR.ToolTipText
End Property
Public Property Let ToolTipText(ByVal New_ToolTipText As String)
   txtCBR.ToolTipText() = New_ToolTipText
   PropertyChanged "ToolTipText"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,TextWidth
Public Function TextWidth(ByVal Str As String) As Single
   TextWidth = UserControl.TextWidth(Str)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,TextHeight
Public Function TextHeight(ByVal Str As String) As Single
   TextHeight = UserControl.TextHeight(Str)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Size
Public Sub Size(ByVal Width As Single, ByVal Height As Single)
   UserControl.Size Width, Height
End Sub

Private Sub UserControl_Show()
   RaiseEvent Show
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtCBR,txtCBR,-1,SelText
Public Property Get SelText() As String
   SelText = txtCBR.SelText
End Property
Public Property Let SelText(ByVal New_SelText As String)
   txtCBR.SelText() = New_SelText
   PropertyChanged "SelText"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtCBR,txtCBR,-1,SelStart
Public Property Get SelStart() As Long
   SelStart = txtCBR.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
   'txtCBR.SelStart() = New_SelStart
   txtCBR.SelStart() = New_SelStart
   PropertyChanged "SelStart"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtCBR,txtCBR,-1,SelLength
Public Property Get SelLength() As Long
   'SelLength = txtCBR.SelLength
   SelLength = txtCBR.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
   'txtCBR.SelLength() = New_SelLength
   txtCBR.SelLength() = New_SelLength
   PropertyChanged "SelLength"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtCBR,txtCBR,-1,ScrollBars
Public Property Get ScrollBars() As Integer
   ScrollBars = txtCBR.ScrollBars
End Property
Public Sub PopupMenu(ByVal Menu As Object, Optional ByVal Flags As Variant, Optional ByVal X As Variant, Optional ByVal Y As Variant, Optional ByVal DefaultMenu As Variant)
   UserControl.PopupMenu Menu, Flags, X, Y, DefaultMenu
End Sub

'The Underscore following "Point" is necessary because it
'is a Reserved Word in VBA.
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Point
Public Function Point(X As Single, Y As Single) As Long
   Point = UserControl.Point(X, Y)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Picture
Public Property Get Picture() As Picture
   Set Picture = UserControl.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
   Set UserControl.Picture = New_Picture
   PropertyChanged "Picture"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtCBR,txtCBR,-1,PasswordChar
Public Property Get PasswordChar() As String
   PasswordChar = txtCBR.PasswordChar
End Property

Public Property Let PasswordChar(ByVal New_PasswordChar As String)
   txtCBR.PasswordChar() = New_PasswordChar
   PropertyChanged "PasswordChar"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,PaletteMode
Public Property Get PaletteMode() As Integer
   PaletteMode = UserControl.PaletteMode
End Property

Public Property Let PaletteMode(ByVal New_PaletteMode As Integer)
   UserControl.PaletteMode() = New_PaletteMode
   PropertyChanged "PaletteMode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Palette
Public Property Get Palette() As Picture
   Set Palette = UserControl.Palette
End Property

Public Property Set Palette(ByVal New_Palette As Picture)
   Set UserControl.Palette = New_Palette
   PropertyChanged "Palette"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,PaintPicture
Public Sub PaintPicture(ByVal Picture As Picture, ByVal X1 As Single, ByVal Y1 As Single, Optional ByVal Width1 As Variant, Optional ByVal Height1 As Variant, Optional ByVal X2 As Variant, Optional ByVal Y2 As Variant, Optional ByVal Width2 As Variant, Optional ByVal Height2 As Variant, Optional ByVal Opcode As Variant)
   UserControl.PaintPicture Picture, X1, Y1, Width1, Height1, X2, Y2, Width2, Height2, Opcode
End Sub

Private Sub UserControl_Paint()
   RaiseEvent Paint
End Sub

Private Sub UserControl_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
   RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

Private Sub UserControl_OLESetData(Data As DataObject, DataFormat As Integer)
   RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub UserControl_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
   RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,OLEDropMode
Public Property Get OLEDropMode() As Integer
   OLEDropMode = UserControl.OLEDropMode
End Property

Public Property Let OLEDropMode(ByVal New_OLEDropMode As Integer)
   UserControl.OLEDropMode() = New_OLEDropMode
   PropertyChanged "OLEDropMode"
End Property

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
   RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, Y, State)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtCBR,txtCBR,-1,OLEDragMode
Public Property Get OLEDragMode() As Integer
   OLEDragMode = txtCBR.OLEDragMode
End Property

Public Property Let OLEDragMode(ByVal New_OLEDragMode As Integer)
   txtCBR.OLEDragMode() = New_OLEDragMode
   PropertyChanged "OLEDragMode"
End Property

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
   RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,OLEDrag
Public Sub OLEDrag()
   UserControl.OLEDrag
End Sub

Private Sub UserControl_OLECompleteDrag(Effect As Long)
   RaiseEvent OLECompleteDrag(Effect)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtCBR,txtCBR,-1,MultiLine
Public Property Get MultiLine() As Boolean
   MultiLine = txtCBR.MultiLine
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As Integer
   MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
   UserControl.MousePointer() = New_MousePointer
   PropertyChanged "MousePointer"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
   Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
   Set UserControl.MouseIcon = New_MouseIcon
   PropertyChanged "MouseIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtCBR,txtCBR,-1,MaxLength
Public Property Get MaxLength() As Long
    MaxLength = txtCBR.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
    txtCBR.MaxLength() = New_MaxLength
    PropertyChanged "MaxLength"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtCBR,txtCBR,-1,Locked
Public Property Get Locked() As Boolean
   Locked = txtCBR.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
   txtCBR.Locked() = New_Locked
   PropertyChanged "Locked"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hwnd() As Long
   hwnd = UserControl.hwnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtCBR,txtCBR,-1,HideSelection
Public Property Get HideSelection() As Boolean
   HideSelection = txtCBR.HideSelection
End Property

Private Sub UserControl_Hide()
   RaiseEvent Hide
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hDC
Public Property Get hDC() As Long
   hDC = UserControl.hDC
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,HasDC
Public Property Get HasDC() As Boolean
   HasDC = UserControl.HasDC
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtCBR,txtCBR,-1,FontUnderline
Public Property Get FontUnderline() As Boolean
   FontUnderline = txtCBR.FontUnderline
End Property

Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
   txtCBR.FontUnderline() = New_FontUnderline
   PropertyChanged "FontUnderline"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FontTransparent
Public Property Get FontTransparent() As Boolean
   FontTransparent = UserControl.FontTransparent
End Property

Public Property Let FontTransparent(ByVal New_FontTransparent As Boolean)
   UserControl.FontTransparent() = New_FontTransparent
   PropertyChanged "FontTransparent"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtCBR,txtCBR,-1,FontStrikethru
Public Property Get FontStrikethru() As Boolean
   FontStrikethru = txtCBR.FontStrikethru
End Property

Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
   txtCBR.FontStrikethru() = New_FontStrikethru
   PropertyChanged "FontStrikethru"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtCBR,txtCBR,-1,FontSize
Public Property Get FontSize() As Single
   FontSize = txtCBR.FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
   txtCBR.FontSize = New_FontSize
   PropertyChanged "FontSize"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtCBR,txtCBR,-1,FontName
Public Property Get FontName() As String
   FontName = txtCBR.FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
   txtCBR.FontName = New_FontName
   PropertyChanged "FontName"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtCBR,txtCBR,-1,FontItalic
Public Property Get FontItalic() As Boolean
   FontItalic = txtCBR.FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
   txtCBR.FontItalic() = New_FontItalic
   PropertyChanged "FontItalic"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtCBR,txtCBR,-1,FontBold
Public Property Get FontBold() As Boolean
   FontBold = txtCBR.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
   txtCBR.FontBold = New_FontBold
   PropertyChanged "FontBold"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FillColor
Public Property Get FillColor() As OLE_COLOR
   FillColor = UserControl.FillColor
End Property

Public Property Let FillColor(ByVal New_FillColor As OLE_COLOR)
   UserControl.FillColor() = New_FillColor
   PropertyChanged "FillColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Controls
Public Property Get Controls() As Object
   Set Controls = UserControl.Controls
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ContainerHwnd
Public Property Get ContainerHwnd() As Long
   ContainerHwnd = UserControl.ContainerHwnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Cls
Public Sub Cls()
   UserControl.Cls
End Sub

Private Sub txtCBR_Change()
   RaiseEvent Change
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtCBR,txtCBR,-1,CausesValidation
Public Property Get CausesValidation() As Boolean
   CausesValidation = txtCBR.CausesValidation
End Property

Public Property Let CausesValidation(ByVal New_CausesValidation As Boolean)
   txtCBR.CausesValidation() = New_CausesValidation
   PropertyChanged "CausesValidation"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtCBR,txtCBR,-1,Appearance
Public Property Get Appearance() As mAppearance
   Appearance = txtCBR.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As mAppearance)
   txtCBR.Appearance() = New_Appearance
   PropertyChanged "Appearance"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ActiveControl
Public Property Get ActiveControl() As Object
   Set ActiveControl = UserControl.ActiveControl
End Property

Private Sub UserControl_InitProperties()
   RaiseEvent InitProperties
End Sub
Private Sub txtCBR_Validate(Cancel As Boolean)
   RaiseEvent Validate(Cancel)
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As mBorderStyle
   BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As mBorderStyle)
   UserControl.BorderStyle() = New_BorderStyle
   PropertyChanged "BorderStyle"
End Property

Public Sub RefreshProperties(strProperty As String)
    PropertyChanged strProperty
End Sub

Public Property Get IPAllowPoints() As IPMaskType
    IPAllowPoints = m_blnPoints
End Property

Public Property Let IPAllowPoints(blnAllowPnts As IPMaskType)
    m_blnPoints = blnAllowPnts
    PropertyChanged "IPAllowPoints"
End Property

Public Property Get CurrencySign() As Boolean
    CurrencySign = m_blnSign
End Property
Public Property Let CurrencySign(blnAllowSign As Boolean)
    m_blnSign = blnAllowSign
    PropertyChanged "CurrencySign"
End Property

Public Property Get AllowAlpha() As Boolean
    AllowAlpha = m_blnAlpha
End Property

Public Property Let AllowAlpha(blnAllowAlpha As Boolean)
    m_blnAlpha = blnAllowAlpha
    PropertyChanged "AllowAlpha"
End Property

Public Property Get AllowNumeric() As Boolean
    AllowNumeric = m_blnNumeric
End Property

Public Property Let AllowNumeric(blnAllowNumber As Boolean)
    m_blnNumeric = blnAllowNumber
    PropertyChanged "AllowNumeric"
End Property
Public Property Get CharAllowed() As String
    CharAllowed = m_strCharAllowed
End Property
Public Property Let CharAllowed(strCharAllowed As String)
    m_strCharAllowed = strCharAllowed
    PropertyChanged "CharAllowed"
End Property
Public Property Get MaxAllowed() As Integer
    MaxAllowed = m_intMaxAllowed
End Property
Public Property Let MaxAllowed(intMaxAllowed As Integer)
    m_intMaxAllowed = intMaxAllowed
    PropertyChanged "MaxAllowed"
End Property
Public Property Get DateTypeMask() As DateMaskType
    DateTypeMask = m_intDateTypeMask
End Property
Public Property Let DateTypeMask(intDateTypeMask As DateMaskType)
    m_intDateTypeMask = intDateTypeMask
    PropertyChanged "DateTypeMask"
End Property
Public Property Get EmailType() As EmailMaskType
    EmailType = m_intEmailMask
End Property
Public Property Let EmailType(intEmailMask As EmailMaskType)
    m_intEmailMask = intEmailMask
    PropertyChanged "EmailType"
End Property
Public Property Get IZIPAllowSpaces() As Boolean
    IZIPAllowSpaces = m_blnSpaces
End Property
Public Property Let IZIPAllowSpaces(blnAllowSpaces As Boolean)
    m_blnSpaces = blnAllowSpaces
    PropertyChanged "IZIPAllowSpaces"
End Property
Public Property Get AreaCode() As Boolean
    AreaCode = m_blnArea
End Property
Public Property Let AreaCode(blnAreaCode As Boolean)
    m_blnArea = blnAreaCode
    PropertyChanged "AreaCode"
End Property
Public Property Get Parenthesis() As Boolean
    Parenthesis = m_blnPars
End Property
Public Property Let Parenthesis(blnPars As Boolean)
    m_blnPars = blnPars
    PropertyChanged "Parenthesis"
End Property
Public Property Get PhnSpaces() As Boolean
    PhnSpaces = m_blnSpc
End Property
Public Property Let PhnSpaces(blnSpaces As Boolean)
    m_blnSpc = blnSpaces
    PropertyChanged "PhnSpaces"
End Property
Public Property Get PhnDashes() As Boolean
    PhnDashes = m_blnDash
End Property
Public Property Let PhnDashes(blnDashes As Boolean)
    m_blnDash = blnDashes
    PropertyChanged "PhnDashes"
End Property
Public Property Get Extension() As Boolean
    Extension = m_blnExt
End Property
Public Property Let Extension(blnExtension As Boolean)
    m_blnExt = blnExtension
    PropertyChanged "Extension"
End Property
Public Property Get SSNDashes() As Boolean
    SSNDashes = m_blnDashes
End Property
Public Property Let SSNDashes(blnDashes As Boolean)
    m_blnDashes = blnDashes
    PropertyChanged "SSNDashes"
End Property
Public Property Get ZipOnlyFive() As Boolean
    ZipOnlyFive = m_blnOnlyFive
End Property
Public Property Let ZipOnlyFive(blnOnlyFive As Boolean)
    m_blnOnlyFive = blnOnlyFive
    PropertyChanged "ZipOnlyFive"
End Property
Public Property Get PhoneMaskType() As String
    PhoneMaskType = m_strMask
End Property
Public Property Let PhoneMaskType(strPhoneMaskType As String)
    m_strMask = strPhoneMaskType
    PropertyChanged "PhoneMaskType"
End Property
Public Property Get AllowNegative() As AllowNegatives
    AllowNegative = m_blnAllowNegative
End Property
Public Property Let AllowNegative(ByVal bAllowNegatives As AllowNegatives)
    m_blnAllowNegative = bAllowNegatives
    PropertyChanged "AllowNegative"
End Property
Public Property Get UsePrecision() As UsePrecision
    UsePrecision = m_blnUsePrecision
End Property
Public Property Let UsePrecision(ByVal bUsePrecision As UsePrecision)
    m_blnUsePrecision = bUsePrecision
    PropertyChanged "UsePrecision"
End Property
Public Property Get Precision() As Integer
    Precision = m_intPrecision
End Property
Public Property Let Precision(iPrecision As Integer)
    m_intPrecision = iPrecision
    PropertyChanged "Precision"
End Property
Public Property Get AllCaps() As Boolean
    AllCaps = m_blnAllCaps
End Property
Public Property Let AllCaps(ByVal bAllCaps As Boolean)
    m_blnAllCaps = bAllCaps
    PropertyChanged "AllCaps"
End Property

