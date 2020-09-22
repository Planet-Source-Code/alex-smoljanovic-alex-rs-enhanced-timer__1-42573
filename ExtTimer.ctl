VERSION 5.00
Begin VB.UserControl ExtTimer 
   ClientHeight    =   390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   390
   InvisibleAtRuntime=   -1  'True
   Picture         =   "ExtTimer.ctx":0000
   ScaleHeight     =   390
   ScaleWidth      =   390
   ToolboxBitmap   =   "ExtTimer.ctx":06EA
   Begin VB.Timer tmrTimer 
      Left            =   240
      Top             =   360
   End
End
Attribute VB_Name = "ExtTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***********************************************************************
'This component was explicitly developed for
'PSC(Planet Source Code) Users as an Open Source Project.
'This code is property of it's author.
'
'If you compile this User Control you MAY redistribute it.
'
'Alex Smoljanovic, Salex Software (c) 2001-2003
'salex_software@shaw.ca
'***********************************************************************


Private Declare Function GetTickCount Lib "kernel32" () As Long
'Retrieves the amount of milliseconds representing the duration of time the system has been running.

Private tmpTC& 'Temporary Tick Count, this variable is used in comparison to future tick counts...
Dim m_Hours&, m_Minutes&, m_Seconds&, m_Enabled As Boolean, m_Interval&
'variables representing this components properties...

Public Event Timer()
'This even is raised when the specified interval has elapsed...

Private Sub tmrTimer_Timer()
 If EventsFrozen = True Then Exit Sub
 'Function EventsFrozen returns a boolean value;
 'Returns true if the client application is ignoring events(user is editing the form on which this control is contained)
  If (GetTickCount - tmpTC) >= m_Interval Then
  'Determine the amount of elapsed time since the this user control last enabled this control, or since the last time the Timer event was raised
  'If the elapsed time evaluates to variable m_Interval(m_Interval's value is calculated based upon the Hours, Minutes, and Seconds property values) then...
   tmrTimer = False 'Disable the Timer control
    RaiseEvent Timer 'Raise the Timer Event
     EnableTmr 'This will initialize tmpTC with the current tick count, and will re-enable the timer control
  End If
End Sub

Private Sub EnableTmr()
tmpTC = GetTickCount 'Retrieve elapsed time in milliseconds since this system started
 If m_Interval * 0.1 > 65535 Then
 'If 10% of variable m_Interval(m_Interval evaluates to the total milliseconds of properties Hours, Minutes, and Seconds combined) is greater than the maximum interval allowed by the Timer Control then...
  tmrTimer.Interval = 65535
  'Set the timers interval property to its maximum value
 Else
  tmrTimer.Interval = m_Interval * 0.1
  'Set the timers interval property to 10% of m_Interval's value
 End If
  tmrTimer.Enabled = m_Enabled
  'Set the timer's enabled property to the value of m_Enabled(m_Enabled evaluates to this controls Enabled property)
End Sub

Private Sub CalcInterval()
Dim Hours&, Minutes&, Seconds&
'dimensionalize Hours as long type, Minutes as long type, Seconds as long type
 Hours = ((m_Hours * 60) * 60) * 1000
 'initialize Hours with the amount of milliseconds which is equal to the amount of hours specified by the Hours property of this control
  Minutes = (m_Minutes * 60) * 1000
  'initialize Minutes with the amount of milliseconds which is equal to the amount of minutes specified by the Minutes property of this control
   Seconds = m_Seconds * 1000
   'initialize Seconds with the amount of milliseconds which is equal to the amount of seconds specified by the Seconds property of this control
    m_Interval = (Hours + Minutes) + Seconds
    'initialize variable m_Interval's value to the combination of variables Hours, Minutes, and Seconds which all represent time in milliseconds
End Sub

Private Sub UserControl_Initialize()
 CalcInterval 'See CalcInterval for more info...
  EnableTmr 'See EnableTmr for more info...
End Sub

Private Sub UserControl_Resize()
 UserControl.Width = 390
  UserControl.Height = 390
  'Set this object's width and height properties to their initial values
End Sub

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
'This procedure is called when a client is trying to determine this controls Enabled property value
 Enabled = m_Enabled 'return the variable m_Enabled's value
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
 m_Enabled = New_Enabled 'initialize m_Enabled with the new_Enabled's argument value
  PropertyChanged "Enabled" 'This method is called to notify this components container(parent) that the specified property's value has been modified..
   EnableTmr 'See EnableTmr for more info...
End Property

Public Property Get Interval() As Long
 Interval = m_Interval
 'This properties Let procedure is not included so that this property is Read-Only during all states(Run time, Design time)
End Property

Private Sub UserControl_InitProperties()
 m_Enabled = False
  m_Hours = 0
   m_Minutes = 0
    m_Seconds = 0
     m_Interval = 0
     'Initialize the properties (variable representatives)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
 m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
  m_Hours = PropBag.ReadProperty("Hours", m_def_Hours)
   m_Minutes = PropBag.ReadProperty("Minutes", m_def_Minutes)
    m_Seconds = PropBag.ReadProperty("Seconds", m_def_Seconds)
     m_Interval = PropBag.ReadProperty("Interval", m_def_Interval)
     'This procedure will load the property values from storage
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
 Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
  Call PropBag.WriteProperty("Hours", m_Hours, m_def_Hours)
   Call PropBag.WriteProperty("Minutes", m_Minutes, m_def_Minutes)
    Call PropBag.WriteProperty("Seconds", m_Seconds, m_def_Seconds)
     Call PropBag.WriteProperty("Interval", m_Interval, m_def_Interval)
     'Save the property variables in storage
End Sub
     'The way in which a controls save information is almost identical to the way in which any executable pe image file stores resources(strings,accelerators,bitmaps,icons,dialogs...) except that this information will be used by VB to write this information in its container when its container is being compiled
     'So basically, when the project who contains this control is compiled by VB, VB will actually write this information in the object file(compiled object[.frm file or otherwise]) prior to the project being fully compiled so that this information can be used to set the components properties when its parent is executed.

Public Property Get Hours() As Long
Attribute Hours.VB_Description = "Specified the amount of Hours in the Timers Interval"
 Hours = m_Hours
End Property '...

Public Property Let Hours(ByVal New_Hours As Long)
 m_Hours = New_Hours
  PropertyChanged "Hours"
   CalcInterval
End Property '...

Public Property Get Minutes() As Long
Attribute Minutes.VB_Description = "Specified the amount of Minutes in the Timers Interval"
 Minutes = m_Minutes
End Property '...

Public Property Let Minutes(ByVal New_Minutes As Long)
 m_Minutes = New_Minutes
  PropertyChanged "Minutes"
   CalcInterval
End Property '...

Public Property Get Seconds() As Long
Attribute Seconds.VB_Description = "Specified the amount of Seconds in the Timers Interval"
 Seconds = m_Seconds
End Property '...

Public Property Let Seconds(ByVal New_Seconds As Long)
 m_Seconds = New_Seconds
  PropertyChanged "Seconds"
   CalcInterval
End Property '...

