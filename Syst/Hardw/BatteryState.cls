VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BatteryState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum PowerInformationLevel
    SystemPowerPolicyAc = 0
    SystemPowerPolicyDc
    VerifySystemPolicyAc
    VerifySystemPolicyDc
    SystemPowerCapabilities
    SystemBatteryState
    SystemPowerStateHandler
    ProcessorStateHandler
    SystemPowerPolicyCurrent
    AdministratorPowerPolicy
    SystemReserveHiberFile
    ProcessorInformation
    SystemPowerInformation
    ProcessorStateHandler2
    LastWakeTime
    LastSleepTime
    SystemExecutionState
    SystemPowerStateNotifyHandler
    ProcessorPowerPolicyAc
    ProcessorPowerPolicyDc
    VerifyProcessorPowerPolicyAc
    VerifyProcessorPowerPolicyDc
    ProcessorPowerPolicyCurrent
    SystemPowerStateLogging
    SystemPowerLoggingEntry
    SetPowerSettingValue
    NotifyUserPowerSetting
    PowerInformationLevelUnused0
    PowerInformationLevelUnused1
    SystemVideoState
    TraceApplicationPowerMessage
    TraceApplicationPowerMessageEnd
    ProcessorPerfStates
    ProcessorIdleStates
    ProcessorCap
    SystemWakeSource
    SystemHiberFileInformation
    TraceServicePowerMessage
    ProcessorLoad
    PowerShutdownNotification
    MonitorCapabilities
    SessionPowerInit
    SessionDisplayState
    PowerRequestCreate
    PowerRequestAction
    GetPowerRequestList
    ProcessorInformationEx
    NotifyUserModeLegacyPowerEvent
    GroupPark
    ProcessorIdleDomains
    WakeTimerList
    SystemHiberFileSize
    PowerInformationLevelMaximum
End Enum

Private Type SystemBatteryState
    IsACOnLine         As Byte ' [MarshalAs(UnmanagedType.I1)]
    IsBatteryPresent   As Byte ' [MarshalAs(UnmanagedType.I1)]
    IsCharging         As Byte ' [MarshalAs(UnmanagedType.I1)]
    IsDischarging      As Byte ' [MarshalAs(UnmanagedType.I1)]
    Spare1             As Byte
    spare2             As Byte
    spare3             As Byte
    Spare4             As Byte
    MaxCapacity        As Long ' public uint
    RemainingCapacity  As Long ' public uint
    Rate               As Long ' public uint
    EstimatedTime      As Long ' public uint
    DefaultAlert1      As Long ' public uint
    DefaultAlert2      As Long ' public uint
End Type

Private Declare Function CallNtPowerInformation Lib "powrprof" (ByVal informationLevel As PowerInformationLevel, ByVal pinputBuffer As Long, ByVal inputBufferSize As Long, ByRef outputBuffer As SystemBatteryState, ByVal outputBufferSize As Long) As Long

'// Microsoft.WindowsAPICodePack.ApplicationServices.BatteryState
Private m_SystemBatteryState As SystemBatteryState

Private Sub Class_Initialize()
'        PowerManagementNativeMethods.SystemBatteryState systemBatteryState = Power.GetSystemBatteryState();
'        if (!systemBatteryState.BatteryPresent)
'        {
'            throw new InvalidOperationException(LocalizedMessages.PowerManagerBatteryNotPresent);
'        }
    Recall False
'    m_SystemBatteryState = MPower.GetSystemBatteryState
'    If Not CBool(m_SystemBatteryState.boolBatteryPresent) Then
'        'Err.Description = "InvalidOperationException(LocalizedMessages.PowerManagerBatteryNotPresent)"
'        'Err.Raise 0
'        MsgBox "InvalidOperationException( LocalizedMessages.PowerManagerBatteryNotPresent )"
'    End If
'        ACOnline = systemBatteryState.AcOnLine;
'        MaxCharge = (int)systemBatteryState.MaxCapacity;
'        CurrentCharge = (int)systemBatteryState.RemainingCapacity;
'        ChargeRate = (int)systemBatteryState.Rate;
'        uint estimatedTime = systemBatteryState.EstimatedTime;
'        if (estimatedTime != uint.MaxValue)
'        {
'            EstimatedTimeRemaining = new TimeSpan(0, 0, (int)estimatedTime);
'        }
'        Else
'        {
'            EstimatedTimeRemaining = TimeSpan.MaxValue;
'        }
'        SuggestedCriticalBatteryCharge = (int)systemBatteryState.DefaultAlert1;
'        SuggestedBatteryWarningCharge = (int)systemBatteryState.DefaultAlert2;

End Sub

Public Sub Recall(Optional ByVal loud As Boolean = False)
    m_SystemBatteryState = GetSystemBatteryState
    If Not CBool(m_SystemBatteryState.IsBatteryPresent) Then
        'Err.Description = "InvalidOperationException(LocalizedMessages.PowerManagerBatteryNotPresent)"
        'Err.Raise 0
        If loud Then MsgBox "InvalidOperationException( LocalizedMessages.PowerManagerBatteryNotPresent )"
    End If
End Sub

'internal static class Power
'{
'internal static PowerManagementNativeMethods.SystemBatteryState GetSystemBatteryState()
'{
'    if (PowerManagementNativeMethods.CallNtPowerInformation(PowerManagementNativeMethods.PowerInformationLevel.SystemBatteryState, IntPtr.Zero, 0u, out PowerManagementNativeMethods.SystemBatteryState outputBuffer, (uint)Marshal.SizeOf(typeof(PowerManagementNativeMethods.SystemBatteryState))) == 3221225506u)
'    {
'        throw new UnauthorizedAccessException(LocalizedMessages.PowerInsufficientAccessBatteryState);
'    }
'    return outputBuffer;
'}
Private Function GetSystemBatteryState() As SystemBatteryState
    Dim hr As Long: hr = CallNtPowerInformation(PowerInformationLevel.SystemBatteryState, 0, 0, GetSystemBatteryState, LenB(GetSystemBatteryState))
    If hr = &HC0000022 Then Err.Raise hr, , "UnauthorizedAccessException(LocalizedMessages.PowerInsufficientAccessBatteryState)"
End Function

Public Property Get IsACOnLine() As Boolean
    IsACOnLine = m_SystemBatteryState.IsACOnLine
End Property

Public Property Get IsBatteryPresent() As Boolean
    IsBatteryPresent = m_SystemBatteryState.IsBatteryPresent
End Property

Public Property Get IsCharging() As Boolean
    IsCharging = m_SystemBatteryState.IsCharging
End Property

Public Property Get IsDischarging() As Boolean
    IsDischarging = m_SystemBatteryState.IsDischarging
End Property

Public Property Get ChargeRate() As Long
    ChargeRate = m_SystemBatteryState.Rate
End Property

Public Property Get CurrentCharge() As Long
    CurrentCharge = m_SystemBatteryState.RemainingCapacity
End Property

Public Property Get EstimatedTimeRemaining() As Long 'Date
    EstimatedTimeRemaining = m_SystemBatteryState.EstimatedTime
End Property

Public Property Get MaxCharge() As Long
    MaxCharge = m_SystemBatteryState.MaxCapacity
End Property

Public Property Get SuggestedBatteryWarningCharge() As Long
    SuggestedBatteryWarningCharge = m_SystemBatteryState.DefaultAlert2
End Property

Public Property Get SuggestedCriticalBatteryCharge() As Long
    SuggestedCriticalBatteryCharge = m_SystemBatteryState.DefaultAlert1
End Property

Public Property Get RemainingCapacity() As Long ' public uint
    RemainingCapacity = m_SystemBatteryState.RemainingCapacity
End Property

Public Property Get MaxCapacity() As Long      ' public uint
    MaxCapacity = m_SystemBatteryState.MaxCapacity
End Property

Private Function Secs_ToStr(ByVal secs As Long) As String
    Dim est As Long: est = secs
    Dim sec As Long: sec = est Mod 60
    Dim min As Long: min = (est \ 60) Mod 60
    Dim hor As Long: hor = Int(est \ 3600)
    Secs_ToStr = CStr(hor) & ":" & Format(min, "00") & ":" & Format(sec, "00")
End Function

Public Property Get BatteryLifePercent() As Long
    'Dim aSystemBatteryState As SystemBatteryState: aSystemBatteryState = MPower.GetSystemBatteryState
    If Not m_SystemBatteryState.IsBatteryPresent Then
        'Err.Raise -1, , "InvalidOperationException(LocalizedMessages.PowerManagerBatteryNotPresent)"
        MsgBox "InvalidOperationException(LocalizedMessages.PowerManagerBatteryNotPresent)"
    End If
    BatteryLifePercent = m_SystemBatteryState.RemainingCapacity / m_SystemBatteryState.MaxCapacity * 100
End Property

'[BatteryStateStringRepresentation,
'ACOnline: {1}
'{0}Max Charge: {2} mWh
'{0}Current Charge: {3} mWh
'{0}Discharge Rate: {4} mWh
'{0}Estimated Time Remaining: {5}
'{0}Suggested Critical Battery Charge: {6} mWh
'{0}Suggested Battery Warning Charge: {7} mWh
'{0}]

Public Function ToStr() As String
'return string.Format( CultureInfo.InvariantCulture, _
'                      LocalizedMessages.BatteryStateStringRepresentation, _
'                      Environment.NewLine, _
'                      ACOnline, _
'                      MaxCharge, _
'                      CurrentCharge, _
'                      ChargeRate, _
'                      EstimatedTimeRemaining, _
'                      SuggestedCriticalBatteryCharge, _
'                      SuggestedBatteryWarningCharge)
    Dim s As String
        s = "BatteryState" & vbCrLf & _
            "{" & vbCrLf & _
            "    IsACOnline                     = " & Me.IsACOnLine & ";" & vbCrLf & _
            "    IsBatteryPresent               = " & Me.IsBatteryPresent & ";" & vbCrLf & _
            "    IsCharging                     = " & Me.IsCharging & ";" & vbCrLf & _
            "    IsDischarging                  = " & Me.IsDischarging & ";" & vbCrLf & _
            "    Max Charge                     = " & Me.MaxCharge & "mWh;" & vbCrLf
If Me.MaxCharge <> 0 Then
    s = s & "    CurrentCharge                  = " & Me.CurrentCharge & " (" & CLng(Me.CurrentCharge / Me.MaxCharge * 100) & "%)" & "mWh;" & vbCrLf
End If
    s = s & "    ChargeRate                     = " & Me.ChargeRate & "mWh;" & vbCrLf & _
            "    RemainingCapacity              = " & Me.RemainingCapacity & "mWh;" & vbCrLf & _
            "    MaxCapacity                    = " & Me.MaxCapacity & "mWh;" & vbCrLf & _
            "    EstimatedTimeRemaining         = " & Secs_ToStr(Me.EstimatedTimeRemaining) & ";" & vbCrLf & _
            "    SuggestedCriticalBatteryCharge = " & Me.SuggestedCriticalBatteryCharge & "mWh;" & vbCrLf & _
            "    SuggestedBatteryWarningCharge  = " & Me.SuggestedBatteryWarningCharge & "mWh;" & vbCrLf & _
            "}"
    ToStr = s
End Function


'/// <summary>A snapshot of the state of the battery.</summary>
'public class BatteryState
'{
'    /// <summary>Gets a value that indicates whether the battery charger is operating on external power.</summary>
'    /// <value>A <see cref="T:System.Boolean" /> value. <b>True</b> indicates the battery charger is operating on AC power.</value>
'    public bool ACOnline
'    {
'        get;
'        private set;
'    }
'
'    /// <summary>Gets the rate of discharge for the battery (in mW).</summary>
'    /// <remarks>
'    /// If plugged in, fully charged: DischargeRate = 0. If plugged in, charging: DischargeRate = positive mW per hour. If unplugged:
'    /// DischargeRate = negative mW per hour.
'    /// </remarks>
'    /// <value>An <see cref="T:System.Int32" /> value.</value>
'    public int ChargeRate
'    {
'        get;
'        private set;
'    }
'
'    /// <summary>Gets the current charge of the battery (in mW).</summary>
'    /// <value>An <see cref="T:System.Int32" /> value.</value>
'    public int CurrentCharge
'    {
'        get;
'        private set;
'    }
'
'    /// <summary>Gets the estimated time remaining until the battery is empty.</summary>
'    /// <value>A <see cref="T:System.TimeSpan" /> object.</value>
'    public TimeSpan EstimatedTimeRemaining
'    {
'        get;
'        private set;
'    }
'
'    /// <summary>Gets the maximum charge of the battery (in mW).</summary>
'    /// <value>An <see cref="T:System.Int32" /> value.</value>
'    public int MaxCharge
'    {
'        get;
'        private set;
'    }
'
'    /// <summary>Gets the manufacturer's suggested battery charge level that should cause a warning to be sent to the user.</summary>
'    /// <value>An <see cref="T:System.Int32" /> value.</value>
'    public int SuggestedBatteryWarningCharge
'    {
'        get;
'        private set;
'    }
'
'    /// <summary>Gets the manufacturer's suggested battery charge level that should cause a critical alert to be sent to the user.</summary>
'    /// <value>An <see cref="T:System.Int32" /> value.</value>
'    public int SuggestedCriticalBatteryCharge
'    {
'        get;
'        private set;
'    }
'
'    internal BatteryState()
'    {
'        PowerManagementNativeMethods.SystemBatteryState systemBatteryState = Power.GetSystemBatteryState();
'        if (!systemBatteryState.BatteryPresent)
'        {
'            throw new InvalidOperationException(LocalizedMessages.PowerManagerBatteryNotPresent);
'        }
'        ACOnline = systemBatteryState.AcOnLine;
'        MaxCharge = (int)systemBatteryState.MaxCapacity;
'        CurrentCharge = (int)systemBatteryState.RemainingCapacity;
'        ChargeRate = (int)systemBatteryState.Rate;
'        uint estimatedTime = systemBatteryState.EstimatedTime;
'        if (estimatedTime != uint.MaxValue)
'        {
'            EstimatedTimeRemaining = new TimeSpan(0, 0, (int)estimatedTime);
'        }
'        Else
'        {
'            EstimatedTimeRemaining = TimeSpan.MaxValue;
'        }
'        SuggestedCriticalBatteryCharge = (int)systemBatteryState.DefaultAlert1;
'        SuggestedBatteryWarningCharge = (int)systemBatteryState.DefaultAlert2;
'    }
'
'    /// <summary>Generates a string that represents this <b>BatteryState</b> object.</summary>
'    /// <returns>A <see cref="T:System.String" /> representation of this object's current state.</returns>
'    public override string ToString()
'    {
'        return string.Format(CultureInfo.InvariantCulture, LocalizedMessages.BatteryStateStringRepresentation, Environment.NewLine, ACOnline, MaxCharge, CurrentCharge, ChargeRate, EstimatedTimeRemaining, SuggestedCriticalBatteryCharge, SuggestedBatteryWarningCharge);
'    }
'}
