VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PowerCaps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum ESystemPowerState
    Unspecified
    Working
    Sleeping1
    Sleeping2
    Sleeping3
    Hibernate
    Shutdown
    Maximum
End Enum

Public Enum EPowerSource
    AC      '/// <summary> The computer is powered by an AC power source or a similar device, such as a laptop powered by a 12V automotive adapter.</summary>
    Battery '/// <summary> The computer is powered by a built-in battery. A battery has a limited amount of power; applications should conserve resources where possible. </summary>
    Ups     '/// <summary>The computer is powered by a short-term power source such as a UPS device.</summary>
End Enum

Public Enum EExecutionStates
    None = &H0&            '/// <summary>No state configured.</summary>
    SystemRequired = &H1&  '/// <summary>Forces the system to be in the working state by resetting the system idle timer.</summary>
    DisplayRequired = &H2& '/// <summary>Forces the display to be on by resetting the display idle timer.</summary>

    '/// <summary>
    '/// Enables away mode. This value must be specified with ES_CONTINUOUS. Away mode should be used only by media-recording and
    '/// media-distribution applications that must perform critical background processing on desktop computers while the computer appears
    '/// to be sleeping. See Remarks.
    '///
    '/// Windows Server 2003 and Windows XP/2000: ES_AWAYMODE_REQUIRED is not supported.
    '/// </summary>
    AwayModeRequired = &H40&

    '/// <summary>
    '/// Informs the system that the state being set should remain in effect until the next call that uses ES_CONTINUOUS and one of the
    '/// other state flags is cleared.
    '/// </summary>
    Continuous = -2147483647 'int.MinValue
End Enum

Private Type VBGuid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Public Enum EPowerPersonality
    PPUnknown       '/// <summary>The power personality Guid does not match a known value.</summary>
    HighPerformance '/// <summary>Power settings designed to deliver maximum performance at the expense of power consumption savings.</summary>
    PowerSaver      '/// <summary>Power settings designed consume minimum power at the expense of system performance and responsiveness.</summary>
    Automatic       '/// <summary>Power settings designed to balance performance and power consumption.</summary>
End Enum

Private Type BatteryReportingScale
    Granularity As Long
    Capacity    As Long
End Type

'https://docs.microsoft.com/en-us/windows/win32/api/winnt/ns-winnt-system_power_capabilities
Private Type SystemPowerCapabilities
    IsPowerButtonPresent      As Byte ' [MarshalAs(UnmanagedType.I1)]
    IsSleepButtonPresent      As Byte ' [MarshalAs(UnmanagedType.I1)]
    IsLidPresent              As Byte ' [MarshalAs(UnmanagedType.I1)]
    IsSystemS1                As Byte ' [MarshalAs(UnmanagedType.I1)]
    IsSystemS2                As Byte ' [MarshalAs(UnmanagedType.I1)]
    IsSystemS3                As Byte ' [MarshalAs(UnmanagedType.I1)]
    IsSystemS4                As Byte ' [MarshalAs(UnmanagedType.I1)]
    IsSystemS5                As Byte ' [MarshalAs(UnmanagedType.I1)]
    IsHiberFilePresent        As Byte ' [MarshalAs(UnmanagedType.I1)]
    IsFullWake                As Byte ' [MarshalAs(UnmanagedType.I1)]
    IsVideoDimPresent         As Byte ' [MarshalAs(UnmanagedType.I1)]
    IsApmPresent              As Byte ' [MarshalAs(UnmanagedType.I1)]
    IsUpsPresent              As Byte ' [MarshalAs(UnmanagedType.I1)]
    IsThermalControl          As Byte ' [MarshalAs(UnmanagedType.I1)]
    IsProcessorThrottle       As Byte ' [MarshalAs(UnmanagedType.I1)]
    ProcessorMinimumThrottle  As Byte
    ProcessorMaximumThrottle  As Byte
    IsFastSystemS4            As Byte ' [MarshalAs(UnmanagedType.I1)]
    spare2(1 To 3)            As Byte ' [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    IsDiskSpinDown            As Byte ' [MarshalAs(UnmanagedType.I1)]
    spare3(1 To 8)            As Byte ' public byte[] [MarshalAs(UnmanagedType.ByValArray, SizeConst = 8)]
    SystemBatteriesArePresent As Byte '[MarshalAs(UnmanagedType.I1)]
    BatteriesAreShortTerm     As Byte '[MarshalAs(UnmanagedType.I1)]
    BatteryScale(1 To 3)      As BatteryReportingScale '[MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    ACOnlineWake              As ESystemPowerState
    SoftLidWake               As ESystemPowerState
    RtcWake                   As ESystemPowerState
    MinimumDeviceWakeState    As ESystemPowerState
    DefaultLowLatencyWake     As ESystemPowerState
End Type

'[DllImport("powrprof.dll")]
Private Declare Function CallNtPowerInformation Lib "powrprof" (ByVal informationLevel As PowerInformationLevel, ByVal pinputBuffer As Any, ByVal inputBufferSize As Long, ByRef outputBuffer As SystemPowerCapabilities, ByVal outputBufferSize As Long) As Long                                                     'uint

Private m_All             As VBGuid
Private m_Automatic       As VBGuid
Private m_HighPerformance As VBGuid
Private m_PowerSaver      As VBGuid


Private m_SystemPowerCaps As SystemPowerCapabilities
Private m_BatteryState    As BatteryState

Private m_monitoronlock     As Object
Private m_isMonitorOn       As Boolean '? Nullable
Private m_monitorRequired   As Boolean
Private m_requestBlockSleep As Boolean

Private Sub Class_Initialize()
    m_SystemPowerCaps = GetSystemPowerCapabilities
    Set m_BatteryState = New BatteryState
    InitGuids
End Sub

Private Sub InitGuids()
                m_All = New_VBGuid(&H68A1E95E, &H13EA, &H41E1, 128, 17, 12, 73, 108, 164, 144, 176)
          m_Automatic = New_VBGuid(&H381B4222, &HF694, &H41F0, 150, 133, 255, 91, 178, 96, 223, 46)
    m_HighPerformance = New_VBGuid(&H8C5E7FDA, &HE8BF, &H4A96, 154, 133, 166, 226, 58, 140, 99, 92)
         m_PowerSaver = New_VBGuid(&HA1841308, &H3541, &H4FAB, 188, 129, 247, 21, 86, 242, 11, 74)
End Sub

Private Function New_VBGuid(ByVal Data1 As Long, ByVal Data2 As Integer, ByVal Data3 As Integer, _
                         ByVal Data40 As Byte, ByVal Data41 As Byte, ByVal Data42 As Byte, ByVal Data43 As Byte, _
                         ByVal Data44 As Byte, ByVal Data45 As Byte, ByVal Data46 As Byte, ByVal Data47 As Byte) As VBGuid
    With New_VBGuid
        .Data1 = Data1: .Data2 = Data2: .Data3 = Data3
        .Data4(0) = Data40: .Data4(1) = Data41: .Data4(2) = Data42: .Data4(3) = Data43
        .Data4(4) = Data44: .Data4(5) = Data45: .Data4(6) = Data46: .Data4(7) = Data47
    End With
End Function

Private Function VBGuid_IsEqual(g1 As VBGuid, g2 As VBGuid) As Boolean
    If g1.Data1 <> g2.Data1 Then Exit Function
    If g1.Data2 <> g2.Data2 Then Exit Function
    If g1.Data3 <> g2.Data3 Then Exit Function
    Dim i As Long
    For i = 0 To 7
        If g1.Data4(i) <> g2.Data4(i) Then Exit Function
    Next
End Function

'Private Function VBGuid_ToEnum(aGuid As VBGuid) As EPowerPersonality
'    Dim e As EPowerPersonality
'    If VBGuid_IsEqual(aGuid, m_HighPerformance) Then
'        e = EPowerPersonality.HighPerformance
'    ElseIf VBGuid_IsEqual(aGuid, m_PowerSaver) Then
'        e = EPowerPersonality.PowerSaver
'    ElseIf VBGuid_IsEqual(aGuid, m_Automatic) Then
'        e = EPowerPersonality.Automatic
'    Else
'        e = EPowerPersonality.PPUnknown
'    End If
'    VBGuid_ToEnum = e
'End Function
Private Function VBGuid_ToEnum(aGuid As VBGuid) As EPowerPersonality
    If VBGuid_IsEqual(aGuid, m_HighPerformance) Then VBGuid_ToEnum = EPowerPersonality.HighPerformance: Exit Function
    If VBGuid_IsEqual(aGuid, m_PowerSaver) Then VBGuid_ToEnum = EPowerPersonality.PowerSaver: Exit Function
    If VBGuid_IsEqual(aGuid, m_Automatic) Then VBGuid_ToEnum = EPowerPersonality.Automatic: Exit Function
    VBGuid_ToEnum = EPowerPersonality.PPUnknown
End Function

Public Property Get IsPowerButtonPresent() As Boolean
    IsPowerButtonPresent = m_SystemPowerCaps.IsPowerButtonPresent
End Property

Public Property Get IsSleepButtonPresent() As Boolean
    IsSleepButtonPresent = m_SystemPowerCaps.IsSleepButtonPresent
End Property

Public Property Get IsLidPresent() As Boolean
    IsLidPresent = m_SystemPowerCaps.IsLidPresent
End Property

Public Property Get IsSystemS1() As Boolean
    IsSystemS1 = m_SystemPowerCaps.IsSystemS1
End Property

Public Property Get IsSystemS2() As Boolean
    IsSystemS2 = m_SystemPowerCaps.IsSystemS2
End Property

Public Property Get IsSystemS3() As Boolean
    IsSystemS3 = m_SystemPowerCaps.IsSystemS3
End Property

Public Property Get IsSystemS4() As Boolean
    IsSystemS4 = m_SystemPowerCaps.IsSystemS4
End Property

Public Property Get IsSystemS5() As Boolean
    IsSystemS5 = m_SystemPowerCaps.IsSystemS5
End Property

Public Property Get IsHiberFilePresent() As Boolean
    IsHiberFilePresent = m_SystemPowerCaps.IsHiberFilePresent
End Property

Public Property Get IsFullWake() As Boolean
    IsFullWake = m_SystemPowerCaps.IsFullWake
End Property

Public Property Get IsVideoDimPresent() As Boolean
    IsVideoDimPresent = m_SystemPowerCaps.IsVideoDimPresent
End Property

Public Property Get IsApmPresent() As Boolean
    IsApmPresent = m_SystemPowerCaps.IsApmPresent
End Property

Public Property Get IsUpsPresent() As Boolean
    IsUpsPresent = m_SystemPowerCaps.IsUpsPresent
End Property

Public Property Get IsThermalControl() As Boolean
    IsThermalControl = m_SystemPowerCaps.IsThermalControl
End Property

Public Property Get IsProcessorThrottle() As Boolean
    IsProcessorThrottle = m_SystemPowerCaps.IsProcessorThrottle
End Property

Public Property Get ProcessorMinimumThrottle() As Boolean
    ProcessorMinimumThrottle = m_SystemPowerCaps.ProcessorMinimumThrottle
End Property

Public Property Get ProcessorMaximumThrottle() As Boolean
    ProcessorMaximumThrottle = m_SystemPowerCaps.ProcessorMaximumThrottle
End Property

Public Property Get IsFastSystemS4() As Boolean
    IsFastSystemS4 = m_SystemPowerCaps.IsFastSystemS4
End Property

Public Property Get IsDiskSpinDown() As Boolean
    IsDiskSpinDown = m_SystemPowerCaps.IsDiskSpinDown
End Property

Public Property Get SystemBatteriesArePresent() As Boolean
    SystemBatteriesArePresent = m_SystemPowerCaps.SystemBatteriesArePresent
End Property

Public Property Get BatteriesAreShortTerm() As Boolean
    BatteriesAreShortTerm = m_SystemPowerCaps.BatteriesAreShortTerm
End Property

Public Property Get ACOnlineWake() As Boolean
    ACOnlineWake = m_SystemPowerCaps.ACOnlineWake
End Property

Public Property Get SoftLidWake() As Boolean
    SoftLidWake = m_SystemPowerCaps.SoftLidWake
End Property

Public Property Get RtcWake() As Boolean
    RtcWake = m_SystemPowerCaps.RtcWake
End Property

Public Property Get MinimumDeviceWakeState() As Boolean
    MinimumDeviceWakeState = m_SystemPowerCaps.MinimumDeviceWakeState
End Property

Public Property Get DefaultLowLatencyWake() As Boolean
    DefaultLowLatencyWake = m_SystemPowerCaps.DefaultLowLatencyWake
End Property

Public Property Get PowerSource() As EPowerSource
    Dim BatteryState As BatteryState: Set BatteryState = GetCurrentBatteryState
    If IsUpsPresent Then
        PowerSource = EPowerSource.Ups
    ElseIf Not BatteryState.IsBatteryPresent And BatteryState.IsACOnLine Then
        PowerSource = EPowerSource.AC
    Else
        PowerSource = EPowerSource.Battery
    End If
End Property

Private Function GetCurrentBatteryState() As BatteryState
    Set GetCurrentBatteryState = New BatteryState
End Function
    
Public Function PowerSource_ToStr(aps As EPowerSource) As String
    Dim s As String
    Select Case aps
    Case EPowerSource.Ups:      s = "PowerSource.Ups"
    Case EPowerSource.AC:       s = "PowerSource.AC"
    Case EPowerSource.Battery:  s = "PowerSource.Battery"
    End Select
    PowerSource_ToStr = s
End Function

Public Property Get IsBatteryShortTerm() As Boolean
    IsBatteryShortTerm = m_SystemPowerCaps.BatteriesAreShortTerm
End Property

Public Function ToStr() As String
    Dim s As String
    With Me
        s = "PowerCapabilities" & vbCrLf & "{" & vbCrLf
        s = s & "    IsPowerButtonPresent     : " & CBool(.IsPowerButtonPresent) & ";" & vbCrLf
        s = s & "    IsSleepButtonPresent     : " & CBool(.IsSleepButtonPresent) & ";" & vbCrLf
        s = s & "    IsLidPresent             : " & CBool(.IsLidPresent) & ";" & vbCrLf
        s = s & "    IsSystemS1               : " & CBool(.IsSystemS1) & ";" & vbCrLf
        s = s & "    IsSystemS2               : " & CBool(.IsSystemS2) & ";" & vbCrLf
        s = s & "    IsSystemS3               : " & CBool(.IsSystemS3) & ";" & vbCrLf
        s = s & "    IsSystemS4               : " & CBool(.IsSystemS4) & ";" & vbCrLf
        s = s & "    IsSystemS5               : " & CBool(.IsSystemS5) & ";" & vbCrLf
        s = s & "    IsHiberFilePresent       : " & CBool(.IsHiberFilePresent) & ";" & vbCrLf
        s = s & "    IsFullWake               : " & CBool(.IsFullWake) & ";" & vbCrLf
        s = s & "    IsVideoDimPresent        : " & CBool(.IsVideoDimPresent) & ";" & vbCrLf
        s = s & "    IsApmPresent             : " & CBool(.IsApmPresent) & ";" & vbCrLf
        s = s & "    IsUpsPresent             : " & CBool(.IsUpsPresent) & ";" & vbCrLf
        s = s & "    IsThermalControl         : " & CBool(.IsThermalControl) & ";" & vbCrLf
        s = s & "    IsProcessorThrottle      : " & CBool(.IsProcessorThrottle) & ";" & vbCrLf

        s = s & "    ProcessorMinimumThrottle : " & .ProcessorMinimumThrottle & ";" & vbCrLf
        s = s & "    ProcessorMaximumThrottle : " & .ProcessorMaximumThrottle & ";" & vbCrLf
        s = s & "    FastSystemS4             : " & CBool(.IsFastSystemS4) & ";" & vbCrLf
        'spare2(1 To 3)
        s = s & "    IsDiskSpinDown           : " & CBool(.IsDiskSpinDown) & ";" & vbCrLf
        'spare3(1 To 8)

        s = s & "    SystemBatteriesPresent   : " & CBool(.SystemBatteriesArePresent) & ";" & vbCrLf
        s = s & "    BatteriesAreShortTerm    : " & CBool(.BatteriesAreShortTerm) & ";" & vbCrLf
        'BatteryScale(1 To 3)

        s = s & "    AcOnlineWake             : " & ESystemPowerState_ToStr(.ACOnlineWake) & ";" & vbCrLf
        s = s & "    SoftLidWake              : " & ESystemPowerState_ToStr(.SoftLidWake) & ";" & vbCrLf
        s = s & "    RtcWake                  : " & ESystemPowerState_ToStr(.RtcWake) & ";" & vbCrLf
        s = s & "    MinimumDeviceWakeState   : " & ESystemPowerState_ToStr(.MinimumDeviceWakeState) & ";" & vbCrLf
        s = s & "    DefaultLowLatencyWake    : " & ESystemPowerState_ToStr(.DefaultLowLatencyWake) & ";" & vbCrLf & "}"
    End With
    ToStr = s
End Function

Private Function GetSystemPowerCapabilities() As SystemPowerCapabilities
    Dim hr As Long: hr = CallNtPowerInformation(PowerInformationLevel.SystemPowerCapabilities, CLng(0), CLng(0), GetSystemPowerCapabilities, LenB(GetSystemPowerCapabilities))
    If hr = &HC0000022 Then
        Err.Raise hr, , "UnauthorizedAccessException(LocalizedMessages.PowerInsufficientAccessCapabilities)"
    End If
End Function

Public Function ESystemPowerState_ToStr(e As ESystemPowerState) As String
    Dim s As String
    Select Case e
    Case ESystemPowerState.Unspecified: s = "Unspecified"
    Case ESystemPowerState.Working:     s = "Working"
    Case ESystemPowerState.Sleeping1:   s = "Sleeping1"
    Case ESystemPowerState.Sleeping2:   s = "Sleeping2"
    Case ESystemPowerState.Sleeping3:   s = "Sleeping3"
    Case ESystemPowerState.Hibernate:   s = "Hibernate"
    Case ESystemPowerState.Shutdown:    s = "Shutdown"
    Case ESystemPowerState.Maximum:     s = "Maximum"
    End Select
    ESystemPowerState_ToStr = s
End Function

