namespace Windows
{
    using System;
    using System.Collections.Generic;
    using System.Runtime.InteropServices;

    internal static class AudioManager
    {
        public static AudioApp FindApplication(int id)
        {
            foreach (var app in EnumerateApplications())
            {
                if (app.ID == id)
                {
                    return app;
                }
                app.Dispose();
            }

            return null;
        }

        public static IEnumerable<AudioApp> EnumerateApplications()
        {
            var list = new List<AudioApp>();
            var deviceEnumerator = (IMMDeviceEnumerator)new MMDeviceEnumerator();
            Marshal.ThrowExceptionForHR(deviceEnumerator.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia, out IMMDevice device));

            Guid IID_IAudioSessionManager2 = typeof(IAudioSessionManager2).GUID;
            Marshal.ThrowExceptionForHR(device.Activate(ref IID_IAudioSessionManager2, ClsCtx.Local, IntPtr.Zero, out object obj));
            var sessionManager = (IAudioSessionManager2)obj;

            Marshal.ThrowExceptionForHR(sessionManager.GetSessionEnumerator(out IAudioSessionEnumerator sessionEnumerator));
            Marshal.ThrowExceptionForHR(sessionEnumerator.GetCount(out int count));

            for (int i = 0; i < count; i++)
            {
                Marshal.ThrowExceptionForHR(sessionEnumerator.GetSession(i, out IAudioSessionControl sessionControl));
                var sessionControl2 = sessionControl as IAudioSessionControl2;
                if (sessionControl2 != null)
                {
                    Marshal.ThrowExceptionForHR(sessionControl2.GetDisplayName(out string displayName));
                    Marshal.ThrowExceptionForHR(sessionControl2.GetProcessId(out uint id));
                    list.Add(new AudioApp(displayName, Convert.ToInt32(id)));
                    Marshal.ReleaseComObject(sessionControl2);
                }
            }

            Marshal.ReleaseComObject(sessionEnumerator);
            Marshal.ReleaseComObject(sessionManager);
            Marshal.ReleaseComObject(device);
            Marshal.ReleaseComObject(deviceEnumerator);

            return list;
        }

        public static ISimpleAudioVolume GetVolumeInterface(uint processId)
        {
            var deviceEnumerator = (IMMDeviceEnumerator)new MMDeviceEnumerator();
            Marshal.ThrowExceptionForHR(deviceEnumerator.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia, out IMMDevice device));

            var IID_IAudioSessionManager2 = typeof(IAudioSessionManager2).GUID;
            Marshal.ThrowExceptionForHR(device.Activate(ref IID_IAudioSessionManager2, ClsCtx.Local, IntPtr.Zero, out object obj));
            var sessionManager = (IAudioSessionManager2)obj;

            Marshal.ThrowExceptionForHR(sessionManager.GetSessionEnumerator(out IAudioSessionEnumerator sessionEnumerator));
            Marshal.ThrowExceptionForHR(sessionEnumerator.GetCount(out int count));

            ISimpleAudioVolume simpleVolume = null;
            for (int i = 0; i < count; i++)
            {
                Marshal.ThrowExceptionForHR(sessionEnumerator.GetSession(i, out IAudioSessionControl sessionControl));
                var sessionControl2 = sessionControl as IAudioSessionControl2;
                if (sessionControl2 != null)
                {
                    Marshal.ThrowExceptionForHR(sessionControl2.GetProcessId(out uint id));
                    if (id == processId)
                    {
                        simpleVolume = sessionControl2 as ISimpleAudioVolume;
                        break;
                    }

                    Marshal.ReleaseComObject(sessionControl2);
                }
            }

            Marshal.ReleaseComObject(sessionEnumerator);
            Marshal.ReleaseComObject(sessionManager);
            Marshal.ReleaseComObject(device);
            Marshal.ReleaseComObject(deviceEnumerator);

            return simpleVolume;
        }
    }

    [ComImport, Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")]
    internal class MMDeviceEnumerator
    {
    }

    [Flags]
    internal enum DeviceState
    {
        Active = 0x00000001,
        Disabled = 0x00000002,
        NotPresent = 0x00000004,
        Unplugged = 0x00000008,
        All = 0x0000000F
    }

    internal enum DataFlow
    {
        Render,
        Capture,
        All
    }

    internal enum Role
    {
        Console,
        Multimedia,
        Communications,
    }

    [Guid("A95664D2-9614-4F35-A746-DE8DB63617E6"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IMMDeviceEnumerator
    {
        int EnumAudioEndpoints(DataFlow dataFlow, DeviceState stateMask, out IMMDeviceCollection devices);

        [PreserveSig]
        int GetDefaultAudioEndpoint(DataFlow dataFlow, Role role, out IMMDevice endpoint);

        int GetDevice(string id, out IMMDevice deviceName);

        int RegisterEndpointNotificationCallback(object client);

        int UnregisterEndpointNotificationCallback(object client);
    }

    [Guid("0BD7A1BE-7A1A-44DB-8397-CC5392387B5E"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IMMDeviceCollection
    {
        int GetCount(out int numDevices);
        int Item(int deviceNumber, out IMMDevice device);
    }

    [Guid("D666063F-1587-4E43-81F1-B948E807363F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IMMDevice
    {
        int Activate(ref Guid id, ClsCtx clsCtx, IntPtr activationParams, [MarshalAs(UnmanagedType.IUnknown)] out object interfacePointer);

        int OpenPropertyStore(object stgmAccess, out object properties);

        int GetId([MarshalAs(UnmanagedType.LPWStr)] out string id);

        int GetState(out DeviceState state);
    }

    [Flags]
    internal enum ClsCtx
    {
        INPROC_SERVER = 0x1,
        INPROC_HANDLER = 0x2,
        LOCAL_SERVER = 0x4,
        INPROC_SERVER16 = 0x8,
        REMOTE_SERVER = 0x10,
        INPROC_HANDLER16 = 0x20,
        RESERVED1 = 0x40,
        RESERVED2 = 0x80,
        RESERVED3 = 0x100,
        RESERVED4 = 0x200,
        NO_CODE_DOWNLOAD = 0x400,
        RESERVED5 = 0x800,
        NO_CUSTOM_MARSHAL = 0x1000,
        ENABLE_CODE_DOWNLOAD = 0x2000,
        NO_FAILURE_LOG = 0x4000,
        DISABLE_AAA = 0x8000,
        ENABLE_AAA = 0x10000,
        FROM_DEFAULT_CONTEXT = 0x20000,
        ACTIVATE_X86_SERVER = 0x40000,
        ACTIVATE_32_BIT_SERVER,
        ACTIVATE_64_BIT_SERVER = 0x80000,
        ENABLE_CLOAKING = 0x100000,
        APPCONTAINER = 0x400000,
        ACTIVATE_AAA_AS_IU = 0x800000,
        RESERVED6 = 0x1000000,
        ACTIVATE_ARM32_SERVER = 0x2000000,
        ALLOW_LOWER_TRUST_REGISTRATION,
        PS_DLL = unchecked((int)0x80000000),
        Local = INPROC_SERVER | INPROC_HANDLER | LOCAL_SERVER
    }

    internal enum AudioSessionState
    {
        AudioSessionStateInactive = 0,
        AudioSessionStateActive = 1,
        AudioSessionStateExpired = 2
    }

    internal enum AudioSessionDisconnectReason
    {
        DisconnectReasonDeviceRemoval = 0,
        DisconnectReasonServerShutdown = 1,
        DisconnectReasonFormatChanged = 2,
        DisconnectReasonSessionLogoff = 3,
        DisconnectReasonSessionDisconnected = 4,
        DisconnectReasonExclusiveModeOverride = 5
    }

    [Guid("BFA971F1-4D5E-40BB-935E-967039BFBEE4"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IAudioSessionManager
    {
        [PreserveSig]
        int GetAudioSessionControl([In, Optional][MarshalAs(UnmanagedType.LPStruct)] Guid sessionId, [In][MarshalAs(UnmanagedType.U4)] UInt32 streamFlags, [Out][MarshalAs(UnmanagedType.Interface)] out IAudioSessionControl sessionControl);

        [PreserveSig]
        int GetSimpleAudioVolume([In, Optional][MarshalAs(UnmanagedType.LPStruct)] Guid sessionId, [In][MarshalAs(UnmanagedType.U4)] UInt32 streamFlags, [Out][MarshalAs(UnmanagedType.Interface)] out ISimpleAudioVolume audioVolume);
    }


    [Guid("77AA99A0-1BD6-484F-8BC7-2C654C9A9B6F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IAudioSessionManager2 : IAudioSessionManager
    {
        [PreserveSig]
        new int GetAudioSessionControl([In, Optional][MarshalAs(UnmanagedType.LPStruct)] Guid sessionId, [In][MarshalAs(UnmanagedType.U4)] UInt32 streamFlags, [Out][MarshalAs(UnmanagedType.Interface)] out IAudioSessionControl sessionControl);

        [PreserveSig]
        new int GetSimpleAudioVolume([In, Optional][MarshalAs(UnmanagedType.LPStruct)] Guid sessionId, [In][MarshalAs(UnmanagedType.U4)] UInt32 streamFlags, [Out][MarshalAs(UnmanagedType.Interface)] out ISimpleAudioVolume audioVolume);

        [PreserveSig]
        int GetSessionEnumerator(out IAudioSessionEnumerator sessionEnum);

        [PreserveSig]
        int RegisterSessionNotification(IAudioSessionNotification sessionNotification);

        [PreserveSig]
        int UnregisterSessionNotification(IAudioSessionNotification sessionNotification);

        [PreserveSig]
        int RegisterDuckNotification(string sessionId, IAudioSessionNotification audioVolumeDuckNotification);

        [PreserveSig]
        int UnregisterDuckNotification(IntPtr audioVolumeDuckNotification);
    }

    [Guid("E2F5BB11-0570-40CA-ACDD-3AA01277DEE8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IAudioSessionEnumerator
    {
        int GetCount(out int sessionCount);

        int GetSession(int sessionCount, out IAudioSessionControl session);
    }

    [Guid("641DD20B-4D41-49CC-ABA3-174B9477BB08"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IAudioSessionNotification
    {
        [PreserveSig]
        int OnSessionCreated(IAudioSessionControl newSession);
    }

    [Guid("24918ACC-64B3-37C1-8CA9-74A66E9957A8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IAudioSessionEvents
    {
        [PreserveSig]
        int OnDisplayNameChanged([In][MarshalAs(UnmanagedType.LPWStr)] string displayName, [In] ref Guid eventContext);

        [PreserveSig]
        int OnIconPathChanged([In][MarshalAs(UnmanagedType.LPWStr)] string iconPath, [In] ref Guid eventContext);

        [PreserveSig]
        int OnSimpleVolumeChanged([In][MarshalAs(UnmanagedType.R4)] float volume, [In][MarshalAs(UnmanagedType.Bool)] bool isMuted, [In] ref Guid eventContext);

        [PreserveSig]
        int OnChannelVolumeChanged([In][MarshalAs(UnmanagedType.U4)] UInt32 channelCount, [In][MarshalAs(UnmanagedType.SysInt)] IntPtr newVolumes, [In][MarshalAs(UnmanagedType.U4)] UInt32 channelIndex, [In] ref Guid eventContext);

        [PreserveSig]
        int OnGroupingParamChanged([In] ref Guid groupingId, [In] ref Guid eventContext);

        [PreserveSig]
        int OnStateChanged([In] AudioSessionState state);

        [PreserveSig]
        int OnSessionDisconnected([In] AudioSessionDisconnectReason disconnectReason);
    }

    [Guid("F4B1A599-7266-4319-A8CA-E70ACB11E8CD"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IAudioSessionControl
    {
        [PreserveSig]
        int GetState([Out] out AudioSessionState state);

        [PreserveSig]
        int GetDisplayName([Out][MarshalAs(UnmanagedType.LPWStr)] out string displayName);

        [PreserveSig]
        int SetDisplayName([In][MarshalAs(UnmanagedType.LPWStr)] string displayName, [In][MarshalAs(UnmanagedType.LPStruct)] Guid eventContext);

        [PreserveSig]
        int GetIconPath([Out][MarshalAs(UnmanagedType.LPWStr)] out string iconPath);

        [PreserveSig]
        int SetIconPath([In][MarshalAs(UnmanagedType.LPWStr)] string iconPath, [In][MarshalAs(UnmanagedType.LPStruct)] Guid eventContext);

        [PreserveSig]
        int GetGroupingParam([Out] out Guid groupingId);

        [PreserveSig]
        int SetGroupingParam([In][MarshalAs(UnmanagedType.LPStruct)] Guid groupingId, [In][MarshalAs(UnmanagedType.LPStruct)] Guid eventContext);

        [PreserveSig]
        int RegisterAudioSessionNotification([In] IAudioSessionEvents client);

        [PreserveSig]
        int UnregisterAudioSessionNotification([In] IAudioSessionEvents client);
    }

    [Guid("bfb7ff88-7239-4fc9-8fa2-07c950be9c6d"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IAudioSessionControl2 : IAudioSessionControl
    {
        [PreserveSig]
        new int GetState([Out] out AudioSessionState state);

        [PreserveSig]
        new int GetDisplayName([Out][MarshalAs(UnmanagedType.LPWStr)] out string displayName);

        [PreserveSig]
        new int SetDisplayName([In][MarshalAs(UnmanagedType.LPWStr)] string displayName, [In][MarshalAs(UnmanagedType.LPStruct)] Guid eventContext);

        [PreserveSig]
        new int GetIconPath([Out][MarshalAs(UnmanagedType.LPWStr)] out string iconPath);

        [PreserveSig]
        new int SetIconPath([In][MarshalAs(UnmanagedType.LPWStr)] string iconPath, [In][MarshalAs(UnmanagedType.LPStruct)] Guid eventContext);

        [PreserveSig]
        new int GetGroupingParam([Out] out Guid groupingId);

        [PreserveSig]
        new int SetGroupingParam([In][MarshalAs(UnmanagedType.LPStruct)] Guid groupingId, [In][MarshalAs(UnmanagedType.LPStruct)] Guid eventContext);

        [PreserveSig]
        new int RegisterAudioSessionNotification([In] IAudioSessionEvents client);

        [PreserveSig]
        new int UnregisterAudioSessionNotification([In] IAudioSessionEvents client);

        [PreserveSig]
        int GetSessionIdentifier([Out][MarshalAs(UnmanagedType.LPWStr)] out string retVal);

        [PreserveSig]
        int GetSessionInstanceIdentifier([Out][MarshalAs(UnmanagedType.LPWStr)] out string retVal);

        [PreserveSig]
        int GetProcessId([Out] out UInt32 retVal);

        [PreserveSig]
        int IsSystemSoundsSession();

        [PreserveSig]
        int SetDuckingPreference(bool optOut);
    }

    [Guid("87CE5498-68D6-44E5-9215-6DA47EF883D8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface ISimpleAudioVolume
    {
        [PreserveSig]
        int SetMasterVolume([In][MarshalAs(UnmanagedType.R4)] float levelNorm, [In][MarshalAs(UnmanagedType.LPStruct)] Guid eventContext);

        [PreserveSig]
        int GetMasterVolume([Out][MarshalAs(UnmanagedType.R4)] out float levelNorm);

        [PreserveSig]
        int SetMute([In][MarshalAs(UnmanagedType.Bool)] bool isMuted, [In][MarshalAs(UnmanagedType.LPStruct)] Guid eventContext);

        [PreserveSig]
        int GetMute([Out][MarshalAs(UnmanagedType.Bool)] out bool isMuted);
    }
}
