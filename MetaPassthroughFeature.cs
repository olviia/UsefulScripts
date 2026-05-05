// MetaPassthroughFeature — OpenXR feature that enables Meta Quest passthrough without the Meta SDK.
//
// How it works:
//   Hooks into xrGetInstanceProcAddr via HookGetInstanceProcAddr to intercept xrEndFrame calls.
//   On every frame, it injects an XrCompositionLayerPassthroughFB at layer index 0 so the
//   real-world camera feed renders behind all other composition layers (AR background effect).
//   The XR_FB_passthrough extension is negotiated automatically by OpenXR; no Meta OVR SDK needed.
//
// How to use:
//   1. Open Edit > Project Settings > XR Plug-in Management > OpenXR.
//   2. Under "OpenXR Feature Groups", enable "Meta Passthrough" (appears after this script compiles).
//   3. Make sure XR_FB_passthrough is listed in the enabled extension strings (done automatically).
//   4. Set your camera's Clear Flags to "Solid Color" with alpha = 0 so the passthrough shows through.
//   5. Build and deploy to a Meta Quest device connected via Link or standalone — passthrough starts
//      automatically when the XR session opens; no additional runtime calls are needed.
//
// Requirements: Unity OpenXR Plugin, Meta Quest device (Quest 2/3/Pro), Android or PC Link build.

#if UNITY_EDITOR
using UnityEditor;
using UnityEditor.XR.OpenXR.Features;
#endif
using System;
using System.Runtime.InteropServices;
using UnityEngine;
using UnityEngine.XR.OpenXR.Features;

#if UNITY_EDITOR
[OpenXRFeature(
    UiName = "Meta Passthrough",
    BuildTargetGroups = new[] { BuildTargetGroup.Standalone, BuildTargetGroup.Android },
    Company = "TUS",
    Desc = "Enables Meta Quest passthrough via XR_FB_passthrough (no Meta SDK required)",
    OpenxrExtensionStrings = "XR_FB_passthrough",
    Version = "0.0.1",
    FeatureId = "com.tus.xr.feature.passthrough")]
#endif
public class MetaPassthroughFeature : OpenXRFeature
{
    private const int XR_TYPE_PASSTHROUGH_CREATE_INFO_FB         = 1000118001;
    private const int XR_TYPE_PASSTHROUGH_LAYER_CREATE_INFO_FB   = 1000118002;
    private const int XR_TYPE_COMPOSITION_LAYER_PASSTHROUGH_FB   = 1000118003;
    private const ulong XR_PASSTHROUGH_IS_RUNNING_AT_CREATION_BIT_FB = 0x00000001;
    private const int XR_PASSTHROUGH_LAYER_PURPOSE_RECONSTRUCTION_FB = 0;

    [StructLayout(LayoutKind.Sequential)]
    private struct XrPassthroughCreateInfoFB
    {
        public int type; public IntPtr next; public ulong flags;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct XrPassthroughLayerCreateInfoFB
    {
        public int type; public IntPtr next;
        public ulong passthrough; public ulong flags; public int purpose;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct XrCompositionLayerPassthroughFB
    {
        public int type; public IntPtr next;
        public ulong flags; public ulong space; public ulong layerHandle;
    }

    // Must match XrFrameEndInfo memory layout on 64-bit:
    // int(4) + pad(4) + ptr(8) + long(8) + int(4) + uint(4) + ptr(8) = 40 bytes
    [StructLayout(LayoutKind.Sequential)]
    private struct XrFrameEndInfo
    {
        public int type; public IntPtr next;
        public long displayTime;
        public int environmentBlendMode;
        public uint layerCount;
        public IntPtr layers;
    }

    // Called FROM native (takes IntPtr for the name string to avoid marshaling issues)
    private delegate int DelegateHookGetProcAddr(ulong instance, IntPtr namePtr, out IntPtr function);
    // For calling native xrGetInstanceProcAddr
    private delegate int DelegateCallGetProcAddr(ulong instance, [MarshalAs(UnmanagedType.LPStr)] string name, out IntPtr function);

    private delegate int DelegateCreatePassthrough(ulong session, ref XrPassthroughCreateInfoFB info, out ulong passthrough);
    private delegate int DelegateStartPassthrough(ulong passthrough);
    private delegate int DelegateDestroyPassthrough(ulong passthrough);
    private delegate int DelegateCreateLayer(ulong session, ref XrPassthroughLayerCreateInfoFB info, out ulong layer);
    private delegate int DelegateResumeLayer(ulong layer);
    private delegate int DelegateDestroyLayer(ulong layer);
    private delegate int DelegateEndFrame(ulong session, IntPtr frameEndInfo);

    // Keep all delegates alive to prevent GC from collecting them while native code holds pointers
    private DelegateCallGetProcAddr m_OriginalProcAddr;
    private DelegateHookGetProcAddr m_HookedProcAddr;
    private DelegateEndFrame m_RealEndFrame;
    private DelegateEndFrame m_HookedEndFrame;

    private ulong m_XrInstance;
    private ulong m_PassthroughHandle;
    private ulong m_LayerHandle;

    protected override IntPtr HookGetInstanceProcAddr(IntPtr func)
    {
        m_OriginalProcAddr = Marshal.GetDelegateForFunctionPointer<DelegateCallGetProcAddr>(func);
        m_HookedProcAddr   = InterceptedGetInstanceProcAddr;
        m_HookedEndFrame   = HookedEndFrame;
        return Marshal.GetFunctionPointerForDelegate(m_HookedProcAddr);
    }

    // Unity calls this for every OpenXR function lookup — we intercept xrEndFrame
    private int InterceptedGetInstanceProcAddr(ulong instance, IntPtr namePtr, out IntPtr function)
    {
        string name = Marshal.PtrToStringAnsi(namePtr);
        int result = m_OriginalProcAddr(instance, name, out function);

        if (name == "xrEndFrame" && result == 0 && function != IntPtr.Zero)
        {
            m_RealEndFrame = Marshal.GetDelegateForFunctionPointer<DelegateEndFrame>(function);
            function = Marshal.GetFunctionPointerForDelegate(m_HookedEndFrame);
        }

        return result;
    }

    // Called instead of xrEndFrame — injects the passthrough composition layer
    private int HookedEndFrame(ulong session, IntPtr frameEndInfoPtr)
    {
        if (m_LayerHandle == 0 || m_RealEndFrame == null)
            return m_RealEndFrame?.Invoke(session, frameEndInfoPtr) ?? 0;

        var info = Marshal.PtrToStructure<XrFrameEndInfo>(frameEndInfoPtr);
        int existing = (int)info.layerCount;

        var ptLayer = new XrCompositionLayerPassthroughFB
        {
            type = XR_TYPE_COMPOSITION_LAYER_PASSTHROUGH_FB,
            next = IntPtr.Zero,
            flags = 0,
            space = 0,
            layerHandle = m_LayerHandle
        };

        IntPtr ptLayerPtr   = Marshal.AllocHGlobal(Marshal.SizeOf<XrCompositionLayerPassthroughFB>());
        IntPtr newLayersPtr = Marshal.AllocHGlobal((existing + 1) * IntPtr.Size);
        IntPtr newInfoPtr   = Marshal.AllocHGlobal(Marshal.SizeOf<XrFrameEndInfo>());

        try
        {
            Marshal.StructureToPtr(ptLayer, ptLayerPtr, false);

            // Passthrough at index 0 so it renders behind everything
            Marshal.WriteIntPtr(newLayersPtr, 0, ptLayerPtr);
            for (int i = 0; i < existing; i++)
                Marshal.WriteIntPtr(newLayersPtr, (i + 1) * IntPtr.Size,
                    Marshal.ReadIntPtr(info.layers, i * IntPtr.Size));

            var modified = info;
            modified.layerCount = (uint)(existing + 1);
            modified.layers = newLayersPtr;
            Marshal.StructureToPtr(modified, newInfoPtr, false);

            return m_RealEndFrame(session, newInfoPtr);
        }
        finally
        {
            Marshal.FreeHGlobal(ptLayerPtr);
            Marshal.FreeHGlobal(newLayersPtr);
            Marshal.FreeHGlobal(newInfoPtr);
        }
    }

    protected override bool OnInstanceCreate(ulong xrInstance)
    {
        m_XrInstance = xrInstance;
        return base.OnInstanceCreate(xrInstance);
    }

    protected override void OnSessionCreate(ulong xrSession)
    {
        TryEnablePassthrough(xrSession);
    }

    protected override void OnSessionDestroy(ulong xrSession)
    {
        TryCleanup();
    }

    private void TryEnablePassthrough(ulong xrSession)
    {
        try
        {
            var fnCreate      = GetDelegate<DelegateCreatePassthrough>("xrCreatePassthroughFB");
            var fnStart       = GetDelegate<DelegateStartPassthrough>("xrPassthroughStartFB");
            var fnCreateLayer = GetDelegate<DelegateCreateLayer>("xrCreatePassthroughLayerFB");
            var fnResumeLayer = GetDelegate<DelegateResumeLayer>("xrPassthroughLayerResumeFB");

            var createInfo = new XrPassthroughCreateInfoFB
                { type = XR_TYPE_PASSTHROUGH_CREATE_INFO_FB, next = IntPtr.Zero, flags = 0 };

            int r = fnCreate(xrSession, ref createInfo, out m_PassthroughHandle);
            if (r != 0) { Debug.LogError($"[Passthrough] xrCreatePassthroughFB failed: {r}"); return; }

            r = fnStart(m_PassthroughHandle);
            if (r != 0) { Debug.LogError($"[Passthrough] xrPassthroughStartFB failed: {r}"); return; }

            var layerInfo = new XrPassthroughLayerCreateInfoFB
            {
                type = XR_TYPE_PASSTHROUGH_LAYER_CREATE_INFO_FB,
                next = IntPtr.Zero,
                passthrough = m_PassthroughHandle,
                flags = XR_PASSTHROUGH_IS_RUNNING_AT_CREATION_BIT_FB,
                purpose = XR_PASSTHROUGH_LAYER_PURPOSE_RECONSTRUCTION_FB
            };

            r = fnCreateLayer(xrSession, ref layerInfo, out m_LayerHandle);
            if (r != 0) { Debug.LogError($"[Passthrough] xrCreatePassthroughLayerFB failed: {r}"); return; }

            r = fnResumeLayer(m_LayerHandle);
            if (r != 0) { Debug.LogError($"[Passthrough] xrPassthroughLayerResumeFB failed: {r}"); return; }

            Debug.Log("[Passthrough] Passthrough enabled.");
        }
        catch (Exception e)
        {
            Debug.LogError($"[Passthrough] Failed: {e.Message}");
        }
    }

    private void TryCleanup()
    {
        try
        {
            if (m_LayerHandle != 0)
            {
                GetDelegate<DelegateDestroyLayer>("xrDestroyPassthroughLayerFB")(m_LayerHandle);
                m_LayerHandle = 0;
            }
            if (m_PassthroughHandle != 0)
            {
                GetDelegate<DelegateDestroyPassthrough>("xrDestroyPassthroughFB")(m_PassthroughHandle);
                m_PassthroughHandle = 0;
            }
        }
        catch (Exception e)
        {
            Debug.LogError($"[Passthrough] Cleanup failed: {e.Message}");
        }
    }

    private T GetDelegate<T>(string funcName) where T : Delegate
    {
        m_OriginalProcAddr(m_XrInstance, funcName, out IntPtr ptr);
        if (ptr == IntPtr.Zero)
            throw new Exception($"Function pointer not found: {funcName}");
        return Marshal.GetDelegateForFunctionPointer<T>(ptr);
    }
}
