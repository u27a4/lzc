using NAudio.CoreAudioApi;
using System;
using System.ComponentModel.Composition;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace lzc.ext.audioswitch
{
    [PartCreationPolicy(CreationPolicy.NonShared)]
    class AudioSwitch
    {
        [MenuItem]
        // 再生デバイスを変更するメニューを作成
        static MenuItem CreateMenuItem()
        {
            // 再生デバイスの列挙
            var enumerator = new MMDeviceEnumerator();
            var current = enumerator.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
            var collection = enumerator.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active);

            // 親メニューの生成
            var menu = new MenuItem("再生デバイスを変更", OnMenuItemClicked)
            {
                Tag = collection.Concat(collection).SkipWhile(device => current.ID != device.ID).First()?.ID,
            };

            // 子メニューの生成
            menu.MenuItems.AddRange(collection.Select(device => new MenuItem(device.FriendlyName, OnMenuItemClicked)
            {
                Tag = device.ID,
                Checked = (current.ID == device.ID),
            }).ToArray());

            return menu;
        }

        // メニューがクリックされた
        static void OnMenuItemClicked(object sender, EventArgs e)
        {
            if (sender is MenuItem menu)
            {
                if (menu.Tag is string id)
                {
                    SetDefaultEndpoint(id);
                }
            }
        }

        // 再生デバイスを変更
        static void SetDefaultEndpoint(string endpointID, uint role = 1U)
        {
            if (string.IsNullOrWhiteSpace(endpointID)) return;

            switch (new PolicyConfig())
            {
                case IPolicyConfig config:
                    Marshal.ThrowExceptionForHR(config.SetDefaultEndpoint(endpointID, role));
                    break;

                default:
                    throw new NotImplementedException("Corresponded interface not found while trying to call IPolicyConfig::SetDefaultEndpoint");
            }
        }
    }

    // 再生デバイス変更に必要なCOM関係
    [ComImport]
    [Guid("870AF99C-171D-4F9E-AF0D-E63DF40C2BC9")]
    internal class PolicyConfig { }

    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("F8679F50-850A-41CF-9C72-430F290290C8")]
    internal interface IPolicyConfig
    {
        int Unused1();
        int Unused2();
        int Unused3();
        int Unused4();
        int Unused5();
        int Unused6();
        int Unused7();
        int Unused8();
        int Unused9();
        int Unused10();

        [PreserveSig]
        int SetDefaultEndpoint(string pszDeviceName, uint role);
    }
}