using System;
using System.Windows.Forms;

namespace OpenChoose
{
    /// <summary>“设为默认打开方式”的共用流程（界面 / 托盘 / 命令行共用）。</summary>
    internal static class DefaultProgramHelper
    {
        /// <summary>
        /// 注册文件关联并尝试设为默认；showUi 为 false 时静默执行（只写注册表）。
        /// </summary>
        public static SetDefaultResult Apply(IWin32Window owner, bool showUi)
        {
            SetDefaultResult result;
            try
            {
                result = FileAssociation.SetAsDefault();
            }
            catch (Exception ex)
            {
                Log.Write("设置默认打开方式失败", ex);
                if (showUi)
                {
                    MessageBox.Show(owner, "设置默认打开方式失败：\r\n" + ex.Message,
                        Program.AppName, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                return null;
            }

            if (!showUi || result == null)
            {
                return result;
            }

            if (result.NeedsSettings)
            {
                string extensions = string.Join("、", result.ProtectedExtensions.ToArray());
                string message =
                    "OpenChoose 已加入“打开方式”列表。\r\n\r\n" +
                    "Windows 10 / 11 会保护默认应用的设置（" + extensions + " 由系统的 UserChoice 决定），" +
                    "程序无法直接修改。\r\n\r\n" +
                    "接下来请在“默认应用”窗口中，把 .pptx / .ppt 改为\r\n" +
                    "“" + Program.FriendlyAppName + "”。\r\n\r\n是否现在打开该设置窗口？";

                if (MessageBox.Show(owner, message, Program.AppName,
                        MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                {
                    Launcher.OpenDefaultAppsSettings();
                }
            }
            else
            {
                MessageBox.Show(owner,
                    "已把 OpenChoose 设为 .ppt / .pptx 等文件的默认打开方式。\r\n双击 PPT 文件即可弹出选择窗口。",
                    Program.AppName, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            return result;
        }

        /// <summary>静默设置默认打开方式（供安装程序调用）。</summary>
        public static SetDefaultResult ApplySilently()
        {
            return Apply(null, false);
        }
    }
}
