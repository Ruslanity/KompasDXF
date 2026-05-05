using Microsoft.Win32;
using System;
using System.Management;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Multitool
{
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class Multitool
    {
        [return: MarshalAs(UnmanagedType.BStr)] public string GetLibraryName()
        {
            return "Multitool";
        }
        // Головная функция библиотеки
        public void ExternalRunCommand([In] short command, [In] short mode, [In, MarshalAs(UnmanagedType.IDispatch)] object kompas_)
        {
            if (LicenseValidator.IsValid(out string errorMessage))
            {
                MainForm.GetInstance().Show();
            }
            else
            {
                MessageBox.Show(errorMessage);
            }
        }

        public string GenerateMS()
        {
            string motherboardSerial = string.Empty;
            try
            {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT SerialNumber FROM Win32_BaseBoard");
                foreach (ManagementObject obj in searcher.Get())
                {
                    motherboardSerial = obj["SerialNumber"].ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return motherboardSerial;
        }

        #region COM Registration
        // Эта функция выполняется при регистрации класса для COM
        // Она добавляет в ветку реестра компонента раздел Kompas_Library,
        // который сигнализирует о том, что класс является приложением Компас,
        // а также заменяет имя InprocServer32 на полное, с указанием пути.
        // Все это делается для того, чтобы иметь возможность подключить
        // библиотеку на вкладке ActiveX.
        private const string KompasAddInPath = @"SOFTWARE\ASCON\KOMPAS-3D\AddIns\Multitool";

        [ComRegisterFunction]
        public static void RegisterKompasLib(Type t)
        {
            try
            {
                RegistryKey regKey = Registry.LocalMachine;
                string keyName = @"SOFTWARE\Classes\CLSID\{" + t.GUID.ToString() + "}";
                regKey = regKey.OpenSubKey(keyName, true);
                regKey.CreateSubKey("Kompas_Library");
                regKey = regKey.OpenSubKey("InprocServer32", true);
                regKey.SetValue(null, System.Environment.GetFolderPath(Environment.SpecialFolder.System) + @"\mscoree.dll");
                regKey.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("При регистрации класса для COM-Interop произошла ошибка:\n{0}", ex));
            }

            try
            {
                string assemblyPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
                RegistryKey addInKey;
                try
                {
                    addInKey = Registry.LocalMachine.CreateSubKey(KompasAddInPath);
                }
                catch
                {
                    addInKey = Registry.CurrentUser.CreateSubKey(KompasAddInPath);
                }
                using (addInKey)
                {
                    addInKey.SetValue("AutoConnect", 1, RegistryValueKind.DWord);
                    addInKey.SetValue("Path", assemblyPath);
                    addInKey.SetValue("ProgID", "Multitool.Multitool");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Ошибка регистрации аддина в КОМПАС:\n{0}", ex));
            }
        }

        // Эта функция удаляет раздел Kompas_Library из реестра
        [ComUnregisterFunction]
        public static void UnregisterKompasLib(Type t)
        {
            RegistryKey regKey = Registry.LocalMachine;
            string keyName = @"SOFTWARE\Classes\CLSID\{" + t.GUID.ToString() + "}";
            RegistryKey subKey = regKey.OpenSubKey(keyName, true);
            if (subKey != null)
            {
                subKey.DeleteSubKeyTree("Kompas_Library", false);
                subKey.Close();
            }

            try { Registry.LocalMachine.DeleteSubKey(KompasAddInPath, false); } catch { }
            try { Registry.CurrentUser.DeleteSubKey(KompasAddInPath, false); } catch { }
        }
        #endregion
    }
}
