using System;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace Multitool
{
    /// <summary>
    /// Класс для проверки лицензии приложения
    /// </summary>
    public static class LicenseValidator
    {
        private const string LicenseFileName = "bug";

        /// <summary>
        /// Проверяет наличие и валидность лицензии
        /// </summary>
        /// <returns>true если лицензия валидна, false в противном случае</returns>
        public static bool IsValid()
        {
            return IsValid(out _);
        }

        /// <summary>
        /// Проверяет наличие и валидность лицензии с выводом причины ошибки
        /// </summary>
        /// <param name="errorMessage">Сообщение об ошибке, если лицензия невалидна</param>
        /// <returns>true если лицензия валидна, false в противном случае</returns>
        public static bool IsValid(out string errorMessage)
        {
            errorMessage = null;
            string licenseFilePath = GetLicenseFilePath();

            if (!File.Exists(licenseFilePath))
            {
                errorMessage = "Отсутствует файл лицензии";
                return false;
            }

            try
            {
                string licenseContent = File.ReadAllText(licenseFilePath);
                string decodedHex = DecodeFromHex(licenseContent);
                string decodedLicense = DecodeFromBase64(decodedHex);
                string machineFingerprint = GetMachineFingerprint();

                if (decodedLicense != machineFingerprint)
                {
                    errorMessage = "Лицензия недействительна для данного компьютера";
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                errorMessage = "Ошибка проверки лицензии: " + ex.Message;
                return false;
            }
        }

        /// <summary>
        /// Генерирует уникальный отпечаток машины
        /// </summary>
        public static string GetMachineFingerprint()
        {
            return Environment.UserName + Environment.OSVersion + Environment.MachineName;
        }

        /// <summary>
        /// Возвращает путь к файлу лицензии
        /// </summary>
        public static string GetLicenseFilePath()
        {
            return Path.Combine(Application.StartupPath, LicenseFileName);
        }

        private static string DecodeFromBase64(string base64EncodedData)
        {
            byte[] base64EncodedBytes = Convert.FromBase64String(base64EncodedData);
            return Encoding.UTF8.GetString(base64EncodedBytes);
        }

        private static string DecodeFromHex(string hex)
        {
            if (hex.Length % 2 != 0)
                throw new ArgumentException("Hex string must have an even length.");

            byte[] bytes = new byte[hex.Length / 2];

            for (int i = 0; i < hex.Length; i += 2)
            {
                bytes[i / 2] = Convert.ToByte(hex.Substring(i, 2), 16);
            }

            return Encoding.UTF8.GetString(bytes);
        }
    }
}
