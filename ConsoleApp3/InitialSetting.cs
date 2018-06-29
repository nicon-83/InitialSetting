using System;
using System.Diagnostics;
using System.Management;
using System.Threading;
using System.IO;

namespace MVA.ConsoleApp
{
    class InitialSetting
    {

        private static readonly String userName = "user";
        private static readonly String windir = Environment.GetEnvironmentVariable("windir");
        private static readonly String robocopyPath = windir + @"\system32\robocopy.exe";
        private static readonly String[] robocopyArgs =
        {
            @"\\usl033\Distr$\Microsoft\SCCM\Client C:\Temp\\Client /E /Z",
            @"\\usl033\Distr$\TM C:\Temp *64*"
        };
        private static readonly String sccmPath = @"c:\temp\client\ccmsetup.exe";
        private static readonly String trendMicroPath = @"c:\temp\USL050_Client_64bit.exe";
        private static Process SCCM = new Process();
        private static Process TrendMicro = new Process();
        private static Boolean SccmEventHandled = false;
        private static Boolean TrendMicroEventHandled = false;
        private const int SLEEP_AMOUNT = 30000;

        private class Type
        {
            public static String success = "success";
            public static String information = "information";
            public static String error = "error";
        }

        private static void Main(string[] args)
        {
            PrintMessage("Будет выполнена настройка компьютера", Type.information);
            Console.WriteLine("Для продолжения нажмите Enter или любую другую клавишу для отмены...");
            ConsoleKeyInfo key = Console.ReadKey();

            if (key.Key.ToString() != "Enter")
            {
                return;
            }

            DeleteUserProfile(userName);
            DeleteLocalUser(userName);
            CopyDistributives();
            InstallApp();
            DeleteDistributives();

            Console.WriteLine("Для продолжения нажмите любую клавишу...");
            Console.ReadKey();
        }

        public static void DeleteUserProfile(string userName)
        {
            #region Удаляем профиль пользователя

            PrintMessage(String.Format("{0:T} Удаляем профиль пользователя с именем {1}", DateTime.Now, userName), Type.information);
            ManagementObject user = null;
            SelectQuery query = new SelectQuery("Select * from win32_userprofile");
            ManagementObjectSearcher searcher = null;
            try
            {
                searcher = new ManagementObjectSearcher(query);
                foreach (ManagementObject userProfile in searcher.Get())
                {
                    if (userProfile["localpath"].ToString().ToLower() == "c:\\users\\" + userName.ToLower())
                    {
                        user = userProfile;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            if (user != null)
            {
                try
                {
                    user.Delete();
                    PrintMessage(String.Format("Профиль пользователя {0} успешно удален", userName), Type.success);
                }
                catch (Exception ex)
                {
                    PrintMessage(String.Format("Произошла ошибка при удалении профиля пользователя {0}" + ex.Message, userName), Type.error);
                }
            }
            else
            {
                PrintMessage(String.Format("Не найден профиль пользователя с именем {0}", userName), Type.error);
            }

            Console.WriteLine("Список профилей пользователей:");
            foreach (ManagementObject userProfile in searcher.Get())
            {
                Console.WriteLine("\t" + userProfile["localpath"]);
            }
            Console.WriteLine();

            #endregion
        }

        public static void DeleteLocalUser(string userName)
        {
            #region Удаляем локальную учетную запись пользователя

            PrintMessage(String.Format("{0:T} Удаляем локальную учетную запись пользователя с именем {1}", DateTime.Now, userName), Type.information);
            ManagementObject localUser = null;
            SelectQuery query = new SelectQuery("Select * from Win32_UserAccount Where LocalAccount = True");
            ManagementObjectSearcher searcher = null;
            try
            {
                searcher = new ManagementObjectSearcher(query);
                foreach (ManagementObject lusr in searcher.Get())
                {
                    if (lusr["Name"].ToString().ToLower() == userName.ToLower())
                    {
                        localUser = lusr;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            if (localUser != null)
            {
                try
                {
                    Process cmd = new Process();
                    cmd.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                    cmd.StartInfo.FileName = windir + @"\system32\cmd.exe";
                    cmd.StartInfo.Arguments = @"/c net user " + userName + " /delete";
                    cmd.Start();
                    cmd.WaitForExit();

                    foreach (ManagementObject lusr in searcher.Get())
                    {
                        if (lusr.GetPropertyValue("name").ToString().ToLower() == userName.ToLower())
                        {
                            localUser = lusr;
                            break;
                        }
                        else
                        {
                            localUser = null;
                        }
                    }

                    if (localUser == null)
                    {
                        PrintMessage(String.Format("Локальный пользователь {0} успешно удален", userName), Type.success);
                    }
                    else
                    {
                        PrintMessage(String.Format("Произошла ошибка при удалении учетной записи пользователя {0}: ", userName), Type.error);
                    }
                }
                catch (Exception ex)
                {
                    PrintMessage(String.Format("Произошла ошибка при удалении учетной записи пользователя {0}: " + ex.Message, userName), Type.error);
                }
            }
            else
            {
                PrintMessage(String.Format("Не найден пользователь с именем {0}", userName), Type.error);
            }

            Console.WriteLine("Список локальных пользователей:");
            foreach (ManagementObject lusr in searcher.Get())
            {
                Console.WriteLine("\t" + lusr["Name"]);
            }
            Console.WriteLine();

            #endregion
        }

        private static void CopyDistributives()
        {
            #region Копируем дистрибутивы SCCM и TrendMicro в папку c:\temp

            PrintMessage(String.Format("{0:T} Копируем дистрибутивы SCCM и TrendMicro в папку c:\\temp", DateTime.Now), Type.information);

            Process SCCM = new Process();
            try
            {
                ProcessStartInfo startInfo = new ProcessStartInfo
                {
                    FileName = robocopyPath,
                    Arguments = robocopyArgs[0],
                    CreateNoWindow = true
                };
                SCCM.StartInfo = startInfo;
                SCCM.Start();
                SCCM.WaitForExit();
            }
            catch (Exception ex)
            {
                PrintMessage(String.Format("Произошла ошибка при копировании клиента SCCM:" + ex.Message), Type.error);
            }

            Process TrendMicro = new Process();
            try
            {
                ProcessStartInfo startInfo = new ProcessStartInfo
                {
                    FileName = robocopyPath,
                    Arguments = robocopyArgs[1],
                    CreateNoWindow = true
                };
                TrendMicro.StartInfo = startInfo;
                TrendMicro.Start();
                TrendMicro.WaitForExit();
            }
            catch (Exception ex)
            {
                PrintMessage(String.Format("Произошла ошибка при копировании клиента TrendMicro:" + ex.Message), Type.error);
            }

            if (SCCM.ExitCode == 1 && TrendMicro.ExitCode == 1)
            {
                PrintMessage(String.Format("Копирование успешно завершено"), Type.success);
            }
            else
            {
                if (SCCM.ExitCode != 1)
                {
                    PrintMessage(String.Format("Копирование клиента SCCM завершено с ошибками! ExitCode = " + SCCM.ExitCode), Type.error);
                }
                if (TrendMicro.ExitCode != 1)
                {
                    PrintMessage(String.Format("Копирование клиента TrendMicro завершено с ошибками! ExitCode = " + TrendMicro.ExitCode), Type.error);
                }
            }

            #endregion
        }

        private static void InstallApp()
        {
            #region Запускаем установку клиента CSSM и TrendMicro

            SCCM.StartInfo.FileName = sccmPath;
            SCCM.StartInfo.CreateNoWindow = true;
            SCCM.EnableRaisingEvents = true;
            SCCM.Exited += SCCM_Exited;
            try
            {
                SCCM.Start();
            }
            catch (Exception ex)
            {
                PrintMessage(String.Format("Ошибка при запуске процесса ccmsetup.exe: " + ex.Message), Type.error);
            }
            PrintMessage(String.Format("{0:T} Запущен процесс установки клиента CSSM", DateTime.Now), Type.information);

            TrendMicro.StartInfo.FileName = trendMicroPath;
            TrendMicro.StartInfo.CreateNoWindow = true;
            TrendMicro.EnableRaisingEvents = true;
            TrendMicro.Exited += TrendMicro_Exited;
            try
            {
                TrendMicro.Start();
            }
            catch (Exception ex)
            {
                PrintMessage(String.Format("Ошибка при запуске процесса установки TrendMicro: " + ex.Message), Type.error);
            }

            PrintMessage(String.Format("{0:T} Запущен процесс установки клиента TrendMicro", DateTime.Now), Type.information);

            while (!SccmEventHandled || !TrendMicroEventHandled)
            {
                try
                {
                    Thread.Sleep(SLEEP_AMOUNT);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }

            #endregion
        }

        private static void DeleteDistributives()
        {
            #region Выполняем очистку директории c:\temp

            PrintMessage(String.Format("{0:T} Выполняем очистку директории c:\\temp", DateTime.Now), Type.information);

            String dirPath = @"c:\Temp\";

            try
            {
                if (Directory.Exists(dirPath))
                {
                    DirectoryInfo directory = new DirectoryInfo(dirPath);
                    DirectoryInfo[] innerDirectories = directory.GetDirectories();
                    foreach (DirectoryInfo dir in innerDirectories)
                    {
                        dir.Delete(true);
                    }
                    FileInfo[] innerFiles = directory.GetFiles();
                    foreach (FileInfo file in innerFiles)
                    {
                        file.Delete();
                    }
                    PrintMessage(String.Format("Выполнена очистка директории {0}", dirPath), Type.success);
                }
                else
                {
                    PrintMessage(String.Format("Не найдена директория {0}", dirPath), Type.error);
                }
            }
            catch (Exception ex)
            {
                PrintMessage(String.Format("Произошла ошибка при очистке директории {0} :" + ex.Message, dirPath), Type.error);
            }

            #endregion
        }

        private static void PrintMessage(string messageText, string messageType)
        {
            #region Метод для печати сообщений
            switch (messageType)
            {
                case "success":
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine(messageText);
                    Console.ResetColor();
                    Console.WriteLine();
                    break;
                case "information":
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine(new string('-', 70));
                    Console.WriteLine(messageText);
                    Console.WriteLine(new string('-', 70));
                    Console.ResetColor();
                    Console.WriteLine();
                    break;
                case "error":
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine(messageText);
                    Console.ResetColor();
                    Console.WriteLine();
                    break;
                default:
                    Console.WriteLine(messageText);
                    Console.WriteLine();
                    break;
            }
            #endregion
        }

        private static void SCCM_Exited(object sender, EventArgs e)
        {
            #region Обработчик события - завершение процесса установки SCCM

            if (SCCM.ExitCode == 0)
            {
                PrintMessage(String.Format("{0:T} Приложение SCCM клиент успешно установлено", SCCM.ExitTime), Type.success);
            }
            else
            {
                PrintMessage(String.Format("При установке приложения SCCM клиент произошла ошибка с кодом завершения: {0}", SCCM.ExitCode), Type.error);
            }

            SccmEventHandled = true;

            #endregion
        }

        private static void TrendMicro_Exited(object sender, EventArgs e)
        {
            #region Обработчик события - завершение процесса установки TrendMicro

            if (TrendMicro.ExitCode == 0)
            {
                PrintMessage(String.Format("{0:T} Приложение TrendMicro успешно установлено", TrendMicro.ExitTime), Type.success);
            }
            else
            {
                PrintMessage(String.Format("При установке приложения TrendMicro произошла ошибка с кодом завершения: {0}", TrendMicro.ExitCode), Type.error);
            }

            TrendMicroEventHandled = true;

            #endregion
        }

    }
}