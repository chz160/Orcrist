using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.Win32;
using Renci.SshNet;

namespace Orcrist
{
    class Program
    {
        [DllImport("wininet.dll")]
        public static extern bool InternetSetOption(IntPtr hInternet, int dwOption, IntPtr lpBuffer, int dwBufferLength);
        public const int INTERNET_OPTION_SETTINGS_CHANGED = 39;
        public const int INTERNET_OPTION_REFRESH = 37;

        private static string localhost = "127.0.0.1";

        static void Main(string[] args)
        {
            if (args.Any() && args[0] == @"\?")
            {
                Usage();
            }
            if (args.Length == 0 || args.Length != 5)
            {
                Console.WriteLine("Invalid arguments...");
                Usage();
            }

            var host = args[0];
            var destinationPort = int.Parse(args[1]);
            var localPort = uint.Parse(args[2]);
            var username = args[3];
            var keyPath = args[4];

            Console.WriteLine("Reading private key...");
            var privateKeyFile = new PrivateKeyFile(keyPath);

            try
            {
                Connect(host, destinationPort, localPort, username, privateKeyFile);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                Console.WriteLine("\r\nPress any key...");
                Console.ReadKey();
            }
        }

        static void Connect(string host, int destinationPort, uint localPort, string username, PrivateKeyFile privateKeyFile)
        {
            bool hasDisconnected;
            using (var client = new SshClient(host, destinationPort, username, privateKeyFile))
            {

                Console.WriteLine("Connecting...");
                client.Connect();
                var port = new ForwardedPortDynamic(localhost, localPort);
                client.AddForwardedPort(port);

                //client.ErrorOccurred += delegate (object sender, ExceptionEventArgs e)
                //{
                //    Console.WriteLine(e.Exception.ToString());
                //};

                //port.Exception += delegate (object sender, ExceptionEventArgs e)
                //{
                //    Console.WriteLine(e.Exception.ToString());
                //};

                Console.WriteLine("Opening port...");
                port.Start();
                Console.WriteLine("Connected!");
                CheckProxySettings(localPort, true);

                Console.WriteLine("Press ESC to exit.");
                hasDisconnected = RunTillEscape(client, port);


                CheckProxySettings(localPort, false);
                port.Stop();
                try { client.Disconnect(); } catch { }
            }

            if (hasDisconnected)
            {
                Console.WriteLine("Disconnected, trying to reconnect...");
                var retry = true;
                while (retry)
                {
                    retry = false;
                    try
                    {
                        Connect(host, destinationPort, localPort, username, privateKeyFile);
                    }
                    catch (Exception ex)
                    {
                        retry = true;
                    }
                }
            }
            else
            {
                Console.WriteLine("Closing connection...");
            }
        }

        static bool RunTillEscape(SshClient client, ForwardedPortDynamic port)
        {
            var hasDisconnected = false;
            do
            {
                var counter = 0;
                while (!Console.KeyAvailable && hasDisconnected == false)
                {
                    counter++;
                    if (counter >= 1000)
                    {
                        if (client.IsConnected)
                        {
                            counter = 0;
                            try
                            {
                                client.RunCommand("cd ~");
                                CheckProxySettings(port.BoundPort, true);
                            }
                            catch
                            {
                                port.Stop();
                                try { client.Disconnect(); } catch { }
                            }
                        }
                        else
                        {
                            hasDisconnected = true;
                        }
                    }
                    Thread.Sleep(10);
                }
                if (hasDisconnected) break;
            } while (Console.ReadKey(true).Key != ConsoleKey.Escape);
            return hasDisconnected;
        }

        static void CheckProxySettings(uint proxyPort, bool proxyEnabled)
        {
            const string userRoot = "HKEY_CURRENT_USER";
            const string subkey = "Software\\Microsoft\\Windows\\CurrentVersion\\Internet Settings";
            const string keyName = userRoot + "\\" + subkey;

            if (proxyEnabled)
            {
                Registry.SetValue(keyName, "ProxyServer", $"socks={localhost}:{proxyPort}");
            }
            else
            {
                using (var key = Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\Internet Settings", true))
                {
                    if (key != null)
                    {
                        key.DeleteValue("ProxyServer");
                    }
                }
            }
            Registry.SetValue(keyName, "ProxyEnable", proxyEnabled ? 0x00000001 : 0x00000000);

            // These lines implement the Interface in the beginning of program 
            // They cause the OS to refresh the settings, causing IP to realy update
            InternetSetOption(IntPtr.Zero, INTERNET_OPTION_SETTINGS_CHANGED, IntPtr.Zero, 0);
            InternetSetOption(IntPtr.Zero, INTERNET_OPTION_REFRESH, IntPtr.Zero, 0);
        }

        static void Usage()
        {
            Console.WriteLine("Usage: \r\n\tOrcrist.exe <host> <port> <local port> <username> <private key path>\r\n\r\nPress any key...");
            Console.ReadKey();
            Environment.Exit(0);
        }
    }
}
