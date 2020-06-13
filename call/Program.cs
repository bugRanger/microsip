using System;
using AutoItX3Lib;

namespace call
{
    class Program
    {
        AutoItX3 softPhone = new AutoItX3();
        AutoItX3 virtSound = new AutoItX3();
        public static string title, account, transport, server, port, fileForSound = "";
        public static string titleAccount = "Account";
        public static string titleSettings = "Settings";
        public static string action = "";
        static void Main(string[] args)
        {
            try
            {
                title = args[0];
                action = args[1];
                account = args[2];
                Program program = new Program();
                switch (action)
                {
                    case "runAndSetup":
                        transport = args[3];
                        server = args[4];
                        port = args[5];
                        program.Run();
                        program.Register(account, transport, server, port);
                        program.Setup();
                        break;
                    case "call":
                        title = title.Replace("-", " - ");
                        fileForSound = args[3];
                        program.Call(account, fileForSound);
                        break;
                    case "mic":
                        title = title.Replace("-", " - ");
                        program.Mic(args[3]);
                        break;
                    default:
                        break;
                }
            }
            catch(IndexOutOfRangeException)
            {
                Console.WriteLine("You should use mandatory prameters: title, action, account. Optional parameters: transport, port");
                Console.WriteLine("For example: call.exe MicroSip runAndSetup 101 tcp 172.16.0.100 5060");
                Console.WriteLine("For example: call.exe \"MicroSip - 102\" call 101");
            }
        }

        private void Mic(string v)
        {
            softPhone.WinWait(title);
            softPhone.WinActivate(title);
            softPhone.WinWaitActive(title);
            softPhone.ControlClick(title, "", "Button25");
        }

        private void Call(string number, string fileForSound)
        {
            softPhone.WinWait(title);
            softPhone.WinActivate(title);
            softPhone.WinWaitActive(title);
            softPhone.ControlClick(title, "", "Edit1");

            virtSound.Run(@"C:\Program Files (x86)\e2eSoft\VSC\VSCMain.exe", "", virtSound.SW_SHOW);
            virtSound.WinWait("e2eSoft VSC");
            virtSound.WinActivate("e2eSoft VSC");
            virtSound.WinWaitActive("e2eSoft VSC");
            virtSound.ControlCommand("e2eSoft VSC", "", "ComboBox1", "SelectString", "Microphone (e2eSoft VAudio)");
            virtSound.Sleep(300);
            virtSound.ControlCommand("e2eSoft VSC", "", "ComboBox2", "SelectString", "Microsoft Sound Mapper");
            virtSound.Sleep(200);
            virtSound.ControlClick("e2eSoft VSC", "", "Button2");
            

            softPhone.WinWait(title);
            softPhone.WinActivate(title);
            softPhone.WinWaitActive(title);
            softPhone.ControlClick(title, "", "Edit1");
            softPhone.Sleep(100);
            softPhone.Send(number);
            softPhone.Sleep(100);
            
            softPhone.ControlClick(title, "", "Button19");

            System.Diagnostics.Process.Start(fileForSound);

            softPhone.WinWait(title);
            softPhone.WinActivate(title);
            softPhone.WinWaitActive(title);

            softPhone.Sleep(10000);
            softPhone.Send("{ENTER}");

            virtSound.WinActivate("e2eSoft VSC");
            virtSound.WinWaitActive("e2eSoft VSC");
            virtSound.ControlClick("e2eSoft VSC", "", "Button2");
            virtSound.WinClose("e2eSoft VSC");
        }

        private void Run()
        {
            softPhone.Run(@"microsip.exe", "", softPhone.SW_SHOW);
            softPhone.WinWait(title);
            softPhone.WinActivate(title);
            softPhone.WinWaitActive(title);
        }
        private void Register(string number, string transport, string server, string port)
        {
            softPhone.ControlClick(title, "", "Button1");
            softPhone.Send("{DOWN}");
            softPhone.Send("{ENTER}");
            softPhone.WinWait(titleAccount);
            softPhone.WinActivate(titleAccount);
            softPhone.WinWaitActive(titleAccount);
            softPhone.ControlClick(titleAccount, "", "Edit1");
            softPhone.Send(number);
            softPhone.Sleep(100);
            softPhone.Send("{TAB}");
            softPhone.Sleep(100);
            softPhone.Send($"{server}:{port}");
            softPhone.Sleep(100);
            softPhone.Send("{TAB}");
            softPhone.Sleep(100);
            softPhone.Send("{TAB}");
            softPhone.Sleep(100);
            softPhone.Send(number);
            softPhone.Sleep(100);
            softPhone.Send("{TAB}");
            softPhone.Sleep(100);
            softPhone.Send($"{server}:{port}");
            softPhone.Sleep(100);
            softPhone.Send("{TAB}");
            softPhone.Sleep(100);
            softPhone.Send(number);
            softPhone.Sleep(100);
            softPhone.Send("{TAB}");
            softPhone.Sleep(100);
            softPhone.Send("111111");
            softPhone.Sleep(100);
            softPhone.ControlCommand(titleAccount, "", "ComboBox2", "SelectString", transport);
            softPhone.Sleep(100);
            softPhone.Send("{DOWN 2}");
            softPhone.Sleep(100);
            softPhone.ControlClick(titleAccount, "", "Edit13");
            softPhone.Send("00");
            softPhone.Sleep(100);
            softPhone.ControlClick(titleAccount, "", "Button7");
            title += $" - {number}";
            softPhone.WinWait(title);
            softPhone.WinActivate(title);
            softPhone.WinWaitActive(title);
        }
        private void Setup()
        {
            softPhone.ControlClick(title, "", "Button1");
            softPhone.Sleep(100);
            softPhone.Send("{DOWN 4}");
            softPhone.Send("{ENTER}");
            softPhone.WinWait(titleSettings);
            softPhone.WinActivate(titleSettings);
            softPhone.WinWaitActive(titleSettings);
            softPhone.ControlClick(titleSettings, "", "Button15");
            softPhone.Sleep(100);
            softPhone.ControlClick(titleSettings, "", "ComboBox5");
            softPhone.Sleep(100);
            softPhone.Send("{DOWN 2}");
            softPhone.Send("{ENTER}");
            softPhone.Sleep(100);
            softPhone.ControlClick(titleSettings, "", "Button26");
        }
    }
}
