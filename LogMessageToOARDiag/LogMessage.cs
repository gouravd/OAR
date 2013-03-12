using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.Serialization.Formatters.Binary;
using System.IO;


namespace LogMessageToOARDiag
{
    public class LogMessage
    {
        //private System.Exception paramEx;
        string SubDir = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\" + "OARsFiles";
        public int fnLogExceptions(System.Exception ex, string logfileName)
        {
            try
            {
                FileInfo fDiagInfo = new FileInfo(SubDir + "\\" + logfileName);
                lock(this)
                {
                    if (!fDiagInfo.Exists)
                    {
                        StreamWriter OARDiagStrWriter = fDiagInfo.CreateText();

                        OARDiagStrWriter.WriteLine();
                        OARDiagStrWriter.WriteLine(SubDir + "\\" + logfileName);
                        OARDiagStrWriter.WriteLine("=========================");
                        OARDiagStrWriter.WriteLine(System.DateTime.Now);
                        OARDiagStrWriter.WriteLine("-------------------------");
                        OARDiagStrWriter.Write("Source: " + ex.Source);
                        OARDiagStrWriter.WriteLine("");
                        OARDiagStrWriter.Write("Message: " + ex.Message);
                        OARDiagStrWriter.WriteLine("");
                        OARDiagStrWriter.Write("Stack: " + ex.StackTrace);
                        OARDiagStrWriter.WriteLine("*************************");
                        OARDiagStrWriter.WriteLine("=========================");

                        OARDiagStrWriter.Close();
                    }
                    else
                    {
                        StreamWriter OARDiagStrWriter = fDiagInfo.AppendText();

                        OARDiagStrWriter.WriteLine();
                        OARDiagStrWriter.WriteLine(SubDir + "\\" + logfileName);
                        OARDiagStrWriter.WriteLine("=========================");
                        OARDiagStrWriter.WriteLine(System.DateTime.Now);
                        OARDiagStrWriter.WriteLine("-------------------------");
                        OARDiagStrWriter.Write("Source: " + ex.Source);
                        OARDiagStrWriter.WriteLine("");
                        OARDiagStrWriter.Write("Message: " + ex.Message);
                        OARDiagStrWriter.WriteLine("");
                        OARDiagStrWriter.Write("Stack: " + ex.StackTrace);
                        OARDiagStrWriter.WriteLine("*************************");
                        OARDiagStrWriter.WriteLine("=========================");

                        OARDiagStrWriter.Close();
                    }
                    return 0x11111111;
                }
            }

            catch
            {
                return 0x013;
            }
        }
        public int fnLogStartUPMessage(string logfileName)
        {
            try
            {
                FileInfo fDiagInfo = new FileInfo(SubDir + "\\" + logfileName);

                if (!fDiagInfo.Exists)
                {
                    StreamWriter OARDiagStrWriter = fDiagInfo.CreateText();
                    
                    OARDiagStrWriter.WriteLine();
                    OARDiagStrWriter.WriteLine(SubDir + "\\" + logfileName);
                    OARDiagStrWriter.WriteLine("=========================");
                    OARDiagStrWriter.WriteLine(System.DateTime.Now);
                    OARDiagStrWriter.WriteLine("-------------------------");
                    OARDiagStrWriter.WriteLine("Machine Name: " + System.Environment.MachineName);
                    OARDiagStrWriter.WriteLine("OS: " + System.Environment.OSVersion);
                    OARDiagStrWriter.WriteLine("UserName: " + System.Environment.UserName);
                    OARDiagStrWriter.WriteLine("Interactive Mode: " + System.Environment.UserInteractive);
                    OARDiagStrWriter.WriteLine("*************************");
                    OARDiagStrWriter.WriteLine("=========================");

                    OARDiagStrWriter.Close();
                }
                else
                {
                    StreamWriter OARDiagStrWriter = fDiagInfo.AppendText();

                    OARDiagStrWriter.WriteLine();
                    OARDiagStrWriter.WriteLine(SubDir + "\\" + logfileName);
                    OARDiagStrWriter.WriteLine("=========================");
                    OARDiagStrWriter.WriteLine(System.DateTime.Now);
                    OARDiagStrWriter.WriteLine("-------------------------");
                    OARDiagStrWriter.WriteLine("Machine Name: " + System.Environment.MachineName);
                    OARDiagStrWriter.WriteLine("OS: " + System.Environment.OSVersion);
                    OARDiagStrWriter.WriteLine("UserName: " + System.Environment.UserName);
                    OARDiagStrWriter.WriteLine("Interactive Mode: " + System.Environment.UserInteractive);
                    OARDiagStrWriter.WriteLine("*************************");
                    OARDiagStrWriter.WriteLine("=========================");

                    OARDiagStrWriter.Close();
                }
                return 0x11111111;
            }
            catch
            {
                return 0x014;
            }

        }
        public void fnCreateWordList(ref FileInfo fInfo)
        {
            StreamWriter OARSw = fInfo.CreateText();

            //Adding some default words
            OARSw.WriteLine("attach");
            OARSw.WriteLine("attached");
            OARSw.WriteLine("attaching");
            OARSw.WriteLine("attachment");
            OARSw.WriteLine("enclose");
            OARSw.WriteLine("enclosing");
            OARSw.WriteLine("enclosure");

            OARSw.WriteLine("SUB:-:attach");
            OARSw.WriteLine("SUB:-:attached");
            OARSw.WriteLine("SUB:-:attaching");
            OARSw.WriteLine("SUB:-:attachment");
            OARSw.WriteLine("SUB:-:enclose");
            OARSw.WriteLine("SUB:-:enclosing");
            OARSw.WriteLine("SUB:-:enclosure");

            OARSw.WriteLine("SIZE:-:100000");

            OARSw.Close();

        }
        public void fnLogSaveAttachment(ref FileInfo fInfo,System.Exception exption)
        {
            lock (this)
            {
                TextWriter OARTw;

                if (!fInfo.Exists)
                {
                    OARTw = fInfo.CreateText();
                }
                else
                {
                    OARTw = fInfo.AppendText();
                }

                OARTw.WriteLine(System.DateTime.Now);
                OARTw.WriteLine("======================================");

                OARTw.WriteLine(exption.Message);
                OARTw.WriteLine("\n\n");
                OARTw.Close();
            }
                      
        }
    }
}
