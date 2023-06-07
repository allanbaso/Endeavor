using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;


using AutConnMgrTypeLibrary;
using AutConnListTypeLibrary;
using AutOIATypeLibrary;
using AutPSTypeLibrary;
using AutSessTypeLibrary;
using AutWinMetricsTypeLibrary;
using AutXferTypeLibrary;
using AutScreenDescTypeLibrary;
using AutScreenRecoTypeLibrary;
using Endeavor.Properties;

namespace Endeavor
{
    public partial class Form1 : Form
    {
        char SessionLetter;
        AutPS PresSpace = new AutPS();

        public Form1()
        {
            InitializeComponent();
        }
        public void beginSession()
        {
            //This sample assumes that the WS file host.ws exists in the 
            //user data directory.
            //string WorkStationProfile = @"PROFILE=CIBC.ws";
            string WorkStationProfile = @"PROFILE=" + textBox4.Text + "";

            string SessLetterName = "CONNNAME=";
            
            AutConnMgr ConnMgr = new AutConnMgr();
            //Find the first free session letter and request to start that session letter
            ((IAutConnList)ConnMgr.autECLConnList).Refresh();
            for (SessionLetter = 'A'; SessionLetter <= 'Z'; SessionLetter++)
            {
                Object Session = ((IAutConnList)ConnMgr.autECLConnList).FindConnectionByName(SessionLetter.ToString());
                if (Session == null)
                    break;
            }
            string SessionString = WorkStationProfile + " " + SessLetterName + SessionLetter;
            try
            {
                ConnMgr.StartConnection(SessionString);
            }
            catch (System.Exception E)
            {
                MessageBox.Show("Could not start the session" + E.Message);
                return;
            }
            //No exeception occured. Wait for the session to start up.
            int Seconds = 0;//10 seconds
            while (Seconds <= 10)//If the session did not start in 10 seconds get out.
            {
                ((IAutConnList)ConnMgr.autECLConnList).Refresh();
                Object Session = ((IAutConnList)ConnMgr.autECLConnList).FindConnectionByName(SessionLetter.ToString());
                if (Session != null)//Session started properly
                {
                    Session = null;
                    break;
                }
                System.Threading.Thread.Sleep(1000);
                Seconds++;
            }
            ConnMgr = null;
        }
        public bool searchText(string text, AutPS con)
        {
            AutPS PresSpace = con;
            string SearchText = text;
            Object Row = 1, Col = 1;

            if (PresSpace.SearchText(SearchText, AutPSTypeLibrary.PsDir.pcNoDirection, ref Row, ref Col))
                return true;
            else
                return false;
        }
        public int searchText(string text, AutPS con, int i)
        {
            AutPS PresSpace = con;
            string SearchText = text;
            Object Row = 1, Col = 1;

            if (PresSpace.SearchText(SearchText, AutPSTypeLibrary.PsDir.pcNoDirection, ref Row, ref Col))
                if (i == 1)
                    return int.Parse(Row.ToString());
                else
                    return int.Parse(Col.ToString());
                    
            else
                return 0;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            Settings.Default.path = this.textBox4.Text;
            Settings.Default.id = this.textBox1.Text;
            Settings.Default.Save();

            beginProcedure();
        }
        public void beginProcedure()
        {
            
            try
            {
                beginSession();
                //AutPS PresSpace = new AutPS();
                PresSpace.SetConnectionByName(SessionLetter.ToString());
                PresSpace.WaitForString("CLIENT", 0, 0, 100000, true, false);

                AutWinMetrics SessionWinmetrics = new AutWinMetrics();
                SessionWinmetrics.SetConnectionByName("A");
                SessionWinmetrics.Maximized = true;
                PresSpace.WaitForString("CLIENT", 0, 0, 10000, true, false);
                PresSpace.SendKeys("nsm02[enter]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                PresSpace.WaitForString("Userid", 0, 0, 10000, true, false);
                PresSpace.SendKeys(textBox1.Text.Trim() + "[tab]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                PresSpace.SendKeys(textBox2.Text.Trim(), PresSpace.CursorPosRow, PresSpace.CursorPosCol);

                PresSpace.SendKeys("[enter]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                PresSpace.SendKeys("[enter]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                if (searchText("DEV", PresSpace))
                {
                    int row = searchText("DEV", PresSpace, 1);
                    int col = searchText("DEV", PresSpace, 2);
                   // PresSpace.SendKeys("s[enter]", row, col - 3);
                    PresSpace.SendKeys("t[enter]", row, col - 21);
                    PresSpace.SendKeys("s[enter]", row, col - 21);
                    PresSpace.WaitForString("ENTER USER", 0, 0, 10000, true, false);
                    if (searchText("ENTER USER", PresSpace))
                    {
                        PresSpace.SendKeys(textBox1.Text.Trim() + "[enter]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                    }
                    PresSpace.SendKeys(textBox2.Text.Trim(), PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                    if (comboBox1.Text.Equals("Libman1"))
                    {
                        PresSpace.SendKeys("ISPFPROC", 10, 20);
                    }
                    else
                    {
                        PresSpace.SendKeys("ISPFDD26", 10, 20);
                    }
                    PresSpace.SendKeys("[enter]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                    PresSpace.WaitForString("***", 0, 0, 10000, true, false);
                    PresSpace.SendKeys("[enter]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                    PresSpace.WaitForString("***", 0, 0, 8000, true, false);
                    PresSpace.SendKeys("[enter]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                    PresSpace.WaitForString("Master Application Menu", 0, 0, 100000, true, false);
                    if (searchText("Master Application Menu", PresSpace))
                    {
                        PresSpace.SendKeys("97[enter]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                        PresSpace.WaitForString("Endevor", 0, 0, 10000, true, false);
                        PresSpace.SendKeys("e[enter]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                        PresSpace.WaitForString("environment", 0, 0, 10000, true, false);
                        if (!searchText("environment", PresSpace))//no se que hace
                        {
                            PresSpace.SendKeys("[enter]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                            PresSpace.WaitForString("environment", 0, 0, 10000, true, false);
                        }
                        PresSpace.SendKeys("2[enter]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                        PresSpace.WaitForString("DEFAULTS", 0, 0, 10000, true, false);
                        PresSpace.SendKeys("3[enter]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                        PresSpace.WaitForString("Batch Options Menu", 0, 0, 10000, true, false);
                        PresSpace.SendKeys("Fernando", 14, 19);
                        PresSpace.SendKeys("N", 12, 50);
                        PresSpace.SendKeys("1[enter]", 2, 15);
                        PresSpace.WaitForString("SCL GENERATION", 0, 0, 7000, true, false);
                        PresSpace.SendKeys("6[enter]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                        PresSpace.WaitForString("DELETE ELEMENTS", 0, 0, 10000, true, false);
                        string linea = richTextBox1.Text;
                        linea = linea.Replace("              ", "!");
                        linea = linea.Replace("             ", "!");
                        linea = linea.Replace("            ", "!");
                        linea = linea.Replace("           ", "!");
                        linea = linea.Replace("          ", "!");
                        linea = linea.Replace("         ", "!");
                        linea = linea.Replace("        ", "!");
                        linea = linea.Replace("       ", "!");
                        linea = linea.Replace("      ", "!");
                        linea = linea.Replace("     ", "!");
                        linea = linea.Replace("    ", "!");
                        linea = linea.Replace("   ", "!");
                        linea = linea.Replace("  ", "!");
                        linea = linea.Replace(" ", "!");
                        linea = linea.Replace("\t", "!");
                        linea = linea.Replace("!!!!", "!");
                        linea = linea.Replace("!!!", "!");
                        linea = linea.Replace("!!", "!");
                        //linea = linea.Trim();
                        string[] words = linea.Split('\n');
                        string[] variables;
                        bool completed = false;
                        for (int i = 0; i < words.Length; i++)
                        {
                            if (words[i].Contains('!'))
                            {
                                variables = words[i].Split('!');
                                PresSpace.SendKeys("[erase eof]", 10, 21);
                                PresSpace.SendKeys(variables[2].Trim(), 10, 21);
                                PresSpace.SendKeys("[erase eof]", 11, 21);
                                PresSpace.SendKeys(variables[4].Trim(), 11, 21);
                                PresSpace.SendKeys("[erase eof]", 12, 21);
                                PresSpace.SendKeys(variables[5].Trim(), 12, 21);
                                PresSpace.SendKeys("[erase eof]", 13, 21);
                                PresSpace.SendKeys(variables[0].Trim(), 13, 21);
                                PresSpace.SendKeys("[erase eof]", 14, 21);
                                PresSpace.SendKeys(variables[1].Trim(), 14, 21);
                                PresSpace.SendKeys("[erase eof]", 15, 21);
                                PresSpace.SendKeys(variables[3].Trim(), 15, 21);
                                PresSpace.SendKeys("Y", 10, 72);
                                PresSpace.SendKeys("prodsupp", 7, 12);
                                PresSpace.SendKeys(textBox3.Text.Trim(), 7, 41);
                                PresSpace.SendKeys("[enter]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                                PresSpace.WaitForString(variables[0].Trim(), 0, 0, 20000, true, false);
                                if (searchText(variables[0].Trim(), PresSpace))
                                {
                                    PresSpace.SendKeys("#", 07, 02);
                                    PresSpace.SendKeys("[enter]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                                    PresSpace.WaitForString("SCL Generated", 0, 0, 40000, true, false);
                                    if (searchText("WRITTEN", PresSpace))
                                    {
                                        PresSpace.SendKeys("[pf3]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                                        PresSpace.WaitForString("DELETE ELEMENTS", 0, 0, 50000, true, false);
                                    }
                                    else
                                    {
                                        PresSpace.SendKeys("[pf3]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                                        PresSpace.WaitForString("DELETE ELEMENTS", 0, 0, 50000, true, false);
                                    }
                                }
                                if (i == words.Length - 1)
                                    completed = true;
                            }
                        }
                        MessageBox.Show("Proccess Finished, Please Continue with the task");
                        /*if (completed)
                        {
                            PresSpace.WaitForString("FROM ENDEVOR", 0, 0, 10000, true, false);
                            PresSpace.SendKeys("[pf3]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                            PresSpace.WaitForString("SCL GENERATION", 0, 0, 10000, true, false);
                            PresSpace.SendKeys("[pf3]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                            PresSpace.WaitForString("Batch Options Menu", 0, 0, 10000, true, false);
                            PresSpace.SendKeys("3[enter]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                            PresSpace.WaitForString("SUBMITTED", 0, 0, 10000, true, false);
                            if (searchText("SUBMITTED", PresSpace))
                            {
                                PresSpace.SendKeys("[pf3]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                                PresSpace.WaitForString("Batch Options Menu", 0, 0, 10000, true, false);
                                PresSpace.SendKeys("[pf3]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                                PresSpace.WaitForString("DEFAULTS", 0, 0, 10000, true, false);
                                PresSpace.SendKeys("[pf3]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                                PresSpace.WaitForString("environment", 0, 0, 10000, true, false);
                                PresSpace.SendKeys("[pf3]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                                PresSpace.WaitForString("Endevor", 0, 0, 10000, true, false);
                                PresSpace.SendKeys("[pf3]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                                PresSpace.WaitForString("Master Application Menu", 0, 0, 10000, true, false);
                                PresSpace.SendKeys("70[enter]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                                PresSpace.WaitForString("OPTION MENU", 0, 0, 10000, true, false);
                                PresSpace.SendKeys("st[enter]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                                PresSpace.WaitForString("ALL CLASSES", 0, 0, 10000, true, false);
                                PresSpace.SendKeys("owner " + textBox1.Text.Trim()+ "[enter]", 4, 21);
                                PresSpace.WaitForString("ALL CLASSES", 0, 0, 10000, true, false);
                                PresSpace.SendKeys("pre *[enter]", 4, 21);
                                PresSpace.WaitForString("ALL CLASSES", 0, 0, 10000, true, false);
                                PresSpace.SendKeys("o[enter]", 7, 2);
                                PresSpace.WaitForString("COMMAND ISSUED", 0, 0, 10000, true, false);
                                PresSpace.SendKeys("o[enter]", 4, 21);
                                PresSpace.WaitForString("ALL CLASSES", 0, 0, 10000, true, false);
                                PresSpace.SendKeys("3[enter]", 7, 54);
                            }
                            while (searchText(textBox1.Text, PresSpace))
                            {
                                PresSpace.Wait(10000);
                                PresSpace.SendKeys("[enter]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                            }
                            PresSpace.SendKeys("[pf3]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                            PresSpace.WaitForString("OPTION MENU", 0, 0, 10000, true, false);
                            PresSpace.SendKeys("[pf3]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                            PresSpace.WaitForString("Master Application Menu", 0, 0, 10000, true, false);
                            PresSpace.SendKeys("74[enter]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                            PresSpace.WaitForString("CLASS3", 0, 0, 10000, true, false);
                            PresSpace.SendKeys("1[enter]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                            PresSpace.WaitForString("Primary Selection", 0, 0, 10000, true, false);
                            PresSpace.SendKeys(textBox1.Text.Trim() + "*", 4, 21);
                            PresSpace.SendKeys("-0", 15, 21);
                            PresSpace.SendKeys("[enter]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                            PresSpace.WaitForString("Sysout Selection List", 0, 0, 10000, true, false);
                           
                            PresSpace.SendKeys("[pf3]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                            PresSpace.SendKeys("[pf3]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                            PresSpace.SendKeys("[pf3]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                            PresSpace.SendKeys("X[enter]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                            PresSpace.SendKeys("logoff[enter]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                            PresSpace.SendKeys("[pf3]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                            PresSpace.SendKeys("1[enter]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                        }*/

                    }
                }
                else
                {
                    MessageBox.Show("No DEV Session, Please Review");
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Error " + e.Message);
            }
        }
        private void ColocaOption(int val, string valor, int posicion)
        {
            //AutPS PresSpace = new AutPS();

            int fila = searchText("" + valor + "", PresSpace, 1);
            int columna = searchText("" + valor + "", PresSpace, 2);
            PresSpace.SendKeys("" + val + "[enter]", fila, columna + posicion);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.textBox1.Text = Settings.Default.id.ToString();
            this.textBox4.Text = Settings.Default.path.ToString();
        }
    }

}
