using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
//Reference ECL types to use
using AutConnMgrTypeLibrary;
using AutConnListTypeLibrary;
using AutOIATypeLibrary;
using AutPSTypeLibrary;
using AutSessTypeLibrary;
using AutWinMetricsTypeLibrary;
using AutXferTypeLibrary;
using AutScreenDescTypeLibrary;
using AutScreenRecoTypeLibrary;
using System.IO;
using listaEndeavor.Properties;

namespace listaEndeavor
{
    public partial class Form1 : Form
    {
        //AutPS PresSpace = new AutPS();
        AutPS PresSpace;
        AutConnMgr ConnMgr;
        char SessionLetter;
        public Form1()
        {
           InitializeComponent();
           textBox1.Text = Properties.Settings.Default["id"].ToString();
           textBox4.Text = Properties.Settings.Default["path"].ToString();

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text.Equals("") | textBox2.Text.Equals("") | textBox3.Text.Equals("") | textBox4.Text.Equals(""))
            {


                MessageBox.Show("There are Empty Fields, Please Review");
            }
            else
            {
                beginProcedure();
                Properties.Settings.Default["id"] = textBox1.Text;
                Settings.Default["id"] = textBox1.Text;
                Properties.Settings.Default["path"] = textBox4.Text;
                Properties.Settings.Default.Save();
                
            }
        }
        public void beginSession()
        {
            //This sample assumes that the WS file host.ws exists in the 
            //user data directory.
            //string WorkStationProfile = @"PROFILE=CIBC.ws";
            //string WorkStationProfile = @"PROFILE=C:\Users\abarrientos2\Desktop\CIBC_Mainframe.WS";
            string WorkStationProfile = @"PROFILE="+textBox4.Text+"";

            string SessLetterName = "CONNNAME=";

            ConnMgr = new AutConnMgr();
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

        public void beginProcedure()
        {
            this.Enabled = false;
            try
            {
                beginSession();
                PresSpace = new AutPS();
                PresSpace.SetConnectionByName(SessionLetter.ToString());
                PresSpace.WaitForString("CLIENT IP", 0, 0, 30000, true, false);
                string listadoExiste = "";
                string listadoNOExiste = "";
                string listadoRepetido = "";
                AutWinMetrics SessionWinmetrics = new AutWinMetrics();
                SessionWinmetrics.SetConnectionByName("A");
                //SessionWinmetrics.Maximized = true;
                PresSpace.SendKeys("nsm02[enter]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                PresSpace.Wait(1000);
                PresSpace.WaitForString("Userid", 0, 0, 10000, true, false);
                PresSpace.SendKeys(textBox1.Text.Trim() + "[tab]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                PresSpace.SendKeys(textBox2.Text.Trim(), PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                PresSpace.SendKeys("[enter]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                PresSpace.SendKeys("[enter]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                if (searchText("DEV", PresSpace))
                {
                    int row = searchText("DEV", PresSpace, 1);
                    int col = searchText("DEV", PresSpace, 2);
                    PresSpace.SendKeys("t[enter]", row, col - 21);
                    PresSpace.SendKeys("s[enter]", row, col - 21);
                    PresSpace.WaitForString("ENTER USER", 0, 0, 10000, true, false);
                    if (searchText("ENTER LOGON", PresSpace))
                    {
                        PresSpace.SendKeys(textBox1.Text.Trim() + "[enter]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                    }
                    PresSpace.SendKeys(textBox2.Text.Trim(), PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                    if(comboBox1.Text.Equals("Libman1"))
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
                    PresSpace.SendKeys("[enter]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                    //PresSpace.WaitForString("***", 0, 0, 10000, true, false);
                    //revisar por si el enter es esperado.
                    /*if (searchText("***", PresSpace))//
                    {
                        PresSpace.SendKeys("[enter]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                        //PresSpace.SendKeys("[enter]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                        //PresSpace.WaitForString("environment", 0, 0, 10000, true, false);
                    }*/
                    //PresSpace.WaitForString("***", 0, 0, 20000, true, false);
                    //PresSpace.SendKeys("[enter]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                    PresSpace.WaitForString("Master Application Menu", 0, 0, 100000, true, false);



                    if (searchText("Master Application Menu", PresSpace))
                    {
                        PresSpace.SendKeys("97[enter]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                        PresSpace.WaitForString("Endevor", 0, 0, 10000, true, false);
                        PresSpace.SendKeys("e[enter]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                        //PresSpace.SendKeys("e[enter]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);

                        PresSpace.WaitForString("environment", 0, 0, 10000, true, false);
                        //si la posicion del option cambia
                        ColocaOption(2, "OPTION", 13);
                        //
                        //PresSpace.SendKeys("2[enter]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                        PresSpace.WaitForString("DEFAULTS", 0, 0, 10000, true, false);
                        ColocaOption(1, "OPTION", 13);
                        //PresSpace.SendKeys("1[enter]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                        PresSpace.WaitForString("Display", 0, 0, 10000, true, false);
                        ColocaOption(1, "OPTION", 13);
                        //PresSpace.SendKeys("1[enter]", 22, 15);
                        PresSpace.WaitForString("FROM LOCATION", 0, 0, 10000, true, false);
                        string linea = richTextBox1.Text;
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
                        linea = linea.Replace("!!", "!");
                        linea = linea.Replace("!!!", "!");
                        string[] words = linea.Split('\n');
                        string[] variables;
                        for (int i = 0; i < words.Length; i++)
                        {
                            if (words[i].Contains('!'))
                            {
                                variables = words[i].Split('!');
                                //PresSpace.SendKeys(variables[2].Trim() + "     ", 13, 21);//////variar columna de posicionamiento
                                int fila = searchText("ENVIRONMENT", PresSpace, 1);
                                int colmna = searchText("ENVIRONMENT", PresSpace, 2);
                                
                                PresSpace.SendKeys("N", 14, 71);
                                colmna += 17;
                                PresSpace.SendKeys("[erase eof]", fila, colmna);//12,21
                                PresSpace.SendKeys(variables[2].Trim() + "", fila, colmna);//system
                                PresSpace.SendKeys("[erase eof]", fila+1, colmna);
                                PresSpace.SendKeys(variables[4].Trim() + "", fila+1, colmna);//subsystem
                                PresSpace.SendKeys("[erase eof]", fila+2, colmna);
                                PresSpace.SendKeys(variables[5].Trim() + "", fila+2, colmna);//element
                                PresSpace.SendKeys("[erase eof]", fila+3, colmna);
                                PresSpace.SendKeys(variables[0].Trim() + "", fila+3, colmna);//type
                                PresSpace.SendKeys("[erase eof]", fila+4, colmna);
                                PresSpace.SendKeys(variables[1].Trim() + "", fila+4, colmna);//stage
                                PresSpace.SendKeys("[erase eof]", fila + 5, colmna);
                                PresSpace.SendKeys(variables[3].Trim() + "", fila + 5, colmna);//stage
           
                                PresSpace.SendKeys("[enter]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                                PresSpace.WaitForString(variables[0].Trim(), 0, 0, 10000, true, false);
                                if (chequearRepetidos(words[i], words, i+1))
                                {
                                    listadoRepetido = listadoRepetido + variables[0] + "\t" + variables[1] + "\t" + variables[2] + "\t" + variables[3] + "\t" + variables[4] + "\t" + variables[5] + "\n";
                                }
                                if (searchText("NO ELEMENTS SELECTED", PresSpace) | searchText("NO ENVIRONMENT MATCH", PresSpace))
                                {
                                    listadoNOExiste = listadoNOExiste + variables[0] + "\t" + variables[1] + "\t" + variables[2] + "\t" + variables[3] + "\t" + variables[4] + "\t" + variables[5] + "\n";
                                }
                                else
                                {
                                    listadoExiste = listadoExiste + variables[0] + "\t" + variables[1] + "\t" + variables[2] + "\t" + variables[3] + "\t" + variables[4] + "\t" + variables[5] + "\n";
                                    PresSpace.SendKeys("[pf3]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                                    PresSpace.WaitForString("FROM LOCATION", 0, 0, 10000, true, false);
                                }
                            }
                        }
                        // create a writer and open the file
                        if (!listadoExiste.Trim().Equals(""))
                        {
                            TextWriter tw = new StreamWriter(textBox3.Text.Trim() + " ListadoExistente.txt");

                            // write a line of text to the file
                            tw.WriteLine(listadoExiste);

                            // close the stream
                            tw.Close();
                        }
                        if (!listadoNOExiste.Trim().Equals(""))
                        {
                            // create a writer and open the file
                            TextWriter tw = new StreamWriter(textBox3.Text.Trim() + " ListadoNOExistente.txt");

                            // write a line of text to the file
                            tw.WriteLine(listadoNOExiste);

                            // close the stream
                            tw.Close();
                        }
                        if (!listadoRepetido.Trim().Equals(""))
                        {
                            TextWriter tw = new StreamWriter(textBox3.Text.Trim() + " ListadoRepetidos.txt");

                            // write a line of text to the file
                            tw.WriteLine(listadoRepetido);

                            // close the stream
                            tw.Close();
                        }
                        PresSpace.SendKeys("[pf3]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                        PresSpace.SendKeys("[pf3]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                        PresSpace.SendKeys("[pf3]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                        PresSpace.SendKeys("[pf3]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                        PresSpace.SendKeys("[pf3]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                        PresSpace.SendKeys("[pf3]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                        PresSpace.SendKeys("[pf3]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);

                        //PresSpace.WaitForString("Master Application Menu", 0, 0, 10000, true, false);
                        //PresSpace.SendKeys("X[enter]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                        //if para la pantalla de salida
                        //Specify Disposition of Log Data Set
                        if (searchText("Specify Disposition of Log Data Set", PresSpace))
                        {
                            //PresSpace.WaitForString("Specify Disposition of Log Data Set", 0, 0, 10000, true, false);
                            PresSpace.WaitForString("Process Option", 0, 0, 10000, true, false);
                            PresSpace.SendKeys("4", 5, 25);
                            PresSpace.SendKeys("[enter]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                        }
                        PresSpace.WaitForString("READY", 0, 0, 10000, true, false);
                        PresSpace.SendKeys("logoff[enter]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                        PresSpace.WaitForString("SUPERSESSION", 0, 0, 10000, true, false);
                        PresSpace.SendKeys("[pf3]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                        PresSpace.SendKeys("1[enter]", PresSpace.CursorPosRow, PresSpace.CursorPosCol);
                        MessageBox.Show("Process Finished Successfully");
                    }
                    else
                    {
                        MessageBox.Show("No Master Application Menu, Please Review");
                    }
                }
                else
                {
                    MessageBox.Show("No DEV Session, Please Review");
                }
                this.Enabled = true;
                
            }
            catch (Exception e)
            {
                MessageBox.Show("Error " + e.Message);
            }
            PresSpace = null;
            ConnMgr = null;
                
        }
        public bool chequearRepetidos(string linea, string[] arreglo, int i)
        {
            bool resultado = false;
            for (; i < arreglo.Length; i++)
            {
                if (arreglo[i].Trim().Equals(linea.Trim()))
                {
                    resultado = true;
                    break;
                }
            }
            return resultado;
        }
        private void ColocaOption(int val, string valor, int posicion)
        {
            //AutPS PresSpace = new AutPS();
            //PresSpace = new AutPS();
            int fila = searchText("" + valor + "", PresSpace, 1);
            int columna = searchText("" + valor + "", PresSpace, 2);
            PresSpace.SendKeys("" + val + "[enter]", fila, columna + posicion);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            textBox1.Text = Properties.Settings.Default.id.ToString();
            textBox4.Text = Properties.Settings.Default.path.ToString();
        }
    }
}
