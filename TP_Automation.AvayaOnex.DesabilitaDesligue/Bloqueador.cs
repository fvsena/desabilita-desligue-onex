using Castle.Core.Logging;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using TestStack.White;
using TestStack.White.Configuration;
using TestStack.White.Factory;
using TestStack.White.UIItems;
using TestStack.White.UIItems.WindowItems;

namespace TP_Automation.AvayaOnex.DesabilitaDesligue
{
    public class Bloqueador
    {
        #region Métodos Externos
        [DllImport("user32.dll", SetLastError = true)]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("kernel32.dll")]
        static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

        static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
        const UInt32 SWP_NOSIZE = 0x0001;
        const UInt32 SWP_NOMOVE = 0x0002;
        const UInt32 SWP_SHOWWINDOW = 0x0040;

        const int SW_HIDE = 0;
        const int SW_SHOW = 5;

        delegate void WinEventDelegate(IntPtr hWinEventHook, uint eventType,
        IntPtr hwnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime);

        [DllImport("user32.dll")]
        static extern IntPtr SetWinEventHook(uint eventMin, uint eventMax, IntPtr
           hmodWinEventProc, WinEventDelegate lpfnWinEventProc, int idProcess,
           uint idThread, uint dwFlags);

        [DllImport("user32.dll")]
        static extern bool UnhookWinEvent(IntPtr hWinEventHook);

        const uint EVENT_OBJECT_LOCATIONCHANGE = 0x800B;
        const uint EVENT_OBJECT_CREATE = 0x8000;
        const uint WINEVENT_OUTOFCONTEXT = 0;
        const uint WINEVENT_OBJECT_VALUES_CHANGE = 0x800E;
        const uint EVENT_OBJECT_NAMECHANGE = 0x800C;
        public static IntPtr ProgramHandle { get; set; }

        // Need to ensure delegate is not collected while we're using it,
        // storing it in a class field is simplest way to do this.
        static WinEventDelegate procMoveDelegate = new WinEventDelegate(WinEventProcMove);
        static WinEventDelegate procNameChangeDelegate = new WinEventDelegate(WinEventProcNameChange);


        #endregion

        #region Construtores
        public Bloqueador()
        {
            CoreAppXmlConfiguration.Instance.LoggerFactory = new ConsoleFactory(LoggerLevel.Error);
        }
        #endregion

        #region Propriedades privadas
        private static bool FormExibido = false;
        private static frmBloqueado Bloqueado;
        private static IUIItem objeto = null;
        private Window window = null;

        private static string processName = "";
        private static string windowName = "";
        private static int processId;
        private static string id = "";
        private static int tipo;
        private static IntPtr objectHandle = IntPtr.Zero;
        private static TipoMonitoramento tipoMonitoramentoProcesso;

        private static bool EmExecucao = false;
        #endregion

        #region Métodos públicos
        /// <summary>
        /// Realiza o bloqueio de um controle ao posicionar um form invisivel em frente ao mesmo
        /// </summary>
        /// <param name="tituloJanela">Nome da janela que contém o elemento</param>
        /// <param name="nomeObjeto">Nome do elemento que será bloqueado</param>
        public void BloquearControle(string nomeProcesso, string tituloJanela, string nomeObjeto, TipoMonitoramento tipoMonitoramento)
        {
            bool carregado = false;
            List<Process> processes = null;

            IntPtr hhook;
            IntPtr hhookNext;

            while (!carregado)
            {
                try
                {
                    Thread.Sleep(10000);
                    processes = Process.GetProcesses().Where(p => p.ProcessName == nomeProcesso).ToList();
                    while (processes.Count < 1 && processes.FirstOrDefault().MainWindowHandle != IntPtr.Zero)
                    {
                        processes = Process.GetProcesses().Where(p => p.ProcessName == nomeProcesso).ToList();
                    }

                    ProgramHandle = processes[0].MainWindowHandle;
                    processId = processes[0].Id;

                    Console.WriteLine(ProgramHandle);
                    Console.WriteLine(processId);

                    processName = nomeProcesso;
                    windowName = tituloJanela;
                    id = nomeObjeto;
                    tipoMonitoramentoProcesso = tipoMonitoramento;

                    carregado = true;
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Não foi possivel carregar o sistema a ser monitorado: "  + ex.Message);
                }
            }

            switch (tipoMonitoramentoProcesso)
            {
                case TipoMonitoramento.EVENT_OBJECT_NAME_CHANGED:
                    hhook = SetWinEventHook(EVENT_OBJECT_NAMECHANGE, EVENT_OBJECT_NAMECHANGE, IntPtr.Zero,
                    procNameChangeDelegate, 0, 0, WINEVENT_OUTOFCONTEXT);

                    hhookNext = SetWinEventHook(EVENT_OBJECT_LOCATIONCHANGE, EVENT_OBJECT_LOCATIONCHANGE, IntPtr.Zero,
                    procMoveDelegate, 0, 0, WINEVENT_OUTOFCONTEXT);

                    break;
                case TipoMonitoramento.EVENT_OBJECT_LOCATION_CHANGED:
                    hhook = SetWinEventHook(EVENT_OBJECT_LOCATIONCHANGE, EVENT_OBJECT_LOCATIONCHANGE, IntPtr.Zero,
                    procMoveDelegate, 0, 0, WINEVENT_OUTOFCONTEXT);
                    break;
                default:
                    break;
            }

            while (true)
            {
                //System.Windows.Forms.MessageBox.Show("nun é pussivo");
                Thread.Sleep(100);
                System.Windows.Forms.Application.DoEvents();
            }

            UnhookWinEvent(hhook);
        }


        static void WinEventProcNameChange(IntPtr hWinEventHook, uint eventType,
            IntPtr hwnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
        {
            try
            {
                Console.WriteLine("Name changed HWND: " + hwnd.ToString());
                if (tipoMonitoramentoProcesso.Equals(TipoMonitoramento.EVENT_OBJECT_NAME_CHANGED) && objectHandle.Equals(IntPtr.Zero))
                {
                    GerenciaObjectHandle();
                }

                if (tipoMonitoramentoProcesso.Equals(TipoMonitoramento.EVENT_OBJECT_NAME_CHANGED) && objectHandle != IntPtr.Zero && objectHandle == hwnd)
                {
                    GerenciaBloqueio();
                }
            }
            catch
            {
            }

        }

        static void WinEventProcMove(IntPtr hWinEventHook, uint eventType,
            IntPtr hwnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
        {
            try
            {
                if (hwnd == ProgramHandle)
                {
                    GerenciaBloqueio();
                }
            }
            catch
            {
            }
            
        }

        private static void GerenciaObjectHandle()
        {
            try
            {
                if (EmExecucao)
                {
                    return;
                }

                EmExecucao = true;

                var application = Application.Attach(processId);
                Window window = application.GetWindow(windowName, InitializeOption.NoCache);

                objeto = window.Items.Find(x => !string.IsNullOrEmpty(x.Id) && x.Id.ToUpper().Equals(id.ToUpper()));

                if (objeto == null)
                {
                    objeto = window.Items.Find(x => !string.IsNullOrEmpty(x.Name) && x.Name.ToUpper().Equals(id.ToUpper()));
                }
                
                if(objeto != null)
                {
                    objectHandle = (IntPtr)objeto.AutomationElement.Current.NativeWindowHandle;
                }
            }
            catch
            {
            }

            EmExecucao = false;

            return;
        }

        private static void GerenciaBloqueio()
        {
            try
            {
                if (EmExecucao)
                {
                    return;
                }

                EmExecucao = true;

                ExibirBloqueio(processName, windowName, id);

                AtualizaPosicao(windowName, id, tipo);

                //Console.WriteLine("MainWindow of Program has updated {0:x8}", hwnd.ToInt32());
            }
            catch
            {
                //Console.WriteLine("Exception | Seguindo processo");
            }

            EmExecucao = false;

            return;
        }

        /// <summary>
        /// Executa o Avaya Onex
        /// </summary>
        public void AbrirAvayaOnex()
        {
            string exe = ConfigurationManager.AppSettings["CAMINHO_AVAYA"];

            Process process = new Process();
            process.StartInfo.FileName = exe;
            process.StartInfo.Arguments = "-n";
            process.StartInfo.WindowStyle = ProcessWindowStyle.Maximized;
            process.Start();
        }

        /// <summary>
        /// Inicia a execução da BAT de configuração do AVAYA Onex
        /// </summary>
        public void IniciarBat()
        {
            try
            {
                string bat = ConfigurationManager.AppSettings["CAMINHO_BAT"];
                if (File.Exists(bat))
                {
                    Process process = new Process();
                    process.StartInfo.FileName = bat;
                    process.StartInfo.Arguments = "-n";
                    process.Start();
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Falha ao iniciar a BAT de configuração do Avaya \r\n Detalhes do erro: {ex.Message}");
            }
        }
        #endregion

        #region Métodos privados
        /// <summary>
        /// Realiza a verificação das janelas e elementos afim de realizar o bloqueio de um específico
        /// </summary>
        /// <param name="tituloJanela">Nome da janela que contém o elemento</param>
        /// <param name="nomeObjeto">Nome do elemento que será bloqueado</param>
        private static void ExibirBloqueio(string nomeProcesso, string tituloJanela, string nomeObjeto)
        {
            try
            {
                if (FormExibido)
                {
                    //Console.WriteLine("Form em exibição");
                    return;
                }

                //Console.WriteLine($"Thread: {tituloJanela} - {nomeObjeto}");
                //Localiza a janela que contém o elemento
                //var windows = Desktop.Instance.Windows().Where(x => x.Name.Equals(tituloJanela));

                //List<Process> processes = Process.GetProcesses().Where(p => p.ProcessName == nomeProcesso).ToList();
                //int count = 0;
                //while (processes.Count < 1 && count < 10)
                //{
                //    processes = Process.GetProcesses().Where(p => p.ProcessName == nomeProcesso).ToList();
                //    count++;
                //}
                //var application = Application.Attach(processes[0].Id);

                var application = Application.Attach(processId);
                Window window = application.GetWindow(tituloJanela, InitializeOption.NoCache);

                tipo = 0;

                //if (windows.Any())
                if (window != null)
                {
                    //Console.WriteLine($"Thread: {tituloJanela} - {nomeObjeto} | Achou janela");
                    //window = windows.FirstOrDefault();

                    //Localiza o elemento que será bloqueado
                    objeto = window.Items.Find(x => !string.IsNullOrEmpty(x.Id) && x.Id.ToUpper().Equals(nomeObjeto.ToUpper()));

                    if (objeto != null)
                    {
                        //Console.WriteLine($"Thread: {tituloJanela} - {nomeObjeto} | Achou elemento");
                        tipo = 1;
                    }
                    else
                    {
                        if (objeto == null)
                        {
                            objeto = window.Items.Find(x => !string.IsNullOrEmpty(x.Name) && x.Name.ToUpper().Equals(nomeObjeto.ToUpper()));

                            if (objeto != null)
                            {
                                //Console.WriteLine($"Thread: {tituloJanela} - {nomeObjeto} | Achou elemento");
                                tipo = 2;
                            }
                        }
                    }

                    if (objeto == null)
                    {
                        //Console.WriteLine("Objeto NULO: " + nomeObjeto);

                        //Elemento não encontrado
                        if (FormExibido)
                        {
                            //Fecha o form aberto
                            //Console.WriteLine("Fechando o form");
                            Bloqueado.Close();
                            FormExibido = false;
                        }
                        return;
                    }

                    //Console.WriteLine("Objeto VISIVEL: " + nomeObjeto);
                    if (!FormExibido)
                    {
                        //Elemento encontrado e form ainda não aberto
                        //Console.WriteLine("Abrindo novo form");
                        Bloqueado = new frmBloqueado();
                        FormExibido = true;
                        SetForegroundWindow(Bloqueado.Handle);
                    }

                    bool isVisivel;
                    //try
                    //{
                    //    //Captura os pontos clicaveis do elemento para determinar se este está maximizado
                    //    //var checkWindow = window.ClickablePoint;
                    //    //var checkObjeto = objeto.ClickablePoint;
                    //    isVisivel = objeto.Visible;
                    //}
                    //catch
                    //{
                    //    isVisivel = false;
                    //}

                    //if (isVisivel)
                    //{
                    //    //Mantém o form sempre a frente e na posição atual do elemento
                    //    //MantemFormEmPosicao(tituloJanela, nomeObjeto, tipo);

                    //    //Fecha o form
                    //    //Bloqueado.Close();
                    //    //FormExibido = false;

                    //    //if (FormExibido)
                    //    //{
                    //    //    //Fecha o form
                    //    //    Bloqueado.Close();
                    //    //    FormExibido = false;
                    //    //}
                    //}
                    //else
                    //{
                    //    //Fecha o form
                    //    Bloqueado.Close();
                    //    FormExibido = false;
                    //}
                    //Bloqueado.Visible = false;
                }
                else if (FormExibido)
                {
                    //Fecha o form
                    Bloqueado.Close();
                    FormExibido = false;
                }
            }
            catch (Exception ex)
            {
                return;
            }
        }

        /// <summary>
        /// Loop para constante verificação da posição do elemento a ser bloqueado, afim de manter o form sempre no mesmo local que o elemento está
        /// </summary>
        private void MantemFormEmPosicao(string window, string id, int tipo)
        {

            //Console.WriteLine("Monitorando botão");
            try
            {
                bool visivel = true;

                while (visivel)
                {
                    //Console.WriteLine("Object name: " + objeto.Name + " | Tipo: " + tipo );
                    if (tipo == 1 && objeto.Id != null && !objeto.Id.Equals(id))
                    {
                        visivel = false;
                    }
                    else if (tipo == 2 && objeto.Name != null && !objeto.Name.Equals(id))
                    {
                        visivel = false;
                    }
                    else if (Double.IsInfinity(objeto.Location.Y))
                    {
                        visivel = false;
                    }

                    //SetForegroundWindow(Bloqueado.Handle);
                    //if (window.ToUpper().Contains("SOFTPHONE"))
                    //{
                    //    SetWindowPos(Bloqueado.Handle, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);
                    //}
                    //

                    //Atribuindo as propriedades ao form para que sempre fique posicionado sobre o elemento
                    Bloqueado.TopMost = true;
                    Bloqueado.BackColor = ColorTranslator.FromHtml("#F0F");
                    Bloqueado.Opacity = 1;// 0.01;
                    Bloqueado.ShowInTaskbar = false;
                    Bloqueado.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
                    Bloqueado.Top = (int)objeto.Location.Y;
                    Bloqueado.Left = (int)objeto.Location.X;
                    Bloqueado.Width = (int)objeto.Bounds.Width;
                    Bloqueado.Height = (int)objeto.Bounds.Height;

                    Bloqueado.Show();

                    System.Windows.Forms.Application.DoEvents();

                    //if (window.ToUpper().Contains("SOFTPHONE"))
                    //{
                    //    SetWindowPos(Bloqueado.Handle, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);
                    //}


                    System.Threading.Thread.Sleep(500);

                }
            }
            catch
            {
                return;
            }
        }

        private static void AtualizaPosicao(string window, string id, int tipo)
        {
            try
            {
                bool visivel = true;

                //Console.WriteLine("Object name: " + objeto.Name + " | Tipo: " + tipo );
                if (tipo == 1 && objeto.Id != null && !objeto.Id.Equals(id))
                {
                    visivel = false;
                }
                else if (tipo == 2 && objeto.Name != null && !objeto.Name.Equals(id))
                {
                    visivel = false;
                }
                else if (Double.IsInfinity(objeto.Location.Y))
                {
                    visivel = false;
                }

                if (!visivel)
                {
                    Bloqueado.Close();
                    FormExibido = false;

                    return;
                }

                //Atribuindo as propriedades ao form para que sempre fique posicionado sobre o elemento
                Bloqueado.TopMost = true;
                Bloqueado.BackColor = ColorTranslator.FromHtml("#F0F");
                Bloqueado.Opacity = 0.01;
                Bloqueado.ShowInTaskbar = false;
                Bloqueado.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
                Bloqueado.Top = (int)objeto.Location.Y;
                Bloqueado.Left = (int)objeto.Location.X;
                Bloqueado.Width = (int)objeto.Bounds.Width;
                Bloqueado.Height = (int)objeto.Bounds.Height;

                Bloqueado.Show();

                SetWindowPos(Bloqueado.Handle, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);

                System.Windows.Forms.Application.DoEvents();

            }
            catch
            {
                Bloqueado.Close();
                FormExibido = false;
                return;
            }
        }
        #endregion
    }
}
