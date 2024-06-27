using System;
using System.Configuration;
using System.Runtime.InteropServices;
using System.Threading;

namespace TP_Automation.AvayaOnex.DesabilitaDesligue
{
    class Program
    {

        [DllImport("user32.dll", SetLastError = true)]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("kernel32.dll")]
        static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        const int SW_HIDE = 0;
        const int SW_SHOW = 5;

        static void Main(string[] args)
        {
            var bloqueador = new Bloqueador();
            var handle = GetConsoleWindow();

            // Hide
            ShowWindow(handle, SW_HIDE);

            bool executarAvaya = true;
            if (args.Length > 0 && args[0].Equals("1"))
            {
                executarAvaya = false;
            }
            
            if(executarAvaya)
            {
                //bloqueador.IniciarBat();
                System.Threading.Thread.Sleep(1000);
                bloqueador.AbrirAvayaOnex();
            }

            var bloqueios = ConfigurationManager.AppSettings["BLOQUEIOS"];
            foreach (var bloqueio in bloqueios.Split(';'))
            {
                new Thread(() =>
                    new Bloqueador().BloquearControle(
                            nomeProcesso: bloqueio.Split('|')[0],
                            tituloJanela: bloqueio.Split('|')[1],
                            nomeObjeto: bloqueio.Split('|')[2],
                            tipoMonitoramento: TipoMonitoramento.EVENT_OBJECT_NAME_CHANGED
                )).Start();
            }
            //bloqueador.IniciarBat();
            //System.Threading.Thread.Sleep(1000);
            //bloqueador.AbrirAvayaOnex();
            //bloqueador.BloquearControle(
            //        tituloJanela: "Avaya one-X® Communicator",
            //        nomeObjeto: "Assign"
            //    //nomeObjeto: "ButtonEnd"
            //    );

            //new Thread(() =>
            //new Bloqueador().BloquearControle(
            //        tituloJanela: "LH.SoftPhone",
            //        nomeObjeto: "backoffice"
            //    )).Start();

            //new Thread(() =>
            //new Bloqueador().BloquearControle(
            //        tituloJanela: "LH.SoftPhone",
            //        nomeObjeto: " desligar"
            //    )).Start();

            //new Thread(() =>
            //new Bloqueador().BloquearControle(
            //        tituloJanela: "Avaya one-X® Communicator",
            //        nomeObjeto: "Assign"
            //    )).Start();

            //new Thread(() =>
            //new Bloqueador().BloquearControle(
            //        tituloJanela: "Avaya one-X® Communicator",
            //        nomeObjeto: "ButtonEnd"
            //    )).Start();


        }
    }
}
