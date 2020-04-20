using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace ЦМПОЛ_passports
{
    static class Program
    {
        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form2());
        }       
    }

    static class Dostup
    {
        public static string Access { get; set; }
        public static string Login { get; set; }
    }
}
