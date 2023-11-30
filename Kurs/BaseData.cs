using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kurs
{
    public class BaseData
    {
        public static int id = 0;
        public static int secretid = 0;
        public static void ClearData() // Это метод для очистки данных
        {
            id = 0;
        }

        public static void ClearSecretData() // Это метод для очистки данных
        {
            id = 0;
        }

        public static int ValueExercise;

        //datagridValue

        public static string dayTraining;
        public static string monthTraining;
        public static string yearTraining;


        public static string monthd;
        public static string yeard;

        //Добавление тренировки в план

        public static string idd;
        public static string plan_name;
        public static string DatePlan;
        public static string Trainingid;
        public static string Idtraining;


    }
   

}
