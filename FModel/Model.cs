using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace F
{
    /// <summary>
    /// информация об футболисте 
    /// </summary>
    public class FootballDto
    {
        /// <summary>
        /// ФИО футболиста
        /// <summary>
        public string FullName { get; set; }
        /// <summary>
        /// Дата заполнения 
        /// <summary> 
        public DateTime Filled { get; set; } 
        /// <summary>
        /// Гражданство футболиста
        /// <summary>
        public string Citizenship { get; set; }
        /// <summary>
        /// Возраст футболиста  
        /// <summary>
        public int Age { get; set; }
        /// <summary>
        /// Рост футболиста
        /// <summary>
        public int Height { get; set; }
        /// <summary>
        /// Вес футболиста
        /// <summary>
        public int Weight { get; set; }
        /// <summary>
        /// Возраст футболиста, когда он начал карьеру
        /// <summary>
        public int AgeStartCareer { get; set; }
        /// <summary>
        /// Опыт игры в футбол у футболиста (полных лет)
        /// <summary>
        public int ExperienceInFootball { get; set; }
        /// <summary>
        /// Позиция футболиста 
        /// <summary>
        public Position Position { get; set; }
        /// <summary>
        /// Рабочая нога футболиста
        /// <summary>
        public WorkingLeg WorkingLeg { get; set; }
        /// <summary>
        /// Количество травм у футболиста за всю карьеру
        /// <summary>
        public int CountTraums { get; set; }
        /// <summary>
        /// Количество матчей, пропущенных футолистом из-за травм за всю карьеру
        /// <summary>
        public int TimeTraums { get; set; }
        /// <summary>
        /// Есть ли у футболиста трамва прямо сейчас
        /// <summary>
        public string TraumаNow { get; set; }
        /// <summary>
        /// Прошлые травмы футболиста
        /// <summary>
        public string Traums { get; set; }
        /// <summary>
        /// Сильные стороны футболиста
        /// <summary>
        public  Strengths Strength1 { get; set; }
        /// <summary>
        /// Слабые стороны футболиста 
        /// <summary>
        public WeakSides WeakSides1 { get; set; }
        /// <summary>
        /// Сильные стороны футболиста
        /// <summary>
        public Strengths Strength2 { get; set; }
        /// <summary>
        /// Слабые стороны футболиста 
        /// <summary>
        public WeakSides WeakSides2 { get; set; }
        /// <summary>
        /// Сильные стороны футболиста
        /// <summary>
        public Strengths Strength3 { get; set; }
        /// <summary>
        /// Слабые стороны футболиста 
        /// <summary>
        public WeakSides WeakSides3 { get; set; }
    }

    public enum TraumaTypes
    {
        Ankle,
        Head,
        Arm,
        Body,
        Knee,
        Foot,
        Groin,
        Femur,
        CalfMuscles,
        CruciformLigaments,
        None,
    }

    public enum WorkingLeg
    {
        Left,
        Right,
        BothLeg,
    }

    public enum WeakSides
    {
        Concentration,
        Speed,
        Shoot,
        Endurance,
        Force,
        HeadGame,
        Dribbling,
        Reaction,
        Jump,
    }

    public enum Strengths
    {
        Shoot,
        Pass,
        Dribbling,
        MasterOfStandards,
        HeadGame,
        Speed,
        Physics,
        Reaction,
        Jump,
    }

    public enum Position
    {
        Goalkeeper,
        Defender,
        Midfielder,
        Forward,
    }
}