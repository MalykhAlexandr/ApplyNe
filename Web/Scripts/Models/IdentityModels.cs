using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity;
using System.Security.Claims;
using System.Threading.Tasks;
using F.Web.Migration;
using Microsoft.AspNet.Identity;
using Microsoft.AspNet.Identity.EntityFramework;

namespace F.Web.Models
{
    // Чтобы добавить данные профиля для пользователя, можно добавить дополнительные свойства в класс ApplicationUser. Дополнительные сведения см. по адресу: http://go.microsoft.com/fwlink/?LinkID=317594.
    public class ApplicationUser : IdentityUser
    {
        public async Task<ClaimsIdentity> GenerateUserIdentityAsync(UserManager<ApplicationUser> manager)
        {
            // Обратите внимание, что authenticationType должен совпадать с типом, определенным в CookieAuthenticationOptions.AuthenticationType
            var userIdentity = await manager.CreateIdentityAsync(this, DefaultAuthenticationTypes.ApplicationCookie);
            // Здесь добавьте утверждения пользователя
            return userIdentity;
        }
    }

    public class ApplicationDbContext : IdentityDbContext<ApplicationUser>
    {
        public ApplicationDbContext()
            : base("DefaultConnection", throwIfV1Schema: false)
        {
            Database.SetInitializer(new MigrateDatabaseToLatestVersion<ApplicationDbContext, Configuration>());
        }

        public static ApplicationDbContext Create()
        {
            return new ApplicationDbContext();
        }

        public DbSet<DbApplyRequest> ApplyRequests { get; set; }
    }
    /// <summary>
    /// Информация о футболисте
    /// </summary>
    public class DbApplyRequest
    {
        [Key, DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }
        /// <summary>
        /// Дата заполнения
        /// </summary>
        public DateTime Filled { get; set; }
        /// <summary>
        /// ФИО футболиста
        /// </summary>
        public string FullName { get; set; }
        /// <summary>
        /// Гражданство футболиста
        /// <summary>
        public string Citizenship { get; set; }
        /// <summary>
        /// Возраст
        /// </summary>
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
        /// Позиция футболиста 
        /// <summary>
        public Position Position { get; set; }
        /// <summary>
        /// Рабочая нога футболиста
        /// <summary>
        public WorkingLeg WorkingLeg { get; set; }
        /// <summary>
        /// Опыт игры в футбол у футболиста (полных лет)
        /// <summary>
        public int AgeStartCareer { get; set; }
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

        public Strengths Strength1 { get; set; }
   
        public WeakSides WeakSides1 { get; set; }
       
        public Strengths Strength2 { get; set; }
        
        public WeakSides WeakSides2 { get; set; }
       
        public Strengths Strength3 { get; set; }
        
        public WeakSides WeakSides3 { get; set; }
    }

    public enum WorkingLeg
    {
        Left,
        Right,
        BothLeg,
    }

    public enum Position
    {
        Goalkeeper,
        Defender,
        Midfielder,
        Forward,
    }
}