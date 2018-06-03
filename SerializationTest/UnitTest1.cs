using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using F;

namespace SerializationTest
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void Test()
        {
            var dto = new FootballDto
            {
                Filled = DateTime.Now,
                FullName = "Малых Александр Александрович",
                Citizenship = "Русский",
                Age = 19,
                Height = 178,
                Weight = 60,
                AgeStartCareer = 15,
                ExperienceInFootball = 3,
                Position = Position.Goalkeeper,
                WorkingLeg = WorkingLeg.Right,
                CountTraums = 0,
                TimeTraums = 0,
                TraumаNow = "Нет",
                Traums = " Не было травм",
                Strength1 = Strengths.Reaction,
                WeakSides1 = WeakSides.HeadGame,
                Strength2 = Strengths.Jump,
                WeakSides2 = WeakSides.Force,
                Strength3 = Strengths.Pass,
                WeakSides3 = WeakSides.Speed,
            };
            var tempFileName = Path.GetTempFileName();
            try
            {
                Helper.WriteToFile(tempFileName, dto);
                var readDto = Helper.LoadFromFile(tempFileName);
                Assert.AreEqual(dto.Filled, readDto.Filled);
            }
            finally
            {
                File.Delete(tempFileName);
            }
        }
    }
}
