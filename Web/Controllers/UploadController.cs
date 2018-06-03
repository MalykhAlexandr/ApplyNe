using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using F;
using F.Web.Models;

namespace Web.Controllers
{
    public class UploadController : Controller
    {
        // GET: Upload
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Print(HttpPostedFileBase file)
        {
            if (file != null && file.ContentLength > 0)
            {
                var dto = Helper.LoadFromStream(file.InputStream);

                using (var db = new ApplicationDbContext())
                {
                    var row = new DbApplyRequest
                    {
                        FullName = dto.FullName,
                        Age = dto.Age,
                        Citizenship = dto.Citizenship,
                        Weight = dto.Weight,
                        Height = dto.Height,
                        Position = (F.Web.Models.Position)dto.Position,
                        WorkingLeg = (F.Web.Models.WorkingLeg)dto.WorkingLeg,
                        Filled = dto.Filled,
                        AgeStartCareer = dto.AgeStartCareer,
                        CountTraums = dto.CountTraums,
                        TimeTraums = dto.TimeTraums,
                        TraumаNow = dto.TraumаNow,
                        Traums = dto.Traums,
                        Strength1 = dto.Strength1,
                        Strength2 = dto.Strength2,
                        Strength3 = dto.Strength3,
                        WeakSides1 = dto.WeakSides1,
                        WeakSides2 = dto.WeakSides2,
                        WeakSides3 = dto.WeakSides3,
                    };

                    /*row.Trauma = new Collection <DbTrauma>();

                    foreach (var wp in dto.Traums)
                    {
                        row.Trauma.Add(new DbTrauma
                        {
                            TraumаNow = wp.,
                            CountTraums = wp.
                            //Type = (Models.WayPointType)(int)wpDto.Type
                        });*/
                    db.ApplyRequests.Add(row);
                    db.SaveChanges();
                }
                return View(dto);
            }
            return RedirectToAction("Index");
        }
    }
}