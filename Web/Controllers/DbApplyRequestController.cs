using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.IO;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Hosting;
using System.Web.Mvc;
using F.Web.Models;
using OfficeOpenXml;

namespace F.Web.Controllers
{
    public class DbApplyRequestController : Controller
    {
        private ApplicationDbContext db = new ApplicationDbContext();

        // GET: DbApplyRequests
        public ActionResult Index()
        {
            return View(db.ApplyRequests.ToList());
        }

        /*private new ActionResult View(object p)
        {
            throw new NotImplementedException();
        }*/

        // GET: DbApplyRequests/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            DbApplyRequest dbApplyRequest = db.ApplyRequests.Find(id);
            if (dbApplyRequest == null)
            {
                return HttpNotFound();
            }
            return View(dbApplyRequest);
        }

        // GET: DbApplyRequests/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: DbApplyRequests/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "Id, Filled, FullName, Age, Citizenship, Position, WorkingLeg, Height, Weight, AgeStartCareer, CountTraums, TimeTraums, TraumаNow, Traums, Strength1, Strength2, Strength3,WeakSides1, WeakSides2, WeakSides3")] DbApplyRequest dbApplyRequest)
        {
            if (ModelState.IsValid)
            {
                db.ApplyRequests.Add(dbApplyRequest);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(dbApplyRequest);
        }

        // GET: DbApplyRequests/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            DbApplyRequest dbApplyRequest = db.ApplyRequests.Find(id);
            if (dbApplyRequest == null)
            {
                return HttpNotFound();
            }
            return View(dbApplyRequest);
        }

        //POST: DbRideRequests/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            DbApplyRequest dbApplyRequest = db.ApplyRequests.Find(id);
            db.ApplyRequests.Remove(dbApplyRequest);
            db.SaveChanges();
            return RedirectToAction("Index");
        }

        // GET: DbRideRequests/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            DbApplyRequest dbApplyRequest = db.ApplyRequests.Find(id);
            if (dbApplyRequest == null)
            {
                return HttpNotFound();
            }
            return View(dbApplyRequest);
        }

        // POST: DbRideRequests/Edit/5
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "Id, Filled, FullName, Age, Citizenship, Position, WorkingLeg, Height, Weight, AgeStartCareer, CountTraums, TimeTraums, TraumаNow, Traums, Strength1, Strength2, Strength3,WeakSides1, WeakSides2, WeakSides3")] DbApplyRequest dbApplyRequest)
        {
            if (ModelState.IsValid)
            {
                db.Entry(dbApplyRequest).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(dbApplyRequest);
        }


        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }

        public ActionResult Download(int? id)
        {
            var ctx = new ApplicationDbContext();
            var f = ctx.ApplyRequests.Find(id);

            ExcelPackage pkg;
            using (var stream = System.IO.File.OpenRead(HostingEnvironment.ApplicationPhysicalPath + "template.xlsx"))
            {
                pkg = new ExcelPackage(stream);
                stream.Dispose();
            }

            var worksheet = pkg.Workbook.Worksheets[1];
            worksheet.Name = "Информация о заявке";

            worksheet.Cells[2, 2].Value = "Имя футболиста";
            worksheet.Cells[2, 3].Value = f.FullName;
            worksheet.Cells[3, 2].Value = "Время запонения";
            worksheet.Cells[3, 3].Value = f.Filled.ToString();
            worksheet.Cells[4, 2].Value = "Гражданство";
            worksheet.Cells[4, 3].Value = f.Citizenship;
            worksheet.Cells[5, 2].Value = "Возраст футболиста";
            worksheet.Cells[5, 3].Value = f.Age;
            worksheet.Cells[6, 2].Value = "Возраст начала карьеры";
            worksheet.Cells[6, 3].Value = f.AgeStartCareer;
            worksheet.Cells[7, 2].Value = "Позиция";
            worksheet.Cells[7, 3].Value = f.Position;
            worksheet.Cells[8, 2].Value = "Рабочая нога";
            worksheet.Cells[8, 3].Value = f.WorkingLeg;
            worksheet.Cells[9, 2].Value = "Рост футболиста";
            worksheet.Cells[9, 3].Value = f.Height;
            worksheet.Cells[10, 2].Value = "Вес футболиста";
            worksheet.Cells[10, 3].Value = f.Weight;

            worksheet.Cells[2, 5].Value = "Сильные стороны футболиста:";
            worksheet.Cells[3, 5].Value = f.Strength1;
            worksheet.Cells[4, 5].Value = f.Strength2;
            worksheet.Cells[5, 5].Value = f.Strength3;

            worksheet.Cells[2, 6].Value = "Слабые стороны футболиста:";
            worksheet.Cells[3, 6].Value = f.WeakSides1;
            worksheet.Cells[4, 6].Value = f.WeakSides2;
            worksheet.Cells[5, 6].Value = f.WeakSides3;

            worksheet.Cells[12, 2].Value = "Информация о травмах:";
            worksheet.Cells[13, 2].Value = "Количество травм";
            worksheet.Cells[13, 3].Value = f.CountTraums;
            worksheet.Cells[14, 2].Value = "Количество матчей, пропущенных из-за травм";
            worksheet.Cells[14, 3].Value = f.TimeTraums;
            worksheet.Cells[15, 2].Value = "Есть ли травма сейчас";
            worksheet.Cells[15, 3].Value = f.TraumаNow;

            worksheet.Cells[17, 2].Value = "Травмы:";
            worksheet.Cells[18, 2].Value = f.Traums;

            worksheet.Cells.AutoFitColumns();
            var ms = new MemoryStream();
            pkg.SaveAs(ms);

            return File(ms.ToArray(), "application/ooxml", ((f.FullName ?? "Без Названия") + f.Filled.ToString()).Replace(" ", "") + ".xlsx");
        }
    }

}